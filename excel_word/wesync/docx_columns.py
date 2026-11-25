"""列结构同步：用表头文本打分定位列、在 Word 中插入/删除列并同步合并区，空白新增列按表忽略。"""

import copy
import itertools
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from .config import MAX_COLUMN_SHIFT, SKIP_COLUMN
from .console import log
from .docx_edit import _remove_text_nodes
from .docx_grid import build_cellmap, cell_text, header_rows, parse_tc_props
from .errors import SyncError
from .valuefmt import display_text


def _set_grid_span(tc, span: int) -> None:
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    node = tcPr.find(qn("w:gridSpan"))
    if span <= 1:
        if node is not None:
            tcPr.remove(node)
        return
    if node is None:
        node = OxmlElement("w:gridSpan")
        tcPr.append(node)
    node.set(qn("w:val"), str(span))


def _clone_empty_tc(tc):
    """克隆单元格格式（宽度/边框/对齐）并清空内容，作为新列的单元格。"""
    new_tc = copy.deepcopy(tc)
    tcPr = new_tc.find(qn("w:tcPr"))
    if tcPr is not None:
        for tag in ("w:vMerge", "w:gridSpan"):
            node = tcPr.find(qn(tag))
            if node is not None:
                tcPr.remove(node)
    for node in list(new_tc.iter(qn("w:bookmarkStart"))) + list(new_tc.iter(qn("w:bookmarkEnd"))):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
    paragraphs = new_tc.findall(qn("w:p"))
    for extra in paragraphs[1:]:
        new_tc.remove(extra)
    if paragraphs:
        _remove_text_nodes(paragraphs[0])
    else:
        new_tc.append(OxmlElement("w:p"))
    return new_tc


def _insert_tc_at(tr, pos: int) -> None:
    col = 0
    prev = None
    for tc in tr.findall(qn("w:tc")):
        span, _ = parse_tc_props(tc)
        if col < pos < col + span:
            _set_grid_span(tc, span + 1)  # 插入点落在横向合并区内部：扩展合并
            return
        if col == pos:
            tc.addprevious(_clone_empty_tc(tc))
            return
        prev = tc
        col += span
    if prev is None:
        raise SyncError("表格行内没有可复制的单元格，无法插入列")
    prev.addnext(_clone_empty_tc(prev))


def insert_word_column(tbl, pos: int) -> None:
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        cols = grid.findall(qn("w:gridCol"))
        if cols:
            source = cols[pos - 1] if 0 < pos <= len(cols) else cols[0]
            new_col = copy.deepcopy(source)
            if pos < len(cols):
                cols[pos].addprevious(new_col)
            else:
                cols[-1].addnext(new_col)
    for tr in tbl.findall(qn("w:tr")):
        _insert_tc_at(tr, pos)


def delete_word_column(tbl, pos: int) -> None:
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        cols = grid.findall(qn("w:gridCol"))
        if len(cols) <= 1:
            raise SyncError("表格只剩一列，无法再删除")
        if pos < len(cols):
            grid.remove(cols[pos])
    for tr in tbl.findall(qn("w:tr")):
        col = 0
        for tc in tr.findall(qn("w:tc")):
            span, _ = parse_tc_props(tc)
            if col <= pos < col + span:
                if span > 1:
                    _set_grid_span(tc, span - 1)
                else:
                    tr.remove(tc)
                break
            col += span


def _excel_cell_display(ws, row: int, col: int, cached=None) -> str:
    cell = ws.cell(row, col)
    value = cell.value
    if isinstance(value, str) and value.startswith("=") and cached is not None:
        value = cached.get((row, col))
    return display_text(value, cell.number_format).strip()


def _word_column_signatures(trs, cellmap, ncols: int) -> list[tuple[str, ...]]:
    header_idx = sorted(header_rows(trs)) or list(range(min(2, len(trs))))
    signatures = []
    for col in range(ncols):
        texts = []
        for row in header_idx:
            if row >= len(cellmap):
                texts.append("")
                continue
            pos = cellmap[row][col]
            texts.append(cell_text(pos.tc).strip() if (pos is not None and pos.owner) else "")
        signatures.append(tuple(texts))
    return signatures


def _excel_column_signatures(ws, grid_rows: list[int], ncols: int, header_count: int, cached=None, col0: int = 2) -> list[tuple[str, ...]]:
    rows = grid_rows[: max(1, header_count)]
    signatures = []
    for col in range(ncols):
        signatures.append(tuple(_excel_cell_display(ws, row, col0 + col, cached) for row in rows))
    return signatures


def _score_mapping(word_sigs, excel_sigs, mapping: list[int | None]) -> tuple[float, int, int]:
    """返回 (匹配得分, 参与比较的列数, 文本真正匹配上的单元格数)。"""
    matched = 0.0
    compared = 0
    non_empty = 0
    for excel_col, word_col in enumerate(mapping):
        if word_col is None or word_col < 0 or word_col >= len(word_sigs):
            continue
        compared += 1
        for word_text, excel_text in zip(word_sigs[word_col], excel_sigs[excel_col]):
            if word_text and excel_text:
                if word_text == excel_text:
                    matched += 1.0
                    non_empty += 1
            elif not word_text and not excel_text:
                matched += 0.5
    score = matched / compared if compared else 0.0
    return score, compared, non_empty


def _column_is_empty(ws, col: int, grid_rows: list[int], col0: int = 2) -> bool:
    return all(_excel_cell_display(ws, row, col0 + col) == "" for row in grid_rows)


def plan_column_alignment(word_sigs, excel_sigs, word_ncols: int, excel_ncols: int, ws, grid_rows, code: str, col0: int = 2):
    """用表头文本逐列比对，返回 Excel 列映射。

    映射值含义：整数 = 对应的旧 Word 列；None = 需要新增的 Word 列；
    SKIP_COLUMN = 该 Excel 列在本表中整列空白，本次忽略（不写入 Word）。
    """
    if excel_ncols == word_ncols:
        return list(range(word_ncols))
    shift = abs(excel_ncols - word_ncols)
    if shift > MAX_COLUMN_SHIFT:
        raise SyncError(
            f"{code}：Excel 与 Word 相差 {shift} 列，超出自动同步范围（最多 {MAX_COLUMN_SHIFT} 列）。"
            f"请重新执行 export，或手工把列结构改回一致后重试"
        )
    best = None
    if excel_ncols > word_ncols:
        for combo in itertools.combinations(range(excel_ncols), shift):
            inserted = set(combo)
            mapping: list[int | None] = []
            word_col = 0
            for excel_col in range(excel_ncols):
                if excel_col in inserted:
                    # 有内容的新增列同步到 Word；整列空白的新增列在本表中忽略
                    mapping.append(SKIP_COLUMN if _column_is_empty(ws, excel_col, grid_rows, col0) else None)
                else:
                    mapping.append(word_col)
                    word_col += 1
            score, compared, non_empty = _score_mapping(word_sigs, excel_sigs, mapping)
            skips = sum(1 for value in mapping if value == SKIP_COLUMN)
            inserts = sum(1 for value in mapping if value is None)
            key = (round(score, 6), non_empty, skips, -inserts, -min(inserted))
            if best is None or key > best[0]:
                best = (key, mapping, min(inserted), compared, non_empty)
    else:
        for combo in itertools.combinations(range(word_ncols), shift):
            deleted = set(combo)
            remaining = [col for col in range(word_ncols) if col not in deleted]
            mapping = list(remaining)
            score, compared, non_empty = _score_mapping(word_sigs, excel_sigs, mapping)
            key = (round(score, 6), non_empty, 0, -min(deleted))
            if best is None or key > best[0]:
                best = (key, mapping, min(deleted), compared, non_empty)
    if best is None:
        raise SyncError(f"{code}：无法判断列位置，请重新执行 export 或手工调整列结构")
    score, compared, non_empty = best[0][0], best[3], best[4]
    if compared < 2 or non_empty < 2 or score < 0.8:
        raise SyncError(
            f"{code}：Word 为 {word_ncols} 列、Excel 为 {excel_ncols} 列，表头文本匹配不足"
            f"（得分 {score:.2f}、可比列 {compared}、真实匹配 {non_empty}），"
            f"为避免写错列已跳过该表。请核对书签是否锚定正确，或重新执行 export"
        )
    return best[1]


def sync_columns(tbl, ws, block: dict, code: str, cached=None) -> tuple[list[int], int, int]:
    """按 Excel 列结构同步 Word 表格的列，返回 (有效 Excel 列顺序, 新增列数, 删除列数)。"""
    trs, cellmap, word_ncols = build_cellmap(tbl)
    excel_ncols = block["ncols"]
    if word_ncols == excel_ncols:
        return list(range(excel_ncols)), 0, 0
    word_sigs = _word_column_signatures(trs, cellmap, word_ncols)
    header_count = len(sorted(header_rows(trs))) or min(2, len(trs))
    excel_sigs = _excel_column_signatures(ws, block["grid_rows"], excel_ncols, header_count, cached, block.get("col0", 2))
    mapping = plan_column_alignment(word_sigs, excel_sigs, word_ncols, excel_ncols, ws, block["grid_rows"], code, block.get("col0", 2))

    inserted = 0
    deleted = 0
    if excel_ncols > word_ncols:
        positions = [col for col, word_col in enumerate(mapping) if word_col is None]
        for pos in sorted(positions, reverse=True):
            insert_word_column(tbl, pos)
            inserted += 1
        skipped = [col for col, word_col in enumerate(mapping) if word_col == SKIP_COLUMN]
        if inserted:
            log(f"[结构] {code}：Word 新增 {inserted} 列（插入于第 {min(positions) + 1} 列起）")
        if skipped:
            log(
                f"[结构] {code}：Excel 第 {'、'.join(str(col + 1) for col in skipped)} 列在本表整列空白，"
                f"已按空白列忽略（填入表头或数据后即可同步到 Word）"
            )
    else:
        used = {col for col in mapping if col is not None and col >= 0}
        positions = [col for col in range(word_ncols) if col not in used]
        for pos in sorted(positions, reverse=True):
            delete_word_column(tbl, pos)
            deleted += 1
        log(f"[结构] {code}：Word 删除 {deleted} 列（第 {min(positions) + 1} 列起）")

    include = [col for col, word_col in enumerate(mapping) if word_col != SKIP_COLUMN]
    _, _, new_ncols = build_cellmap(tbl)
    if new_ncols != len(include):
        raise SyncError(
            f"{code}：列结构同步后 Word 为 {new_ncols} 列、Excel 有效列为 {len(include)} 列，未对齐"
        )
    return include, inserted, deleted
