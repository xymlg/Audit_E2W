"""按金额列排序：锚点行「排序：」设置的读写、数据行重排、导出重建时按书签沿用。

规则：表头行与合计行不动；首列（或第二列）项目名为「其他」的行不参与排序，
统一放在数据区末尾（合计行之前）；解析不了的金额按最小值排在排序段末尾。
数据区含合并单元格的表保守跳过，避免错位。
"""

from __future__ import annotations

import re
from copy import copy

from openpyxl.styles import Alignment, Font

from .config import FILL_ANCHOR, FONT_NAME, FONT_SIZE, META_SCAN_COLUMNS, NUMBER_RE, SHEET_NAME
from .console import log, warn
from .errors import SyncError
from .valuefmt import display_text, evaluate_formula

SORT_PREFIX = "排序："
DESC = "降序"
ASC = "升序"
OTHER_ITEM = "其他"

_RANGE_OR_REF_RE = re.compile(r"(?<![A-Za-z0-9_$])(\$?[A-Za-z]{1,3}\$?)(\d+)(?::(\$?[A-Za-z]{1,3}\$?)(\d+))?")


def sort_setting_text(column: str, descending: bool = True) -> str:
    return f"{SORT_PREFIX}{column.strip()} {DESC if descending else ASC}"


def parse_sort_setting(text: str) -> tuple[str, bool] | None:
    text = (text or "").strip()
    if not text.startswith(SORT_PREFIX):
        return None
    body = text[len(SORT_PREFIX):].strip()
    descending = True
    if body.endswith(DESC):
        body = body[: -len(DESC)].strip()
    elif body.endswith(ASC):
        descending = False
        body = body[: -len(ASC)].strip()
    return (body, descending) if body else None


def _anchor_style(cell) -> None:
    cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True, color="1F4E79")
    cell.fill = FILL_ANCHOR
    cell.alignment = Alignment(horizontal="left", vertical="center")


def write_block_sort(ws, block: dict, column: str, descending: bool = True) -> None:
    """把排序设置写进锚点行第一个空闲元数据列（列数格之后），已有设置则原位更新。"""
    row = block["anchor_row"]
    col0 = block.get("col0", 2)
    target = None
    for col in range(col0 + 4, META_SCAN_COLUMNS + 1):
        text = "" if ws.cell(row, col).value is None else str(ws.cell(row, col).value).strip()
        if text.startswith(SORT_PREFIX):
            target = col
            break
        if target is None and text == "":
            target = col
    if target is None:
        raise SyncError(f"{block['code']}：锚点行元数据列已满，无法写入排序设置")
    cell = ws.cell(row, target, sort_setting_text(column, descending))
    _anchor_style(cell)


def clear_block_sort(ws, block: dict) -> bool:
    row = block["anchor_row"]
    col0 = block.get("col0", 2)
    cleared = False
    for col in range(col0 + 4, META_SCAN_COLUMNS + 1):
        cell = ws.cell(row, col)
        text = "" if cell.value is None else str(cell.value).strip()
        if text.startswith(SORT_PREFIX):
            cell.value = None
            cleared = True
    return cleared


def block_headers(ws, block: dict) -> list[str]:
    """表块首行（表头）各列显示文本，供界面选择金额列。"""
    if not block["grid_rows"]:
        return []
    head = block["grid_rows"][0]
    col0 = block.get("col0", 2)
    return [
        display_text(ws.cell(head, col0 + c).value, ws.cell(head, col0 + c).number_format).strip()
        for c in range(block["ncols"])
    ]


def _row_text(ws, row: int, col: int) -> str:
    cell = ws.cell(row, col)
    return display_text(cell.value, cell.number_format).strip()


def _amount(ws, row: int, col: int, cached) -> float | None:
    cell = ws.cell(row, col)
    raw = cell.value
    if isinstance(raw, str) and raw.startswith("="):
        raw = cached.get((row, col))
        if raw is None:
            raw = evaluate_formula(ws, row, col, cached)
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return float(raw)
    if isinstance(raw, str) and NUMBER_RE.match(raw.strip()):
        try:
            return float(raw.strip().replace(",", ""))
        except ValueError:
            return None
    return None


def _is_other(ws, row: int, ncols: int, col0: int = 2) -> bool:
    if _row_text(ws, row, col0) == OTHER_ITEM:
        return True
    return ncols > 1 and _row_text(ws, row, col0 + 1) == OTHER_ITEM


def _remap_formula(text: str, rowmap: dict[int, int]) -> str:
    """按旧行号→新行号映射改写公式引用；区间取映射后行号的最小/最大（保持求和范围不变）。"""

    def fix_part(part: str) -> str:
        if part.startswith('"'):
            return part

        def repl(m: re.Match) -> str:
            if part[m.end():m.end() + 1] == "(":  # LOG10( 这类函数名，不是单元格引用
                return m.group(0)
            if m.group(3) is not None:
                r1, r2 = int(m.group(2)), int(m.group(4))
                rows = [rowmap.get(r, r) for r in range(min(r1, r2), max(r1, r2) + 1)]
                return f"{m.group(1)}{min(rows)}:{m.group(3)}{max(rows)}"
            row = int(m.group(2))
            return f"{m.group(1)}{rowmap.get(row, row)}"

        return _RANGE_OR_REF_RE.sub(repl, part)

    return "".join(fix_part(p) for p in re.split(r'("[^"]*")', text))


def apply_block_sort(ws, block: dict, cached: dict, dry_run: bool = False) -> dict | None:
    """按表块保存的排序设置重排数据行（原地修改 ws）。无设置/无需改动时返回 None。"""
    from .excel_blocks import block_merges

    setting = parse_sort_setting(block.get("meta", {}).get("sort", ""))
    if setting is None:
        return None
    column, descending = setting
    code = block["code"]
    rows = block["grid_rows"]
    ncols = block["ncols"]
    if len(rows) < 3:
        return None
    headers = block_headers(ws, block)
    matches = [c for c, h in enumerate(headers) if h == column]
    if not matches:
        warn(f"[排序] {code}：找不到金额列「{column}」（表头可能已改），本次跳过排序")
        return None
    sort_col = block.get("col0", 2) + matches[0]
    col0 = block.get("col0", 2)

    data = rows[1:]
    tail_start = next((i for i, r in enumerate(data) if _row_text(ws, r, col0).startswith("合计")), len(data))
    sortable, tail = data[:tail_start], data[tail_start:]
    if len(sortable) < 2:
        return None
    sort_idx = {r: i + 1 for i, r in enumerate(sortable)}  # 相对块内网格的行号（表头为 0）
    for (r, _c, rs, cs) in block_merges(ws, block):
        if (rs > 1 or cs > 1) and any(r <= idx < r + rs for idx in sort_idx.values()):
            warn(f"[排序] {code}：数据区含合并单元格，为避免错位本次跳过排序")
            return None

    normal = [r for r in sortable if not _is_other(ws, r, ncols, col0)]
    others = [r for r in sortable if _is_other(ws, r, ncols, col0)]
    parsed = [r for r in normal if _amount(ws, r, sort_col, cached) is not None]
    unparsed_rows = [r for r in normal if r not in parsed]
    unparsed = [_row_text(ws, r, col0) or f"第{sort_idx[r] + 1}行" for r in unparsed_rows]
    ordered = sorted(parsed, key=lambda r: _amount(ws, r, sort_col, cached), reverse=descending)
    new_data = ordered + unparsed_rows + others + tail
    if new_data == data:
        log(f"[排序] {code}：已是「{column} {DESC if descending else ASC}」顺序，无需改动")
        return None
    moved = sum(1 for old, new in zip(data, new_data) if old != new)
    note = f"（其他 {len(others)} 行置底" + (f"，{len(unparsed)} 行金额无法解析按最小值处理：{'、'.join(unparsed[:5])}" if unparsed else "") + "）"
    if dry_run:
        log(f"[排序·预演] {code}：按「{column}」{DESC if descending else ASC}重排，{moved} 行将移动{note}，合计行不动")
        return {"moved": moved, "others": len(others), "unparsed": unparsed}

    snapshot = {}
    for r in data:
        for c in range(2, col0 + ncols):
            cell = ws.cell(r, c)
            snapshot[(r, c)] = (cell.value, copy(cell._style) if cell.has_style else None,
                                copy(cell.hyperlink) if cell.hyperlink is not None else None)
    heights = {r: ws.row_dimensions[r].height for r in data if r in ws.row_dimensions and ws.row_dimensions[r].height is not None}
    rowmap = {old: new for old, new in zip(data, new_data)}
    for target, source in zip(data, new_data):
        for c in range(2, col0 + ncols):
            value, style, link = snapshot[(source, c)]
            cell = ws.cell(target, c)
            cell.value = value
            if style is not None:
                cell._style = style
            if link is not None:
                link.ref = cell.coordinate
                cell.hyperlink = link
            else:
                cell.hyperlink = None
        if source in heights:
            ws.row_dimensions[target].height = heights[source]
            ws.row_dimensions[target].customHeight = True
    for r in data:
        for c in range(2, col0 + ncols):
            cell = ws.cell(r, c)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = _remap_formula(cell.value, rowmap)
    cached_snapshot = {key: value for key, value in cached.items() if key[0] in rowmap}
    for key in list(cached_snapshot):
        del cached[key]
    for (r, c), value in cached_snapshot.items():
        cached[(rowmap[r], c)] = value
    log(f"[排序] {code}：已按「{column}」{DESC if descending else ASC}重排 {moved} 行{note}，合计行保持末行")
    return {"moved": moved, "others": len(others), "unparsed": unparsed}


def carry_over_sort_settings(wb_new, ws_old, blocks_old: dict) -> list[str]:
    """导出重建工作簿后，把旧表块的排序设置按书签名沿用到新表块。"""
    from .excel_blocks import parse_blocks

    ws_new = wb_new[SHEET_NAME]
    _, blocks_new = parse_blocks(ws_new)
    by_bookmark = {b["bookmark"]: b for b in blocks_new.values() if b.get("bookmark")}
    carried = []
    for block in blocks_old.values():
        setting = parse_sort_setting(block.get("meta", {}).get("sort", ""))
        bookmark = block.get("bookmark", "")
        if setting is None or not bookmark or bookmark not in by_bookmark:
            continue
        write_block_sort(ws_new, by_bookmark[bookmark], setting[0], setting[1])
        carried.append(by_bookmark[bookmark]["code"])
    return carried