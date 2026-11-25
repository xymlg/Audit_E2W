"""内容签名与书签对齐：把工作簿表块与 Word 表格按内容精确配对，并把书签整表锚定。

签名 = (列数, 行数, 首行文本, 首列文本)，两侧文本都先做空白归一化。
只接受**完全一致**的签名，不做模糊猜测；签名相同的重复表按文档顺序 1:1 消歧。
"""

from __future__ import annotations

from collections import defaultdict, deque

from .console import log
from .config import LEGACY_BOOKMARK_RE
from .docx_bookmarks import (
    bookmark_names,
    insert_table_bookmark,
    new_cpah_name,
    next_bookmark_id,
    remove_bookmark,
    table_index_map_by_bookmark,
)
from .docx_grid import build_cellmap, cell_text
from .valuefmt import display_text


def _norm(text) -> str:
    return " ".join(str(text if text is not None else "").split()).strip()


def table_signature(tbl) -> tuple:
    """Word 表格的内容签名。"""
    trs, cellmap, ncols = build_cellmap(tbl)
    first_row = tuple(
        _norm(cell_text(cellmap[0][c].tc)) if cellmap[0][c] is not None and cellmap[0][c].owner else ""
        for c in range(ncols)
    )
    first_col = tuple(
        _norm(cell_text(row[0].tc)) if row[0] is not None and row[0].owner else "" for row in cellmap
    )
    return (ncols, len(trs), first_row, first_col)


def block_signature(block: dict, ws) -> tuple:
    """工作簿表块的内容签名（数值按 Excel 显示文本比较）。"""
    rows = block["grid_rows"]
    col0 = block.get("col0", 2)
    first_row = tuple(
        _norm(display_text(ws.cell(rows[0], col0 + c).value, ws.cell(rows[0], col0 + c).number_format))
        for c in range(block["ncols"])
    ) if rows else ()
    first_col = tuple(_norm(display_text(ws.cell(r, col0).value, ws.cell(r, col0).number_format)) for r in rows)
    return (block["ncols"], len(rows), first_row, first_col)


def align_blocks_to_tables(doc, ws, blocks: dict) -> tuple[dict[str, int], list[int], list[str]]:
    """按内容签名把表块与 Word 表格 1:1 配对。

    返回 ``(块编号→表序号, 未匹配的表序号, 未匹配的块编号)``。
    """
    if not blocks:
        return {}, list(range(len(doc.tables))), []
    groups: dict[tuple, deque] = defaultdict(deque)
    for code, block in blocks.items():
        groups[block_signature(block, ws)].append(code)

    mapping: dict[str, int] = {}
    unmatched_tables: list[int] = []
    for index, tab in enumerate(doc.tables):
        queue = groups.get(table_signature(tab._tbl))
        if queue:
            mapping[queue.popleft()] = index
        else:
            unmatched_tables.append(index)
    used = set(mapping)
    unmatched_codes = [code for code in blocks if code not in used]
    return mapping, unmatched_tables, unmatched_codes


def plan_table_bookmarks(
    doc,
    ws,
    blocks: dict,
    recorded: dict[str, str],
    *,
    name_unmatched: bool = False,
) -> dict:
    """规划每张表的书签名。

    - 能按内容对上的块：沿用工作簿记录的书签名（没有记录则新建一个 Table_ 名）；
    - ``name_unmatched=True`` 时，给“没对上任何块”的表也新建书签（用于首次导出）；
      否则这些表保持原样，只作为“未对齐”报告出来。
    """
    mapping, unmatched_tables, unmatched_codes = align_blocks_to_tables(doc, ws, blocks)
    fresh = set(bookmark_names(doc.element.body))
    assignments: list[tuple[int, str]] = []
    used: set[str] = set()
    for code, index in sorted(mapping.items(), key=lambda item: item[1]):
        name = (recorded.get(code) or "").strip()
        if name and name in used:
            log(f"[书签] {code} 记录的书签名 {name} 与其它表重复，已改用新的书签名")
            name = ""
        if not name:
            name = new_cpah_name(fresh)
        used.add(name)
        assignments.append((index, name))
    if name_unmatched:
        for index in unmatched_tables:
            assignments.append((index, new_cpah_name(fresh)))
    return {
        "assignments": assignments,
        "mapping": mapping,
        "unmatched_tables": unmatched_tables,
        "unmatched_codes": unmatched_codes,
    }


def apply_table_bookmarks(doc, assignments: list[tuple[int, str]], remove_names: set[str]) -> dict:
    """先删除 ``remove_names`` 里的旧书签，再把书签整表锚定到对应表格。"""
    removed = 0
    for name in sorted(remove_names):
        if remove_bookmark(doc, name):
            removed += 1
    bookmark_id = next_bookmark_id(doc)
    inserted = 0
    for index, name in assignments:
        if index < 0 or index >= len(doc.tables):
            continue
        if insert_table_bookmark(doc.tables[index]._tbl, name, bookmark_id):
            bookmark_id += 1
            inserted += 1
        else:
            log(f"[警告] 第 {index + 1} 张表无法插入书签 {name}（表格为空）")
    return {"removed": removed, "inserted": inserted}


def split_bookmark_updates(
    doc, assignments: list[tuple[int, str]], recorded: dict[str, str]
) -> tuple[list[tuple[int, str]], set[str]]:
    """区分“已正确锚定”与“需要重建”的书签。

    返回 ``(需要重建的 (表序号, 名字), 需要移除的旧书签名)``：
    已经锚定在目标表内部的书签不重建（保证重复运行幂等）；
    需要移除的是本工具自己的旧书签（工作簿记录的名字 + 早期 TBL_*）。
    """
    target_names = [name for _, name in assignments]
    in_table = table_index_map_by_bookmark(
        doc, {name for name in recorded.values() if name} | set(target_names)
    )
    keep = {name for index, name in assignments if in_table.get(name) == index}
    removals = {name for name in recorded.values() if name and name not in keep}
    removals |= {
        name for name in bookmark_names(doc.element.body) if LEGACY_BOOKMARK_RE.match(name)
    }
    updates = [(index, name) for index, name in assignments if name not in keep]
    return updates, removals
