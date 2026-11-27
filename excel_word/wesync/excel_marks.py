"""工作簿里的「链接标注」工作表：高亮标记的导出与回读。

列：标记 | 书签 | 位置 | 原文 | 新内容 | 链接地址 | →定位Word | 来源 | 登记序号
* 导出/刷新时按书签名保留已填的「新内容」「链接地址」；
* Word 中已删除的标记保留成行并在位置列标注「（Word 中已删除）」。
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

from .config import FONT_NAME
from .console import log

SHEET_MARKS = "链接标注"
HEADERS = ["标记", "书签", "位置", "原文", "新内容", "链接地址", "→定位Word", "来源", "登记序号"]
HEADER_FILL = PatternFill("solid", fgColor="FFF2CC")
DELETED_FONT = Font(name=FONT_NAME, size=10, color="9AA3AF")
THIN = Side(style="thin", color="E3E6EA")


def read_previous_marks(xlsx_path: Path) -> list[dict]:
    """读取现有工作簿里「链接标注」工作表的历史内容（用于保留已填写项）。"""
    if not xlsx_path.exists():
        return []
    try:
        workbook = load_workbook(str(xlsx_path))
    except Exception:
        return []
    if SHEET_MARKS not in workbook.sheetnames:
        return []
    return parse_marks_rows(workbook[SHEET_MARKS], resolve_formulas=False)


def parse_marks_rows(ws, cached_ws=None, *, resolve_formulas: bool = True) -> list[dict]:
    """把工作表读成行字典；遇到「标记」列空行就停止。

    「新内容」按单元格在 Excel 里的显示效果转换（与表格回写/占位符同口径）：
    日期按数字格式渲染（如 2025年12月31日），不会出现 2025-12-31 00:00:00；
    金额千分位/小数位照显示。公式优先取 Excel 缓存结果，其次轻量求值。
    """
    from .errors import SyncError
    from .valuefmt import display_cell, formula_result

    rows: list[dict] = []
    seen_bookmarks: dict[str, int] = {}
    if ws.max_row is None:
        return rows
    for row_index in range(2, ws.max_row + 1):
        def cell(column: int, display: bool = False) -> str:
            c = ws.cell(row_index, column)
            raw = c.value
            if resolve_formulas and isinstance(raw, str) and raw.startswith("="):
                raw = formula_result(ws, row_index, column, cached_ws)
                if raw is None:
                    raise SyncError(
                        f"「{SHEET_MARKS}」{c.coordinate} 是公式但没有可用的计算结果。"
                        "请先在 Excel 中打开该工作簿，完成重算并保存后再同步。"
                    )
            if display:
                return display_cell(c, raw).strip()
            return "" if raw is None else str(raw).strip()

        label = cell(1)
        if not label:
            break
        bookmark = cell(2)
        if bookmark:
            if bookmark in seen_bookmarks:
                raise SyncError(
                    f"「{SHEET_MARKS}」书签 {bookmark} 重复出现在第 {seen_bookmarks[bookmark]} 行和第 {row_index} 行，"
                    "无法确定同步目标。请合并为一行后再同步。"
                )
            seen_bookmarks[bookmark] = row_index
        sequence = ws.cell(row_index, 9).value
        try:
            sequence = int(sequence)
        except (TypeError, ValueError):
            sequence = row_index - 1
        rows.append(
            {
                "row": row_index,
                "label": label,
                "bookmark": bookmark,
                "position": cell(3),
                "original": cell(4),
                "new": cell(5, display=True),
                "link": cell(6),
                "source": cell(8),
                "sequence": sequence,
            }
        )
    return rows


def build_marks_sheet(
    wb,
    docx_name: str,
    marks: list,
    previous: list[dict],
    processed: list[dict] | None = None,
) -> dict:
    """在 ``wb`` 中创建/刷新「链接标注」工作表，返回统计。

    三类行：待处理（仍带高亮）→ 已处理（同步过、已去色，书签仍在）→ 已删除（Word 里连书签都没了）。
    行号（M01…）按书签沿用上一版，保证 Excel 里引用稳定。
    """
    previous_by_bookmark = {row["bookmark"]: row for row in previous if row["bookmark"]}
    current = {mark.bookmark for mark in marks if mark.bookmark}
    orphans = [row for row in previous if row["bookmark"] and row["bookmark"] not in current]
    processed = processed or []
    processed_names = {item["bookmark"] for item in processed}
    orphans = [row for row in orphans if row["bookmark"] not in processed_names]

    label_by_bookmark = {row["bookmark"]: row["label"] for row in previous if row["bookmark"]}
    used_numbers = set()
    for label in label_by_bookmark.values():
        if isinstance(label, str) and label.startswith("M") and label[1:].isdigit():
            used_numbers.add(int(label[1:]))
    next_number = 1

    def label_for(bookmark: str) -> str:
        nonlocal next_number
        if bookmark in label_by_bookmark:
            return label_by_bookmark[bookmark]
        while next_number in used_numbers:
            next_number += 1
        used_numbers.add(next_number)
        label = f"M{next_number:02d}"
        label_by_bookmark[bookmark] = label
        return label

    if SHEET_MARKS in wb.sheetnames:
        ws = wb[SHEET_MARKS]
        if ws.max_row:
            ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet(SHEET_MARKS)

    for column, title in enumerate(HEADERS, start=1):
        cell = ws.cell(1, column, title)
        cell.font = Font(name=FONT_NAME, size=10, bold=True, color="1F4E79")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    row_index = 2
    for mark in marks:
        old = previous_by_bookmark.get(mark.bookmark, {})
        ws.cell(row_index, 1, label_for(mark.bookmark))
        ws.cell(row_index, 2, mark.bookmark)
        ws.cell(row_index, 3, mark.location)
        ws.cell(row_index, 4, mark.text)
        if old.get("new"):
            ws.cell(row_index, 5, old["new"])
        ws.cell(row_index, 6, old.get("link", ""))
        ws.cell(row_index, 8, "全量")
        ws.cell(row_index, 9, row_index - 1)
        link_cell = ws.cell(row_index, 7, "→定位Word")
        link_cell.hyperlink = Hyperlink(
            ref=link_cell.coordinate, target=docx_name, location=mark.bookmark, display="→定位Word"
        )
        _style_row(ws, row_index)
        row_index += 1

    for item in processed:
        old = previous_by_bookmark.get(item["bookmark"], {})
        ws.cell(row_index, 1, label_for(item["bookmark"]))
        ws.cell(row_index, 2, item["bookmark"])
        ws.cell(row_index, 3, f"{item['location']}（已处理）")
        ws.cell(row_index, 4, item["text"] or old.get("original", ""))
        if old.get("new"):
            ws.cell(row_index, 5, old["new"])
        if old.get("link"):
            ws.cell(row_index, 6, old["link"])
        ws.cell(row_index, 7, "（已处理）")
        ws.cell(row_index, 8, "全量")
        ws.cell(row_index, 9, row_index - 1)
        _style_row(ws, row_index, deleted=True)
        row_index += 1

    for orphan in orphans:
        ws.cell(row_index, 1, label_for(orphan["bookmark"]) if orphan.get("bookmark") else orphan.get("label", ""))
        ws.cell(row_index, 2, orphan.get("bookmark", ""))
        position = (orphan.get("position") or "").replace("（Word 中已删除）", "")
        ws.cell(row_index, 3, f"{position}（Word 中已删除）")
        ws.cell(row_index, 4, orphan.get("original", ""))
        if orphan.get("new"):
            ws.cell(row_index, 5, orphan["new"])
        if orphan.get("link"):
            ws.cell(row_index, 6, orphan["link"])
        ws.cell(row_index, 7, "（已删除）")
        ws.cell(row_index, 8, "全量")
        ws.cell(row_index, 9, row_index - 1)
        _style_row(ws, row_index, deleted=True)
        row_index += 1

    for column, width in {1: 7, 2: 12, 3: 18, 4: 26, 5: 26, 6: 34, 7: 12, 8: 10, 9: 10}.items():
        ws.column_dimensions[get_column_letter(column)].width = width
    ws.freeze_panes = "A2"
    return {"marks": len(marks), "processed": len(processed), "orphans": len(orphans)}


def append_marks_sheet(wb, docx_name: str, marks: list, previous: list[dict]) -> dict:
    """只追加新增标记；既有行、行号、公式和填写值完全不动。"""
    known = {row.get("bookmark") for row in previous if row.get("bookmark")}
    marks = [mark for mark in marks if not mark.bookmark or mark.bookmark not in known]
    if not marks:
        return {"marks": 0, "processed": 0, "orphans": 0}
    if SHEET_MARKS in wb.sheetnames:
        ws = wb[SHEET_MARKS]
    else:
        ws = wb.create_sheet(SHEET_MARKS)
        previous = []
    for column, title in enumerate(HEADERS, start=1):
        cell = ws.cell(1, column, title)
        cell.font = Font(name=FONT_NAME, size=10, bold=True, color="1F4E79")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    known_labels = {str(row.get("label") or "") for row in previous}
    next_number = 1
    while f"XM{next_number:02d}" in known_labels:
        next_number += 1
    sequence = max([int(row.get("sequence") or 0) for row in previous] + [0]) + 1
    added = 0
    for mark in marks:
        while f"XM{next_number:02d}" in known_labels:
            next_number += 1
        label = f"XM{next_number:02d}"
        next_number += 1
        known_labels.add(label)
        row_index = max(ws.max_row or 1, 1) + 1
        values = [label, mark.bookmark, mark.location, mark.text, "", "", "→定位Word", "增量", sequence]
        for column, value in enumerate(values, start=1):
            ws.cell(row_index, column, value)
        link_cell = ws.cell(row_index, 7)
        link_cell.hyperlink = Hyperlink(
            ref=link_cell.coordinate, target=docx_name, location=mark.bookmark, display="→定位Word"
        )
        _style_row(ws, row_index)
        sequence += 1
        added += 1
    for column, width in {1: 7, 2: 12, 3: 18, 4: 26, 5: 26, 6: 34, 7: 12, 8: 10, 9: 10}.items():
        ws.column_dimensions[get_column_letter(column)].width = width
    ws.freeze_panes = "A2"
    return {"marks": added, "processed": 0, "orphans": 0}


def _style_row(ws, row_index: int, deleted: bool = False) -> None:
    for column in range(1, len(HEADERS) + 1):
        cell = ws.cell(row_index, column)
        cell.font = DELETED_FONT if deleted else Font(name=FONT_NAME, size=10)
        cell.alignment = Alignment(vertical="center", wrap_text=column in (3, 4, 5, 6))
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
