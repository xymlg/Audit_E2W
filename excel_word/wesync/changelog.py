"""改动留痕清单：把一次回写的逐格改动（旧值→新值）导出成 Excel 复核底稿。

记录由回写过程收集（见 writeback_cmd），本模块只负责落盘成清单。
清单保存在工作簿同目录，命名 <报告主名>_改动清单_<时间戳>.xlsx，不覆盖报告与工作簿。
"""

from __future__ import annotations

import datetime as _dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from .config import FONT_NAME, FONT_SIZE
from .console import log

HEADERS = ["表号", "表名", "行", "列", "项目", "旧值", "新值", "类型"]
_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")


def export_changelog(records: list[dict], xlsx_path: Path, docx_name: str) -> Path | None:
    """把改动记录导出为 Excel 清单；无记录时返回 None。"""
    if not records:
        return None
    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = Path(docx_name).stem
    target = xlsx_path.parent / f"{stem}_改动清单_{stamp}.xlsx"
    counter = 1
    while target.exists():
        target = xlsx_path.parent / f"{stem}_改动清单_{stamp}_{counter}.xlsx"
        counter += 1

    wb = Workbook()
    ws = wb.active
    ws.title = "改动清单"
    ws.cell(1, 1, f"报告：{docx_name}    生成时间：{_dt.datetime.now():%Y-%m-%d %H:%M:%S}    共 {len(records)} 条改动")
    ws.cell(1, 1).font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True, color="1F4E79")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    for c, title in enumerate(HEADERS, 1):
        cell = ws.cell(2, c, title)
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True, color="FFFFFF")
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for i, rec in enumerate(records, 3):
        values = [rec.get("code", ""), rec.get("name", ""), rec.get("row", ""), rec.get("col", ""),
                  rec.get("item", ""), rec.get("old", ""), rec.get("new", ""), rec.get("kind", "")]
        for c, value in enumerate(values, 1):
            cell = ws.cell(i, c, value)
            cell.font = Font(name=FONT_NAME, size=FONT_SIZE)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    widths = [8, 28, 6, 6, 24, 24, 24, 10]
    for c, width in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(2, c).column_letter].width = width
    ws.freeze_panes = "A3"
    wb.save(str(target))
    log(f"[留痕] 已导出改动清单：{target.name}（{len(records)} 条）")
    return target