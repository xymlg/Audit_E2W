"""正文占位符：扫描/替换 Word 正文与页眉页脚里的 {字段名}，值来自工作簿「正文信息」表。

替换只动占位符文字，保留所在 run 的字体格式；跨 run 的占位符也能整体替换。
未填值的占位符不替换（由调用方决定跳过该份报告）。
"""

from __future__ import annotations

import re

PLACEHOLDER_RE = re.compile(r"\{([^{}\r\n]{1,40})\}")
VALUES_SHEET = "正文信息"


def _iter_paragraphs(doc):
    """正文段落（含表格内）+ 各节页眉页脚（含其中表格）。"""
    def _walk(container):
        yield from container.paragraphs
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    yield from _walk(cell)

    yield from _walk(doc)
    for section in doc.sections:
        for part in (section.header, section.footer,
                     section.first_page_header, section.first_page_footer,
                     section.even_page_header, section.even_page_footer):
            if part is not None:
                yield from _walk(part)


def scan_placeholders(doc) -> list[str]:
    """按出现顺序去重列出全部占位符字段名。"""
    seen: dict[str, None] = {}
    for p in _iter_paragraphs(doc):
        for m in PLACEHOLDER_RE.finditer(p.text):
            field = m.group(1).strip()
            if field:
                seen.setdefault(field, None)
    return list(seen)


def _replace_in_paragraph(p, values: dict[str, str]) -> list[tuple[str, str]]:
    runs = p.runs
    if not runs:
        return []
    full = "".join(r.text for r in runs)
    matches = [m for m in PLACEHOLDER_RE.finditer(full) if m.group(1).strip() in values]
    if not matches:
        return []
    spans = []
    pos = 0
    for r in runs:
        spans.append((pos, pos + len(r.text)))
        pos += len(r.text)
    done = []
    for m in reversed(matches):  # 从右往左改，前面的偏移不受影响
        field = m.group(1).strip()
        replacement = values[field]
        s, e = m.span()
        first = True
        for r, (rs, re_) in zip(runs, spans):
            if re_ <= s or rs >= e:
                continue
            head = r.text[: max(0, s - rs)] if first else ""
            tail = r.text[max(0, min(len(r.text), e - rs)):]
            r.text = head + (replacement if first else "") + tail
            first = False
        done.append((field, replacement))
    return list(reversed(done))


def apply_placeholders(doc, values: dict[str, str], *, dry_run: bool = False) -> dict:
    """返回 {"replaced": [(字段, 新值)], "missing": [字段]}；dry_run 只列计划不改文档。"""
    fields = scan_placeholders(doc)
    missing = [f for f in fields if not str(values.get(f, "")).strip()]
    if dry_run:
        return {"replaced": [(f, str(values[f])) for f in fields if f not in missing], "missing": missing}
    replaced = []
    effective = {f: str(values[f]) for f in fields if f not in missing}
    for p in _iter_paragraphs(doc):
        replaced.extend(_replace_in_paragraph(p, effective))
    return {"replaced": replaced, "missing": missing}


def read_placeholder_values(xlsx_path, *, resolve_formulas: bool = True) -> dict[str, str]:
    """读工作簿「正文信息」表（A=字段 B=值，首行表头）；无该表返回空。

    值按单元格在 Excel 里的显示效果转换（与表格回写同口径）：日期用其数字格式渲染
    （如 2026-12-31、2026年12月31日），不会出现 "2026-12-31 00:00:00" 这类原始值。
    公式优先取 Excel 缓存结果，其次轻量求值；导出保留旧值时可按原公式保留。
    """
    from openpyxl import load_workbook

    from .errors import SyncError
    from .valuefmt import display_text, formula_result

    wb = load_workbook(str(xlsx_path))
    if VALUES_SHEET not in wb.sheetnames:
        return {}
    cached_wb = load_workbook(str(xlsx_path), data_only=True) if resolve_formulas else None
    ws = wb[VALUES_SHEET]
    cached_ws = cached_wb[VALUES_SHEET] if cached_wb is not None else None
    values = {}
    for row in range(2, (ws.max_row or 1) + 1):
        field = str(ws.cell(row, 1).value or "").strip()
        cell = ws.cell(row, 2)
        raw = cell.value
        if resolve_formulas and isinstance(raw, str) and raw.startswith("="):
            raw = formula_result(ws, row, 2, cached_ws)
            if raw is None:
                raise SyncError(
                    f"「{VALUES_SHEET}」{cell.coordinate} 是公式但没有可用的计算结果。"
                    "请先在 Excel 中打开该工作簿，完成重算并保存后再同步。"
                )
        if field:
            if not resolve_formulas and isinstance(raw, str) and raw.startswith("="):
                values[field] = raw
            else:
                values[field] = display_text(raw, cell.number_format)
    return values


def init_placeholder_sheet(wb):
    """在工作簿对象中新建「正文信息」骨架表（表头/列宽），返回工作表。"""
    from openpyxl.styles import Font

    from .config import FONT_NAME, FONT_SIZE

    ws = wb.create_sheet(VALUES_SHEET)
    ws.cell(1, 1, "字段")
    ws.cell(1, 2, "值")
    for c in (1, 2):
        ws.cell(1, c).font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 48
    return ws


def write_placeholder_sheet(xlsx_path, fields: list[str]) -> list[str]:
    """给工作簿生成/补齐「正文信息」填写骨架，保留已填值；返回新增字段。"""
    from .fsops import backup_file, check_file_free, load_workbook_for_save

    check_file_free(xlsx_path, False)
    wb = load_workbook_for_save(xlsx_path)
    created = VALUES_SHEET not in wb.sheetnames
    ws = init_placeholder_sheet(wb) if created else wb[VALUES_SHEET]
    existing = {str(ws.cell(r, 1).value or "").strip() for r in range(2, (ws.max_row or 1) + 1)}
    added = []
    for field in fields:
        if field not in existing:
            ws.append([field, ""])
            added.append(field)
    if added or created:
        backup_file(xlsx_path)
        wb.save(str(xlsx_path))
    wb.close()
    if getattr(wb, "vba_archive", None):
        wb.vba_archive.close()
    return added
