"""C组：链接标注、书签工具、按金额排序。"""
from helpers import *  # noqa
from docx.enum.text import WD_COLOR_INDEX  # noqa


def t09_marks():
    d = ws_dir("t09")
    docx = copy_sample(d)
    api.run_export(docx)
    doc = Document(str(docx))
    p = doc.add_paragraph()
    run = p.add_run("待替换文字ABC")
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.save(str(docx))
    api.run_marks(docx)
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    check("t09 标注表生成", "链接标注" in wb.sheetnames)
    ws = wb["链接标注"]
    row = next((r for r in range(2, ws.max_row + 1) if ws.cell(r, 4).value == "待替换文字ABC"), None)
    check("t09 找到标注行", row is not None)
    if row is None:
        return
    ws.cell(row, 5).value = "替换后的文字XYZ"
    ws.cell(row, 6).value = "https://example.com"
    wb.save(str(xlsx))
    api.run_marks_sync(docx)
    doc = Document(str(docx))
    p = next(p for p in doc.paragraphs if "替换后的文字XYZ" in p.text)
    check("t09 文字已替换", "待替换文字ABC" not in p.text)
    check("t09 高亮已去除", all(r.font.highlight_color is None for r in p.runs), str([r.font.highlight_color for r in p.runs]))
    check("t09 超链接已加", "hyperlink" in p._element.xml and any(
        "example.com" in str(rel.target_ref) for rel in doc.part.rels.values()))
    # 重新导出标注：已处理行保留记录
    api.run_marks(docx)
    wb = load_workbook(str(xlsx))
    ws = wb["链接标注"]
    check("t09 刷新后行仍在", any("替换后的文字XYZ" in str(ws.cell(r, 4).value or "") or "（已处理）" in str(ws.cell(r, 1).value or "") for r in range(2, ws.max_row + 1)))


def t10_bookmark():
    d = ws_dir("t10")
    docx = copy_sample(d)
    api.run_export(docx)
    before = word_tables(docx)
    api.run_bookmark(docx)
    check("t10 书签工具不动数据", word_tables(docx) == before)
    api.run_bookmark(docx)
    check("t10 重复执行幂等", word_tables(docx) == before)
    api.run_clear_bookmarks(docx)
    from wesync.docx_bookmarks import bookmark_names
    names = bookmark_names(Document(str(docx)).element.body)
    check("t10 清理后无工具书签", not any(n.startswith(("Table_", "CPAH_", "TBL_", "MRK_")) for n in names), str(names)[:200])
    check("t10 清理不动数据", word_tables(docx) == before)


def t14_highlight_rules():
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from wesync.docx_marks import apply_mark_change, ensure_mark_bookmarks, find_marks, mark_payload_by_bookmark
    from wesync.errors import SyncError

    d = ws_dir("t14")
    path = d / "高亮.docx"
    doc = Document()
    base = doc.styles.add_style("高亮基样式", WD_STYLE_TYPE.CHARACTER)
    base.font.highlight_color = WD_COLOR_INDEX.YELLOW
    child = doc.styles.add_style("高亮子样式", WD_STYLE_TYPE.CHARACTER)
    child.base_style = base
    doc.add_paragraph().add_run("样式继承标注").style = child
    overridden = doc.add_paragraph().add_run("本地取消高亮")
    overridden.style = child
    rpr = overridden._r.get_or_add_rPr()
    hl = OxmlElement("w:highlight")
    hl.set(qn("w:val"), "none")
    rpr.append(hl)
    paragraph_mark = doc.add_paragraph()
    ppr = paragraph_mark._p.get_or_add_pPr()
    mark_rpr = OxmlElement("w:rPr")
    mark_hl = OxmlElement("w:highlight")
    mark_hl.set(qn("w:val"), "yellow")
    mark_rpr.append(mark_hl)
    ppr.append(mark_rpr)
    paragraph_mark.add_run("整段标注")
    doc.add_paragraph().add_run("   ").font.highlight_color = WD_COLOR_INDEX.YELLOW
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).paragraphs[0].add_run("表内不导出").font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.sections[0].header.paragraphs[0].add_run("页眉标注").font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.save(str(path))

    doc = Document(str(path))
    texts = [mark.text for mark in find_marks(doc)]
    check("t14 样式继承与段落标记", "样式继承标注" in texts and "整段标注" in texts, str(texts))
    check("t14 页眉包含", "页眉标注" in texts, str(texts))
    check("t14 表格与空白排除", "表内不导出" not in texts and "本地取消高亮" not in texts, str(texts))
    marks = find_marks(doc)
    style_mark = next(mark for mark in marks if mark.text == "样式继承标注")
    ensure_mark_bookmarks(doc, marks)
    payload = mark_payload_by_bookmark(doc, style_mark.bookmark)
    check("t14 定位标注", payload is not None)
    if payload:
        p_el, children, runs = payload
        apply_mark_change(doc.part, p_el, children, runs, "已替换", "", doc=doc)
        local = runs[0].find(qn("w:rPr")).find(qn("w:highlight"))
        check("t14 去色使用本地覆盖", local is not None and local.get(qn("w:val")) == "none")

    duplicate = Document(str(path))
    p_el = duplicate.add_paragraph()._p
    for bookmark_id in ("9001", "9002"):
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), bookmark_id)
        start.set(qn("w:name"), "MRK_0001")
        p_el.append(start)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), bookmark_id)
        p_el.append(end)
    try:
        find_marks(duplicate)
        raise AssertionError("t14 重复书签应报错")
    except SyncError as exc:
        check("t14 重复书签明确报错", "重复的标注书签" in str(exc), str(exc))


def t11_sort():
    d = ws_dir("t11")
    rows = [["项目", "金额", "备注"], ["甲", "300", ""], ["乙", "100", ""],
            ["其他", "50", ""], ["丙", "200", ""], ["合计", "650", ""]]
    docx = make_doc(d / "s.docx", [("表1 明细", rows)])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, order, blocks = blocks_of(xlsx)
    ws = wb[SHEET_NAME]
    b = blocks["T01"]
    gr = b["grid_rows"]
    ws.cell(gr[5], 4).value = f"=SUM(D{gr[1]}:D{gr[4]})"
    wb.save(str(xlsx))
    api.set_table_sort(docx, code="T01", column="金额", descending=True)
    api.run_writeback(docx)
    wb, order, blocks = blocks_of(xlsx)
    grid = grid_text(wb[SHEET_NAME], blocks["T01"])
    order_items = [r[0] for r in grid]
    check("t11 Excel排序正确", order_items == ["项目", "甲", "丙", "乙", "其他", "合计"], str(order_items))
    t = word_tables(docx)[0]
    check("t11 Word同步排序", [r[0] for r in t] == ["项目", "甲", "丙", "乙", "其他", "合计"], str([r[0] for r in t]))
    check("t11 合计公式值正确", t[5][1] == "650", str(t[5]))
    wb, order, blocks = blocks_of(xlsx)
    f = wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][5], 4).value
    check("t11 公式引用已平移", f == f"=SUM(D{blocks['T01']['grid_rows'][1]}:D{blocks['T01']['grid_rows'][4]})", str(f))
    api.run_export(docx)
    sorts = api.list_table_sorts(docx)
    check("t11 重导出沿用设置", sorts and sorts[0]["has_sort"] and sorts[0]["sort_column"] == "金额", str(sorts))
    api.clear_table_sort(docx, code="T01")
    check("t11 清除设置", not api.list_table_sorts(docx)[0]["has_sort"])

def t12_marks_display_format():
    from datetime import datetime

    d = ws_dir("t12")
    docx = make_doc(d / "报告.docx", [("表1", [["项目", "金额"], ["甲", "1"]])])
    doc = Document(str(docx))
    r1 = doc.add_paragraph().add_run("日期待填")
    r1.font.highlight_color = WD_COLOR_INDEX.YELLOW
    r2 = doc.add_paragraph().add_run("金额待填")
    r2.font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.save(str(docx))
    api.run_export(docx)
    api.run_marks(docx)
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    ws = wb["链接标注"]
    for r in range(2, ws.max_row + 1):
        orig = str(ws.cell(r, 4).value or "")
        if orig == "日期待填":
            ws.cell(r, 5).value = datetime(2025, 12, 31)
            ws.cell(r, 5).number_format = '[$-804]yyyy"年"m"月"d"日"'
        elif orig == "金额待填":
            ws.cell(r, 5).value = 45805.36
            ws.cell(r, 5).number_format = r'_ * #,##0.00_ ;_ * \-#,##0.00_ ;_ * "-"??_ ;_ @_ '
    wb.save(str(xlsx))
    api.run_marks_sync(docx)
    texts = [p.text for p in Document(str(docx)).paragraphs]
    check("t12 标注日期按显示格式", any("2025年12月31日" in t for t in texts), str(texts))
    check("t12 标注金额按显示格式", any("45,805.36" in t for t in texts), str(texts))
    check("t12 无原始时间串", not any("00:00:00" in t for t in texts), str(texts))


def t13_builtin_short_date():
    from datetime import datetime
    from unittest.mock import patch

    d = ws_dir("t13")
    docx = make_doc(d / "报告.docx", [("表1", [["项目", "日期"], ["甲", ""]])])
    doc = Document(str(docx))
    run = doc.add_paragraph().add_run("日期待填")
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.save(str(docx))
    api.run_export(docx)
    api.run_marks(docx)
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    ws = wb["链接标注"]
    row = next(r for r in range(2, ws.max_row + 1) if ws.cell(r, 4).value == "日期待填")
    cell = ws.cell(row, 5)
    cell.value = datetime(2026, 11, 12)
    cell.number_format = "mm-dd-yy"
    check("t13 使用Excel内置短日期格式", cell._style.numFmtId == 14, str(cell._style.numFmtId))
    wb.save(str(xlsx))
    with patch("wesync.valuefmt._windows_short_date_format", return_value="yyyy/m/d"):
        api.run_marks_sync(docx)
    texts = [p.text for p in Document(str(docx)).paragraphs]
    check("t13 内置短日期按区域显示", any("2026/11/12" in t for t in texts), str(texts))
    check("t13 不出现两位年份", not any("11-12-26" in t for t in texts), str(texts))
