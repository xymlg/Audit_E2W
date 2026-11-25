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
            ws.cell(r, 5).number_format = '_ * #,##0.00_ ;_ * \-#,##0.00_ ;_ * "-"??_ ;_ @_ '
    wb.save(str(xlsx))
    api.run_marks_sync(docx)
    texts = [p.text for p in Document(str(docx)).paragraphs]
    check("t12 标注日期按显示格式", any("2025年12月31日" in t for t in texts), str(texts))
    check("t12 标注金额按显示格式", any("45,805.36" in t for t in texts), str(texts))
    check("t12 无原始时间串", not any("00:00:00" in t for t in texts), str(texts))
