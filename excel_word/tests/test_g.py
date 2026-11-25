"""G：导出时合计行自动写求和公式（改明细后合计自动更新）。"""
from helpers import *  # noqa

ROWS = [["项目", "金额", "备注"], ["甲", "100", ""], ["乙", "200", ""], ["合计", "300", ""]]


def t23_total_formula_export():
    d = ws_dir("t23")
    docx = make_doc(d / "g.docx", [("表1 明细", ROWS)])
    api.run_export(docx)
    wb, order, blocks = blocks_of(api.workbook_path(docx))
    ws = wb[SHEET_NAME]
    gr = blocks["T01"]["grid_rows"]
    check("t23 合计金额列为SUM公式", ws.cell(gr[3], 4).value == f"=SUM(D{gr[1]}:D{gr[2]})",
          str(ws.cell(gr[3], 4).value))
    check("t23 文字备注列不写公式", ws.cell(gr[3], 5).value is None, str(ws.cell(gr[3], 5).value))
    check("t23 明细行仍为数值", ws.cell(gr[1], 4).value == 100, str(ws.cell(gr[1], 4).value))
    check("t23 合计公式格式保留千分位", "#,##0" in (ws.cell(gr[3], 4).number_format or ""),
          str(ws.cell(gr[3], 4).number_format))


def t24_total_auto_update():
    d = ws_dir("t24")
    docx = make_doc(d / "g.docx", [("表1 明细", ROWS)])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, order, blocks = blocks_of(xlsx)
    gr = blocks["T01"]["grid_rows"]
    wb[SHEET_NAME].cell(gr[1], 4).value = 500  # 甲 100 → 500
    wb.save(str(xlsx))
    api.run_writeback(docx)
    t = word_tables(docx)[0]
    check("t24 明细已改", t[1][1] == "500", str(t[1]))
    check("t24 合计随明细自动更新", t[3][1] == "700", str(t[3]))


def t25_total_formula_off():
    d = ws_dir("t25")
    docx = make_doc(d / "g.docx", [("表1 明细", ROWS)])
    api.run_export(docx, totals_formula=False)
    wb, order, blocks = blocks_of(api.workbook_path(docx))
    gr = blocks["T01"]["grid_rows"]
    check("t25 关闭后合计为原数值", wb[SHEET_NAME].cell(gr[3], 4).value == 300,
          str(wb[SHEET_NAME].cell(gr[3], 4).value))


def t26_export_backup_option():
    d = ws_dir("t26")
    docx = make_doc(d / "g.docx", [("表1 明细", ROWS)])
    bak = d / "_backup"

    def n_bak():
        return len(list(bak.glob("*"))) if bak.exists() else 0

    api.run_export(docx)  # 首次：Word 插书签会备份一次
    n1 = n_bak()
    check("t26 默认备份", n1 >= 1, str(n1))
    api.run_export(docx, backup=False)
    check("t26 关备份不新增", n_bak() == n1, str(n_bak()))
    api.run_export(docx, backup=True)
    check("t26 开备份覆盖前留底", n_bak() == n1 + 1, str(n_bak()))