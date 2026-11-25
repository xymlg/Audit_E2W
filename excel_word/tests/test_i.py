"""I：占位符导出/同步、同步内容组合（单份+批量）、标注未填警告、显示格式回写。"""
from datetime import datetime

from helpers import *  # noqa
from wesync.config import MARK_YES  # noqa
from wesync.docx_placeholders import VALUES_SHEET  # noqa
from wesync.errors import SyncError  # noqa
from wesync.excel_marks import SHEET_MARKS  # noqa
from wesync.valuefmt import display_text  # noqa


def _ph_doc(path, *, header=False, cross_run=False):
    doc = Document()
    if cross_run:
        p = doc.add_paragraph("审计期间：")
        p.add_run("{审计")
        p.add_run("期间}")
        p.add_run("，编号：{报告编号}。")
    else:
        doc.add_paragraph("审计期间：{审计期间}，编号：{报告编号}。")
    t = doc.add_table(rows=2, cols=2)
    t.cell(0, 0).text = "项目"
    t.cell(0, 1).text = "金额"
    t.cell(1, 0).text = "资产总额"
    t.cell(1, 1).text = "{资产总额}"
    if header:
        doc.sections[0].header.paragraphs[0].add_run("单位：{单位名称}")
    doc.save(str(path))
    return path


def _fill_values(xlsx, mapping):
    wb = load_workbook(str(xlsx))
    ws = wb[VALUES_SHEET]
    rows = {str(ws.cell(r, 1).value or "").strip(): r for r in range(2, ws.max_row + 1)}
    for field, value in mapping.items():
        ws.cell(rows[field], 2).value = value
    wb.save(str(xlsx))


def t31_placeholder_roundtrip():
    d = ws_dir("t31")
    docx = _ph_doc(d / "i.docx", header=True, cross_run=True)
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    api.run_placeholder_export(docx)
    wb = load_workbook(str(xlsx))
    check("t31 建正文信息表", VALUES_SHEET in wb.sheetnames, str(wb.sheetnames))
    ws = wb[VALUES_SHEET]
    fields = [str(ws.cell(r, 1).value) for r in range(2, ws.max_row + 1)]
    check("t31 字段齐全", set(fields) == {"审计期间", "报告编号", "资产总额", "单位名称"}, str(fields))
    _fill_values(xlsx, {"审计期间": "2025年度", "报告编号": "XYZH/2026BJA1", "单位名称": "某某公司"})
    api.run_placeholder_export(docx)  # 再导出保留已填值
    wb = load_workbook(str(xlsx))
    ws = wb[VALUES_SHEET]
    kept = {str(ws.cell(r, 1).value): ws.cell(r, 2).value for r in range(2, ws.max_row + 1)}
    check("t31 保留已填值", kept["审计期间"] == "2025年度" and kept["资产总额"] in (None, ""), str(kept))
    before = docx.read_bytes()
    LOGS.clear()
    api.run_placeholder_sync(docx, dry_run=True)
    check("t31 预演不写盘", docx.read_bytes() == before)
    check("t31 未填警告", any("未填值" in m and "资产总额" in m for m in LOGS), str(LOGS[-4:]))
    api.run_placeholder_sync(docx)
    doc = Document(str(docx))
    check("t31 跨run正文替换", "2025年度" in doc.paragraphs[0].text and "{审计" not in doc.paragraphs[0].text,
          doc.paragraphs[0].text)
    check("t31 表格内替换", doc.tables[0].cell(1, 1).text == "{资产总额}", doc.tables[0].cell(1, 1).text)
    check("t31 页眉替换", "某某公司" in doc.sections[0].header.paragraphs[0].text,
          doc.sections[0].header.paragraphs[0].text)
    # 补填后再同步
    _fill_values(xlsx, {"资产总额": "5,654.00"})
    api.run_placeholder_sync(docx)
    check("t31 补填后替换", Document(str(docx)).tables[0].cell(1, 1).text == "5,654.00")


def t32_placeholder_zero_fields():
    d = ws_dir("t32")
    docx = make_doc(d / "i.docx", [("表1", [["项目", "金额"], ["甲", "1"]])])
    api.run_export(docx)  # 导出总会建「正文信息」表（此处应为空表）
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    check("t32 导出后正文信息为空表", VALUES_SHEET in wb.sheetnames and wb[VALUES_SHEET].max_row == 1,
          str((wb.sheetnames, wb[VALUES_SHEET].max_row if VALUES_SHEET in wb.sheetnames else None)))
    before = xlsx.read_bytes()
    LOGS.clear()
    api.run_placeholder_export(docx)
    wb = load_workbook(str(xlsx))
    check("t32 无占位符不新增行", wb[VALUES_SHEET].max_row == 1, str(wb[VALUES_SHEET].max_row))
    check("t32 不写盘", xlsx.read_bytes() == before)
    check("t32 有提示", any("占位符" in m for m in LOGS), str(LOGS[-3:]))


def _combo_doc(path):
    doc = Document()
    p = doc.add_paragraph("金额核对：")
    hl = p.add_run("待核对")
    from docx.enum.text import WD_COLOR_INDEX
    hl.font.highlight_color = WD_COLOR_INDEX.YELLOW
    p.add_run("，期间：{审计期间}。")
    t = doc.add_table(rows=2, cols=2)
    t.cell(0, 0).text = "项目"
    t.cell(0, 1).text = "金额"
    t.cell(1, 0).text = "甲"
    t.cell(1, 1).text = "100"
    doc.save(str(path))
    return path


def _prepare_combo(d):
    docx = _combo_doc(d / "i.docx")
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    api.run_placeholder_export(docx)
    wb = load_workbook(str(xlsx))
    _, _, blocks = blocks_of(xlsx)
    wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 4).value = 200
    wb[SHEET_MARKS].cell(2, 5).value = "已核对无误"
    wb.save(str(xlsx))
    _fill_values(xlsx, {"审计期间": "2025年度"})
    return docx, xlsx


def t33_content_combos():
    d = ws_dir("t33")
    docx, xlsx = _prepare_combo(d)
    # 只同步表格：标注/占位符不动
    api.run_sync_content(docx, xlsx, tables=True)
    doc = Document(str(docx))
    check("t33 表格已同步", doc.tables[0].cell(1, 1).text == "200", doc.tables[0].cell(1, 1).text)
    check("t33 标注未动", "待核对" in doc.paragraphs[0].text, doc.paragraphs[0].text)
    check("t33 占位符未动", "{审计期间}" in doc.paragraphs[0].text, doc.paragraphs[0].text)
    # 只同步标注+占位符
    api.run_sync_content(docx, xlsx, tables=False, marks=True, placeholders=True)
    doc = Document(str(docx))
    check("t33 标注已同步", "已核对无误" in doc.paragraphs[0].text, doc.paragraphs[0].text)
    check("t33 占位符已同步", "2025年度" in doc.paragraphs[0].text, doc.paragraphs[0].text)
    check("t33 默认保留高亮", any(r.font.highlight_color is not None for r in doc.paragraphs[0].runs),
          str([r.text for r in doc.paragraphs[0].runs]))
    # 全不勾 → 报错
    try:
        api.run_sync_content(docx, xlsx, tables=False, marks=False, placeholders=False)
        raise AssertionError("t33 空组合应报错")
    except SyncError as exc:
        check("t33 空组合报错", "至少勾选" in str(exc), str(exc))


def t34_content_missing_sheet_and_decolor():
    d = ws_dir("t34")
    docx = make_doc(d / "i.docx", [("表1", [["项目", "金额"], ["甲", "1"]])])
    api.run_export(docx)  # 无高亮 → 无「链接标注」表
    xlsx = api.workbook_path(docx)
    LOGS.clear()
    summary = api.run_sync_content(docx, xlsx, tables=False, marks=True)
    check("t34 缺表警告跳过", any("跳过标注" in m for m in LOGS), str(LOGS[-3:]))
    check("t34 摘要无标注", "marks" not in summary, str(summary))
    # 去高亮口径
    d2 = ws_dir("t34b")
    docx2, xlsx2 = _prepare_combo(d2)
    api.run_sync_content(docx2, xlsx2, tables=False, marks=True, keep_highlight=False)
    doc = Document(str(docx2))
    check("t34 同步后去色", all(r.font.highlight_color is None for r in doc.paragraphs[0].runs),
          str([(r.text, r.font.highlight_color) for r in doc.paragraphs[0].runs]))


def t35_marks_unfilled_warn():
    d = ws_dir("t35")
    docx, xlsx = _prepare_combo(d)
    wb = load_workbook(str(xlsx))
    wb[SHEET_MARKS].cell(2, 5).value = None  # 新内容留空、链接也空
    wb.save(str(xlsx))
    LOGS.clear()
    api.run_marks_sync(docx, dry_run=True)
    check("t35 未填逐行警告", any("均未填" in m for m in LOGS), str(LOGS[-4:]))
    check("t35 原文保留", "待核对" in Document(str(docx)).paragraphs[0].text)
    # 只填链接地址 → 视为已填，不警告
    wb = load_workbook(str(xlsx))
    wb[SHEET_MARKS].cell(2, 6).value = "https://example.com"
    wb.save(str(xlsx))
    LOGS.clear()
    api.run_marks_sync(docx, dry_run=True)
    check("t35 仅链接不警告", not any("均未填" in m for m in LOGS), str(LOGS[-4:]))


def t36_batch_contents():
    d = ws_dir("t36")
    paths = []
    for name in ("a", "b"):
        sub = d / name
        sub.mkdir()
        docx, xlsx = _prepare_combo(sub)
        paths.append(str(docx))
    results = api.run_batch_writeback(paths, contents=("tables", "marks", "placeholders"))
    check("t36 批量全成功", all(r["status"] == "ok" for r in results), str(results))
    for p in paths:
        doc = Document(p)
        check("t36 批量表格", doc.tables[0].cell(1, 1).text == "200", p)
        check("t36 批量标注", "已核对无误" in doc.paragraphs[0].text, p)
        check("t36 批量占位符", "2025年度" in doc.paragraphs[0].text, p)
    results = api.run_batch_writeback(paths, contents=())
    check("t36 空组合逐份失败", all(r["status"] == "failed" and "至少勾选" in r["reason"] for r in results),
          str(results))


def t37_display_format_writeback():
    d = ws_dir("t37")
    docx = make_doc(d / "i.docx", [("表1", [["项目", "金额"], ["甲", "1"], ["乙", "2"], ["丙", "3"], ["丁", "4"]])])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    ws = wb[SHEET_NAME]
    _, _, blocks = blocks_of(xlsx)
    gr = blocks["T01"]["grid_rows"]
    cases = [
        (5654, "#,##0.00", "5,654.00"),
        (datetime(2026, 12, 31), 'yyyy"年"m"月"d"日"', "2026年12月31日"),
        (0.1234, "0.00%", "12.34%"),
        (-1234.5, '#,##0.00;(#,##0.00)', "(1,234.50)"),
    ]
    for i, (value, fmt, _expected) in enumerate(cases):
        cell = ws.cell(gr[1 + i], 4)
        cell.value = value
        cell.number_format = fmt
    wb.save(str(xlsx))
    api.run_writeback(docx)
    got = [row[1] for row in word_tables(docx)[0][1:]]
    check("t37 按显示格式回写", got == [e for _, _, e in cases], str(got))
    # 子集外格式：回退旧口径 + 警告一次
    wb = load_workbook(str(xlsx))
    ws = wb[SHEET_NAME]
    ws.cell(gr[1], 4).value = 1.5
    ws.cell(gr[1], 4).number_format = "[h]:mm"
    wb.save(str(xlsx))
    LOGS.clear()
    api.run_writeback(docx)
    check("t37 怪格式回退", word_tables(docx)[0][1][1] == "2", str(word_tables(docx)))
    check("t37 怪格式警告", any("超出内置显示渲染子集" in m for m in LOGS), str(LOGS[-4:]))


def t38_display_text_units():
    check("t38 General去尾零", display_text(12345.0, "General") == "12345")
    check("t38 General不加千分位", display_text(12345, "General") == "12345")
    check("t38 千分位两位小数", display_text(5654, "#,##0.00") == "5,654.00")
    check("t38 无千分位", display_text(5654, "0.00") == "5654.00")
    check("t38 四舍五入", display_text(1234.565, "#,##0.00") == "1,234.57")
    check("t38 百分比", display_text(0.12, "0.00%") == "12.00%")
    check("t38 负数括号段", display_text(-1234.5, '#,##0.00;(#,##0.00)') == "(1,234.50)")
    check("t38 零值段", display_text(0, '#,##0.00;(#,##0.00);"—"') == "—")
    check("t38 日期序列值", display_text(46387, 'yyyy"年"m"月"d"日"') == "2026年12月31日")
    check("t38 日期时间分钟", display_text(datetime(2026, 1, 2, 3, 4, 5), "h:mm:ss") == "3:04:05")
    check("t38 井号去尾零", display_text(0.5, "0.0#") == "0.5")
    check("t38 千位缩放", display_text(1234567, "#,##0,") == "1,235")
    check("t38 空值", display_text(None, "#,##0.00") == "")


def t39_rejudge_raw_value():
    d = ws_dir("t39")
    docx = make_doc(d / "i.docx", [("表1", [["项目", "金额"], ["甲", ""], ["乙", "1"]])])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    ws = wb[SHEET_NAME]
    _, _, blocks = blocks_of(xlsx)
    cell = ws.cell(blocks["T01"]["grid_rows"][1], 4)  # 甲行金额：显示为空但原始值 0 → 应判✓
    cell.value = 0
    cell.number_format = '0;-0;""'
    wb.save(str(xlsx))
    check("t39 显示为空", display_text(cell.value, cell.number_format) == "")
    api.run_rejudge(docx)
    wb = load_workbook(str(xlsx))
    check("t39 按原始值判✓", str(wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 2).value).strip() == MARK_YES,
          str(wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 2).value))

def t40_template_contents():
    d = ws_dir("t40")
    tpl = _combo_doc(d / "模板.docx")
    api.run_export(tpl)
    xlsx = api.workbook_path(tpl)
    api.run_placeholder_export(tpl)
    wb, _, blocks = blocks_of(xlsx)
    wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 4).value = 200
    wb[SHEET_MARKS].cell(2, 5).value = "已核对无误"
    wb.save(str(xlsx))
    # 正文信息未填值：不勾占位符则不因此跳过
    out = d / "out"
    tpl_bytes, xlsx_bytes = tpl.read_bytes(), xlsx.read_bytes()
    results = api.run_template_batch(
        [tpl], [{"excel": str(xlsx), "template": 1}], out, contents=("tables", "marks")
    )
    check("t40 未填占位符不跳过", results[0]["status"] == "ok", str(results))
    doc = Document(results[0]["output"])
    check("t40 表格已同步", doc.tables[0].cell(1, 1).text == "200", doc.tables[0].cell(1, 1).text)
    check("t40 标注已同步", "已核对无误" in doc.paragraphs[0].text, doc.paragraphs[0].text)
    check("t40 占位符保留", "{审计期间}" in doc.paragraphs[0].text, doc.paragraphs[0].text)
    check("t40 模板未动", tpl.read_bytes() == tpl_bytes)
    check("t40 源Excel未动", xlsx.read_bytes() == xlsx_bytes)
    # 只勾标注：表格不动
    results = api.run_template_batch(
        [tpl], [{"excel": str(xlsx), "template": 1}], out, contents=("marks",), keep_highlight=False
    )
    check("t40 仅标注成功", results[0]["status"] == "ok", str(results))
    doc = Document(results[0]["output"])
    check("t40 表格未同步", doc.tables[0].cell(1, 1).text == "100", doc.tables[0].cell(1, 1).text)
    check("t40 标注已同步2", "已核对无误" in doc.paragraphs[0].text, doc.paragraphs[0].text)
    check("t40 去高亮", all(r.font.highlight_color is None for r in doc.paragraphs[0].runs),
          str([(r.text, r.font.highlight_color) for r in doc.paragraphs[0].runs]))
    # 预演：不产出文件
    results = api.run_template_batch(
        [tpl], [{"excel": str(xlsx), "template": 1}], d / "dry", dry_run=True, contents=("tables", "marks")
    )
    check("t40 预演成功", results[0]["status"] == "ok" and not (d / "dry").exists(), str(results))
    # 空组合 → 报错
    try:
        api.run_template_batch([tpl], [{"excel": str(xlsx), "template": 1}], out, contents=())
        raise AssertionError("t40 空组合应报错")
    except SyncError as exc:
        check("t40 空组合报错", "至少勾选" in str(exc), str(exc))


def t41_export_creates_placeholder_sheet():
    from wesync.docx_placeholders import read_placeholder_values
    d = ws_dir("t41")
    docx = make_doc(d / "i.docx", [("表1", [["项目", "金额"], ["甲", "1"]])])
    doc = Document(str(docx))
    doc.add_paragraph("被审计单位：{公司名称}，文号{报告文号}。")
    doc.save(str(docx))
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    check("t41 导出自动建正文信息表", VALUES_SHEET in wb.sheetnames, str(wb.sheetnames))
    values = read_placeholder_values(xlsx)
    check("t41 字段已写入", values == {"公司名称": "", "报告文号": ""}, str(values))
    check("t41 有日志", any(VALUES_SHEET in m for m in LOGS), str(LOGS[-4:]))
    _fill_values(xlsx, {"公司名称": "甲公司"})
    api.run_export(docx)
    values = read_placeholder_values(xlsx)
    check("t41 重复导出保留已填值", values.get("公司名称") == "甲公司", str(values))
    check("t41 未填字段仍在", values.get("报告文号") == "", str(values))


def t42_export_no_placeholders_empty_sheet():
    d = ws_dir("t42")
    docx = make_doc(d / "i.docx", [("表1", [["项目", "金额"], ["甲", "1"]])])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    check("t42 无占位符也建空表", VALUES_SHEET in wb.sheetnames and wb[VALUES_SHEET].max_row == 1,
          str((wb.sheetnames, wb[VALUES_SHEET].max_row if VALUES_SHEET in wb.sheetnames else None)))


def t43_placeholder_values_display_format():
    from datetime import datetime

    from openpyxl import Workbook

    from wesync.docx_placeholders import VALUES_SHEET as PH_SHEET, read_placeholder_values
    d = ws_dir("t43")
    xlsx = d / "v.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = PH_SHEET
    ws.cell(1, 1, "字段")
    ws.cell(1, 2, "值")
    ws.cell(2, 1, "日期")
    ws.cell(2, 2, datetime(2026, 12, 31))
    ws.cell(2, 2).number_format = "yyyy-mm-dd"
    ws.cell(3, 1, "金额")
    ws.cell(3, 2, 69235.36)
    ws.cell(3, 2).number_format = "#,##0.00"
    ws.cell(4, 1, "文本")
    ws.cell(4, 2, "甲公司")
    wb.save(str(xlsx))
    values = read_placeholder_values(xlsx)
    check("t43 日期按显示格式", values["日期"] == "2026-12-31", str(values))
    check("t43 金额按显示格式", values["金额"] == "69,235.36", str(values))
    check("t43 文本原样", values["文本"] == "甲公司", str(values))


def t44_display_accounting_and_locale_date():
    from datetime import datetime as dt_cls
    acct = '_ * #,##0.00_ ;_ * \-#,##0.00_ ;_ * "-"??_ ;_ @_ '
    cn_date = '[$-804]yyyy"年"m"月"d"日"'
    check("t44 会计正数", display_text(45805.36, acct) == "45,805.36", display_text(45805.36, acct))
    check("t44 会计负数", display_text(-45805.36, acct) == "-45,805.36", display_text(-45805.36, acct))
    check("t44 会计零值段", display_text(0, acct) == "-", display_text(0, acct))
    check("t44 区域前缀日期", display_text(dt_cls(2025, 12, 31), cn_date) == "2025年12月31日",
          display_text(dt_cls(2025, 12, 31), cn_date))
    check("t44 区域前缀日期序列值", display_text(46022, cn_date) == "2025年12月31日", display_text(46022, cn_date))


def t45_export_without_tables_creates_all_sheets():
    from wesync.docx_placeholders import read_placeholder_values as _read
    d = ws_dir("t45")
    # 仅占位符（无表格）
    docx = d / "仅占位符.docx"
    doc = Document()
    doc.add_paragraph("单位：{公司名称}，期间：{审计期间}。")
    doc.save(str(docx))
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    check("t45 无表格仍导出且三个工作表都在", wb.sheetnames == [SHEET_NAME, SHEET_MARKS, VALUES_SHEET], str(wb.sheetnames))
    check("t45 表格汇总为空", wb[SHEET_NAME].max_row == 1, str(wb[SHEET_NAME].max_row))
    check("t45 占位符已写入", set(_read(xlsx)) == {"公司名称", "审计期间"}, str(_read(xlsx)))
    # 仅标注（无表格）
    docx2 = d / "仅标注.docx"
    doc = Document()
    from docx.enum.text import WD_COLOR_INDEX
    run = doc.add_paragraph().add_run("待核对内容")
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.save(str(docx2))
    api.run_export(docx2)
    wb = load_workbook(str(api.workbook_path(docx2)))
    check("t45 仅标注三个工作表都在", wb.sheetnames == [SHEET_NAME, SHEET_MARKS, VALUES_SHEET], str(wb.sheetnames))
    check("t45 标注行已写入", wb[SHEET_MARKS].max_row >= 2, str(wb[SHEET_MARKS].max_row))
    check("t45 正文信息为空表", wb[VALUES_SHEET].max_row == 1, str(wb[VALUES_SHEET].max_row))
    # 纯正文（三者都没有）也能导出
    docx3 = d / "空.docx"
    doc = Document()
    doc.add_paragraph("纯正文。")
    doc.save(str(docx3))
    api.run_export(docx3)
    wb = load_workbook(str(api.workbook_path(docx3)))
    check("t45 三者皆无也导出三表", wb.sheetnames == [SHEET_NAME, SHEET_MARKS, VALUES_SHEET], str(wb.sheetnames))


def t46_marks_formula_uses_cached_result():
    from docx.enum.text import WD_COLOR_INDEX

    d = ws_dir("t46")
    docx = d / "标注公式.docx"
    doc = Document()
    run = doc.add_paragraph().add_run("待核对金额")
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.save(str(docx))
    api.run_export(docx)
    api.run_marks(docx)
    xlsx = api.workbook_path(docx)
    formula = "=[1]Sheet1!$C$5"
    wb = load_workbook(str(xlsx))
    ws = wb[SHEET_MARKS]
    ws["E2"] = formula
    ws["E2"].number_format = "#,##0.00"
    wb.save(str(xlsx))
    set_formula_cached_value(xlsx, SHEET_MARKS, "E2", 5654)
    api.run_marks_sync(docx, keep_highlight=True)
    text = Document(str(docx)).paragraphs[0].text
    check("t46 标注取公式缓存值", "5,654.00" in text and formula not in text, text)
    api.run_marks(docx)
    wb = load_workbook(str(xlsx))
    check("t46 刷新标注保留公式", wb[SHEET_MARKS]["E2"].value == formula, str(wb[SHEET_MARKS]["E2"].value))


def t47_placeholder_formula_result_and_missing_cache():
    from wesync.docx_placeholders import read_placeholder_values

    d = ws_dir("t47")
    docx = d / "占位符公式.docx"
    doc = Document()
    doc.add_paragraph("单位：{公司名称}，金额：{金额}。")
    doc.save(str(docx))
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    formula = "=[1]Sheet1!$C$5"
    wb = load_workbook(str(xlsx))
    ws = wb[VALUES_SHEET]
    ws["B2"] = formula
    ws["B3"] = "=100+200"
    ws["B3"].number_format = "#,##0.00"
    wb.save(str(xlsx))
    set_formula_cached_value(xlsx, VALUES_SHEET, "B2", "乙公司")
    values = read_placeholder_values(xlsx)
    check("t47 占位符取缓存与轻量求值", values == {"公司名称": "乙公司", "金额": "300.00"}, str(values))
    api.run_export(docx)
    wb = load_workbook(str(xlsx))
    check("t47 重导出保留占位符公式", wb[VALUES_SHEET]["B2"].value == formula
          and wb[VALUES_SHEET]["B3"].value == "=100+200", str(wb[VALUES_SHEET]["B2"].value))
    set_formula_cached_value(xlsx, VALUES_SHEET, "B2", "乙公司")
    api.run_placeholder_sync(docx)
    text = Document(str(docx)).paragraphs[0].text
    check("t47 占位符写结果不写公式", "乙公司" in text and "300" in text and formula not in text, text)
    wb = load_workbook(str(xlsx))
    wb[VALUES_SHEET]["B2"] = formula
    wb.save(str(xlsx))
    before = docx.read_bytes()
    try:
        api.run_placeholder_sync(docx)
        raise AssertionError("t47 无缓存外链公式应报错")
    except SyncError as exc:
        check("t47 无缓存给出明确提示", "Excel" in str(exc) and "保存" in str(exc), str(exc))
    check("t47 无缓存不写坏 Word", docx.read_bytes() == before)


def _prepare_formula_combo(d, name):
    docx = _combo_doc(d / f"{name}.docx")
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    _, _, blocks = blocks_of(xlsx)
    wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 4).value = 200
    wb[SHEET_MARKS]["E2"] = "=[1]Sheet1!$C$5"
    wb.save(str(xlsx))
    _fill_values(xlsx, {"审计期间": "=[1]Sheet1!$C$6"})
    set_formula_cached_value(xlsx, SHEET_MARKS, "E2", f"{name}标注")
    set_formula_cached_value(xlsx, VALUES_SHEET, "B2", f"{name}期间")
    return docx, xlsx


def t48_formula_results_cover_combined_batch_template():
    d = ws_dir("t48")
    docx, _ = _prepare_formula_combo(d, "组合")
    api.run_sync_content(docx, tables=True, marks=True, placeholders=True, keep_highlight=True)
    text = Document(str(docx)).paragraphs[0].text
    check("t48 组合同步取公式结果", "组合标注" in text and "组合期间" in text and "Sheet1" not in text, text)
    check("t48 组合表格同步", Document(str(docx)).tables[0].cell(1, 1).text == "200",
          Document(str(docx)).tables[0].cell(1, 1).text)

    paths = [_prepare_formula_combo(d, name)[0] for name in ("批量甲", "批量乙")]
    results = api.run_batch_writeback(paths, contents=("tables", "marks", "placeholders"))
    check("t48 批量同步全部成功", all(r["status"] == "ok" for r in results), str(results))
    for path, name in zip(paths, ("批量甲", "批量乙")):
        doc = Document(str(path))
        text = doc.paragraphs[0].text
        check(f"t48 {name}批量取公式结果", f"{name}标注" in text and f"{name}期间" in text, text)
        check(f"t48 {name}批量表格同步", doc.tables[0].cell(1, 1).text == "200", doc.tables[0].cell(1, 1).text)

    tpl, xlsx = _prepare_formula_combo(d, "模板")
    results = api.run_template_batch(
        [tpl], [{"excel": str(xlsx), "template": 1}], d / "out", contents=("tables", "marks", "placeholders")
    )
    check("t48 模板生成成功", results[0]["status"] == "ok", str(results))
    doc = Document(results[0]["output"])
    text = doc.paragraphs[0].text
    check("t48 模板取公式结果", "模板标注" in text and "模板期间" in text and "Sheet1" not in text, text)
    check("t48 模板表格同步", doc.tables[0].cell(1, 1).text == "200", doc.tables[0].cell(1, 1).text)


def t49_marks_formula_missing_cache_does_not_write_formula():
    from docx.enum.text import WD_COLOR_INDEX

    d = ws_dir("t49")
    docx = d / "标注无缓存.docx"
    doc = Document()
    run = doc.add_paragraph().add_run("待核对")
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.save(str(docx))
    api.run_export(docx)
    api.run_marks(docx)
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    wb[SHEET_MARKS]["E2"] = "=[1]Sheet1!$C$5"
    wb.save(str(xlsx))
    before = docx.read_bytes()
    try:
        api.run_marks_sync(docx)
        raise AssertionError("t49 无缓存外链公式应报错")
    except SyncError as exc:
        check("t49 标注无缓存明确提示", "Excel" in str(exc) and "保存" in str(exc), str(exc))
    check("t49 标注无缓存不写坏 Word", docx.read_bytes() == before)


def t50_combined_formula_missing_cache_stops_before_tables():
    d = ws_dir("t50")
    docx, xlsx = _prepare_formula_combo(d, "无缓存组合")
    wb = load_workbook(str(xlsx))
    wb.save(str(xlsx))
    before = docx.read_bytes()
    try:
        api.run_sync_content(docx, xlsx, tables=True, marks=True, placeholders=True)
        raise AssertionError("t50 组合公式无缓存应在写表格前报错")
    except SyncError as exc:
        check("t50 组合公式无缓存明确提示", "Excel" in str(exc) and "保存" in str(exc), str(exc))
    check("t50 组合公式无缓存不写任何内容", docx.read_bytes() == before)
