"""E：正文占位符（扫描/跨run替换/格式保留/页眉/表格）+ 正文信息骨架 + 模板生成集成。"""
from helpers import *  # noqa

from wesync.docx_placeholders import (  # noqa: E402
    VALUES_SHEET, apply_placeholders, read_placeholder_values, scan_placeholders, write_placeholder_sheet,
)

ROWS = [["项目", "金额"], ["A", "1"], ["合计", "1"]]


def _tpl_doc(path):
    doc = Document()
    p = doc.add_paragraph()
    r1 = p.add_run("我们审计了{公司")
    r1.font.bold = True
    p.add_run("名称}的财务报表。")
    doc.add_paragraph("报告编号：{报告文号}。")
    t = doc.add_table(rows=1, cols=1)
    t.cell(0, 0).text = "被审计单位：{公司名称}"
    doc.sections[0].header.paragraphs[0].text = "{公司名称}审计报告"
    doc.save(str(path))
    return path


def t17_placeholder_core():
    d = ws_dir("t17")
    path = _tpl_doc(d / "占位.docx")
    doc = Document(str(path))
    check("t17 扫描去重有序", scan_placeholders(doc) == ["公司名称", "报告文号"], str(scan_placeholders(doc)))
    values = {"公司名称": "甲公司", "报告文号": "审字〔2026〕1号"}
    plan = apply_placeholders(doc, values, dry_run=True)
    check("t17 预演列计划", len(plan["replaced"]) == 2 and not plan["missing"], str(plan))
    check("t17 预演不改文档", "{公司名称}" in doc.paragraphs[0].text)
    result = apply_placeholders(doc, values)
    check("t17 全部替换", len(result["replaced"]) == 3 and not result["missing"], str(result))
    check("t17 跨run替换", doc.paragraphs[0].text == "我们审计了甲公司的财务报表。", doc.paragraphs[0].text)
    check("t17 格式保留", doc.paragraphs[0].runs[0].font.bold is True)
    check("t17 表格内保留", doc.tables[0].cell(0, 0).text == "被审计单位：{公司名称}")
    check("t17 页眉替换", doc.sections[0].header.paragraphs[0].text == "甲公司审计报告")
    doc2 = Document(str(path))
    part = apply_placeholders(doc2, {"公司名称": "甲公司"})
    check("t17 缺值列出", part["missing"] == ["报告文号"], str(part))
    check("t17 缺值不替换", "{报告文号}" in doc2.paragraphs[1].text)


def t18_placeholder_skeleton():
    d = ws_dir("t18")
    tpl = _tpl_doc(d / "模板.docx")
    xlsx = d / "数据.xlsx"
    from openpyxl import Workbook
    Workbook().save(str(xlsx))
    fields = api.scan_template_placeholders(tpl)
    check("t18 api扫描", fields == ["公司名称", "报告文号"], str(fields))
    results = api.fill_placeholder_sheets(tpl, [xlsx])
    check("t18 建骨架", results[0]["status"] == "ok" and results[0]["added"] == fields, str(results))
    wb = load_workbook(str(xlsx))
    check("t18 表存在", VALUES_SHEET in wb.sheetnames)
    wb[VALUES_SHEET].cell(2, 2).value = "甲公司"
    wb.save(str(xlsx))
    added = write_placeholder_sheet(xlsx, fields + ["新字段"])
    check("t18 增量补齐", added == ["新字段"], str(added))
    check("t18 已填值保留", read_placeholder_values(xlsx)["公司名称"] == "甲公司")


def t19_template_batch_placeholders():
    d = ws_dir("t19")
    tpl = make_doc(d / "模板.docx", [("表1", ROWS)])
    doc = Document(str(tpl))
    doc.add_paragraph("被审计单位：{公司名称}，文号{报告文号}。")
    doc.save(str(tpl))
    api.run_export(tpl)
    xlsx = api.workbook_path(tpl)
    wb, order, blocks = blocks_of(xlsx)
    wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 4).value = 888
    wb.save(str(xlsx))
    api.fill_placeholder_sheets(tpl, [xlsx])
    wb = load_workbook(str(xlsx))
    wb[VALUES_SHEET].cell(2, 2).value = "甲公司"
    wb[VALUES_SHEET].cell(3, 2).value = "审字〔2026〕1号"
    wb.save(str(xlsx))
    tpl_bytes, xlsx_bytes = tpl.read_bytes(), xlsx.read_bytes()
    out = d / "out"
    results = api.run_template_batch([tpl], [{"excel": str(xlsx), "template": 1}], out, dry_run=True)
    check("t19 预演成功", results[0]["status"] == "ok" and not out.exists(), str(results))
    check("t19 预演列占位符", any("占位符·预演" in m and "甲公司" in m for m in LOGS), str(LOGS[-6:]))
    results = api.run_template_batch([tpl], [{"excel": str(xlsx), "template": 1}], out)
    check("t19 生成成功", results[0]["status"] == "ok", str(results))
    produced = Document(results[0]["output"])
    check("t19 正文已替换", "被审计单位：甲公司，文号审字〔2026〕1号。" in [p.text for p in produced.paragraphs],
          str([p.text for p in produced.paragraphs]))
    check("t19 表格已同步", word_tables(results[0]["output"])[0][1][1] == "888")
    check("t19 模板未动", tpl.read_bytes() == tpl_bytes)
    check("t19 源Excel未动", xlsx.read_bytes() == xlsx_bytes)
    xlsx2 = d / "缺值.xlsx"
    import shutil
    shutil.copy2(xlsx, xlsx2)
    wb2 = load_workbook(str(xlsx2))
    wb2[VALUES_SHEET].cell(3, 2).value = None
    wb2.save(str(xlsx2))
    results = api.run_template_batch([tpl], [{"excel": str(xlsx2), "template": 1}], out)
    check("t19 缺值跳过", results[0]["status"] == "skipped" and "占位符" in results[0]["reason"], str(results))

def t20_template_without_tables_placeholders_only():
    d = ws_dir("t20")
    tpl = d / "占位模板.docx"
    doc = Document()
    doc.add_paragraph("被审计单位：{公司名称}。")
    doc.save(str(tpl))
    api.run_export(tpl)
    xlsx = api.workbook_path(tpl)
    wb = load_workbook(str(xlsx))
    wb[VALUES_SHEET].cell(2, 2).value = "甲公司"
    wb.save(str(xlsx))
    results = api.run_template_batch([tpl], [{"excel": str(xlsx), "template": 1}], d / "out", contents=("placeholders",))
    check("t20 无表格模板不因书签跳过", results[0]["status"] == "ok", str(results))
    texts = [p.text for p in Document(results[0]["output"]).paragraphs]
    check("t20 占位符已替换", any("甲公司" in t for t in texts), str(texts))
