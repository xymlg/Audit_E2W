"""D2：模板生成（单/多模板、错配跳过、预演、重名）+ 排序合并区保护。"""
from helpers import *  # noqa

ROWS = [["项目", "金额"], ["A", "1"], ["合计", "1"]]


def _tpl_with_wb(d, name):
    tpl = make_doc(d / name, [("表1", ROWS)])
    api.run_export(tpl)
    return tpl, api.workbook_path(tpl)


def _edit_wb(xlsx, value):
    wb, order, blocks = blocks_of(xlsx)
    wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 4).value = value
    wb.save(str(xlsx))


def t14_template_single():
    d = ws_dir("t14")
    tpl, xlsx = _tpl_with_wb(d, "模板.docx")
    _edit_wb(xlsx, 888)
    tpl_bytes, xlsx_bytes = tpl.read_bytes(), xlsx.read_bytes()
    out = d / "out"
    results = api.run_template_batch([tpl], [{"excel": str(xlsx), "template": 1}], out, dry_run=True)
    check("t14 预演不产出", results[0]["status"] == "ok" and not out.exists(), str(results))
    results = api.run_template_batch([tpl], [{"excel": str(xlsx), "template": 1}], out)
    check("t14 生成成功", results[0]["status"] == "ok", str(results))
    produced = Path(results[0]["output"])
    check("t14 命名规则", produced.name == "模板_模板_表格数据.docx", produced.name)
    check("t14 新Word含888", word_tables(produced)[0][1][1] == "888", str(word_tables(produced)))
    check("t14 模板原件未动", tpl.read_bytes() == tpl_bytes)
    check("t14 源Excel未动", xlsx.read_bytes() == xlsx_bytes)
    results = api.run_template_batch([tpl], [{"excel": str(xlsx), "template": 1}], out)
    check("t14 重名加_2", Path(results[0]["output"]).name.endswith("_2.docx"), results[0]["output"])


def t15_template_multi():
    d = ws_dir("t15")
    tpl1, wb1 = _tpl_with_wb(d, "模板一.docx")
    tpl2, wb2 = _tpl_with_wb(d, "模板二.docx")
    _edit_wb(wb1, 111)
    _edit_wb(wb2, 222)
    out = d / "out"
    items = [
        {"excel": str(wb1), "template": 1},
        {"excel": str(wb2), "template": 2},
        {"excel": str(wb1), "template": 2},
        {"excel": str(wb2), "template": 0},
    ]
    results = api.run_template_batch([tpl1, tpl2], items, out)
    check("t15 分组生成", results[0]["status"] == "ok" and results[1]["status"] == "ok", str(results))
    check("t15 各自内容正确", word_tables(Path(results[0]["output"]))[0][1][1] == "111"
          and word_tables(Path(results[1]["output"]))[0][1][1] == "222")
    check("t15 错配跳过", results[2]["status"] == "skipped" and "书签" in results[2]["reason"], str(results[2]))
    check("t15 未指定模板跳过", results[3]["status"] == "skipped", str(results[3]))


def t16_sort_merge_guard():
    d = ws_dir("t16")
    rows = [["项目", "金额"], ["甲", "300"], ["乙", "100"], ["合计", "400"]]
    docx = make_doc(d / "m.docx", [("表1", rows)])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, order, blocks = blocks_of(xlsx)
    ws = wb[SHEET_NAME]
    gr = blocks["T01"]["grid_rows"]
    ws.merge_cells(start_row=gr[1], start_column=3, end_row=gr[2], end_column=3)
    wb.save(str(xlsx))
    api.set_table_sort(docx, code="T01", column="金额", descending=True)
    LOGS.clear()
    api.run_writeback(docx)
    check("t16 合并区跳过排序", any("合并单元格" in m for m in LOGS), str(LOGS[-5:]))
    wb, order, blocks = blocks_of(xlsx)
    items = [r[0] for r in grid_text(wb[SHEET_NAME], blocks["T01"])]
    check("t16 顺序未变", items == ["项目", "甲", "", "合计"], str(items))