"""增量导出：XT 表格、标注与占位符只追加，既有行与公式位置保持不变。"""
from helpers import *  # noqa
from docx.enum.text import WD_COLOR_INDEX
from wesync.config import SHEET_NAME, SHEET_NEW_TABLES
from wesync.docx_placeholders import VALUES_SHEET
from wesync.errors import SyncError
from wesync.excel_marks import SHEET_MARKS
from wesync.valuefmt import formula_map


def _add_table_at_start(doc, rows):
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = "Table Grid"
    for i, row in enumerate(rows):
        for j, text in enumerate(row):
            table.cell(i, j).text = str(text)
    doc.tables[0]._tbl.addprevious(table._tbl)
    return table


def t04_incremental_export():
    d = ws_dir("t04")
    docx = make_doc(d / "n.docx", [("表1 甲", [["项目", "金额"], ["A", "1"], ["合计", "1"]])])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, order, blocks = blocks_of(xlsx)
    old_block = blocks["T01"]
    old_rows = list(old_block["grid_rows"])
    old_ws = wb[SHEET_NAME]
    old_values = [[old_ws.cell(r, c).value for c in range(1, 7)] for r in range(1, old_ws.max_row + 1)]
    old_ws.cell(old_rows[-1], 4).value = f"=SUM(D{old_rows[1]}:D{old_rows[-2]})"
    old_values = [[old_ws.cell(r, c).value for c in range(1, 7)] for r in range(1, old_ws.max_row + 1)]
    wb.save(str(xlsx))

    doc = Document(str(docx))
    table = _add_table_at_start(doc, [["项目", "本期"], ["新增甲", "100"], ["合计", "100"]])
    table.cell(1, 1).text = "{表内字段}"
    highlight = table.cell(1, 1).paragraphs[0].runs[0]
    highlight.font.highlight_color = WD_COLOR_INDEX.YELLOW
    run = doc.add_paragraph().add_run("新增标注")
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    doc.add_paragraph("单位：{新增单位}。")
    doc.save(str(docx))

    before_docx, before_xlsx = docx.read_bytes(), xlsx.read_bytes()
    api.run_incremental_export(docx, dry_run=True)
    check("t04 预演不写Word", docx.read_bytes() == before_docx)
    check("t04 预演不写Excel", xlsx.read_bytes() == before_xlsx)
    api.run_incremental_export(docx)
    wb = load_workbook(str(xlsx))
    check("t04 新建新增表格sheet", SHEET_NEW_TABLES in wb.sheetnames, str(wb.sheetnames))
    xt_order, xt_blocks = parse_blocks(wb[SHEET_NEW_TABLES])
    check("t04 首增表XT01", xt_order == ["XT01"], str(xt_order))
    check("t04 表块有真实书签", xt_blocks["XT01"]["bookmark"].startswith("Table_"), str(xt_blocks["XT01"]))
    t_order, t_blocks = parse_blocks(wb[SHEET_NAME])
    check("t04 原表编号不变", t_order == ["T01"], str(t_order))
    check("t04 原表行号不变", list(t_blocks["T01"]["grid_rows"]) == old_rows)
    check("t04 原表公式不变", wb[SHEET_NAME].cell(old_rows[-1], 4).value == old_values[-1][3])
    check("t04 原表内容不变", [[wb[SHEET_NAME].cell(r, c).value for c in range(1, 7)] for r in range(1, len(old_values) + 1)] == old_values)
    marks_ws = wb[SHEET_MARKS]
    mark_rows = [r for r in range(2, marks_ws.max_row + 1) if marks_ws.cell(r, 4).value == "新增标注"]
    check("t04 新增标注XM01", len(mark_rows) == 1 and marks_ws.cell(mark_rows[0], 1).value == "XM01", str(mark_rows))
    check("t04 标注来源增量", marks_ws.cell(mark_rows[0], 8).value == "增量")
    check("t04 表内标注未导出", all("表内字段" != str(marks_ws.cell(r, 4).value) for r in range(2, marks_ws.max_row + 1)))
    ph_ws = wb[VALUES_SHEET]
    fields = {str(ph_ws.cell(r, 1).value): r for r in range(2, ph_ws.max_row + 1)}
    check("t04 新占位符追加", "新增单位" in fields and "表内字段" not in fields, str(fields))
    check("t04 占位符来源增量", ph_ws.cell(fields["新增单位"], 3).value == "增量")
    wb.save(str(xlsx))

    doc = Document(str(docx))
    _add_table_at_start(doc, [["项目", "上期"], ["新增乙", "200"], ["合计", "200"]])
    doc.save(str(docx))
    api.run_incremental_export(docx)
    wb = load_workbook(str(xlsx))
    xt_order, xt_blocks = parse_blocks(wb[SHEET_NEW_TABLES])
    check("t04 第二次追加XT02", xt_order == ["XT01", "XT02"], str(xt_order))
    check("t04 原XT01行号不变", xt_blocks["XT01"]["grid_rows"] == parse_blocks(load_workbook(str(xlsx))[SHEET_NEW_TABLES])[1]["XT01"]["grid_rows"])

    unchanged_docx, unchanged_xlsx = docx.read_bytes(), xlsx.read_bytes()
    api.run_incremental_export(docx)
    check("t04 无增量不改Word", docx.read_bytes() == unchanged_docx)
    check("t04 无增量不改Excel", xlsx.read_bytes() == unchanged_xlsx)

    wb = load_workbook(str(xlsx))
    block = parse_blocks(wb[SHEET_NEW_TABLES])[1]["XT02"]
    wb[SHEET_NEW_TABLES].cell(block["grid_rows"][1], block["col0"] + 1).value = 999
    wb.save(str(xlsx))
    api.run_writeback(docx)
    check("t04 XT表参与回写", any(row[1] == ["新增乙", "999"] for row in word_tables(docx)), str(word_tables(docx)))
    api.run_bookmark(docx)
    doc = Document(str(docx))
    from wesync.docx_bookmarks import bookmark_names
    names = bookmark_names(doc.element.body)
    check("t04 书签维护保留XT", all(block["bookmark"] in names for block in xt_blocks.values()))
    wb = load_workbook(str(xlsx))
    xt_name = wb.defined_names.get("EX_XT01")
    expected_ref = f"'{SHEET_NEW_TABLES}'!$C${xt_blocks['XT01']['anchor_row'] + 1}"
    check("t04 XT定义名称指向新增表格", xt_name is not None and xt_name.attr_text == expected_ref, str(xt_name))


def t53_xt_sort_batch_blank_restore_template():
    d = ws_dir("t04b")
    docx = make_doc(d / "排序.docx", [("表1 原表", [["项目", "金额"], ["甲", "10"]])])
    api.run_export(docx)
    doc = Document(str(docx))
    _add_table_at_start(doc, [["项目", "金额"], ["乙", "30"], ["甲", "20"], ["合计", "50"]])
    doc.save(str(docx))
    api.run_incremental_export(docx)
    api.set_table_sort(docx, code="XT01", column="金额", descending=True)
    api.run_writeback(docx)
    check("t04b XT排序回写", any([r[0] for r in table] == ["项目", "乙", "甲", "合计"] for table in word_tables(docx)), str(word_tables(docx)))

    batch_docx = make_doc(d / "批量.docx", [("表1 原表", [["项目", "金额"], ["甲", "1"]])])
    api.run_export(batch_docx)
    doc = Document(str(batch_docx))
    _add_table_at_start(doc, [["项目", "金额"], ["乙", "2"]])
    doc.save(str(batch_docx))
    api.run_incremental_export(batch_docx)
    xlsx = api.workbook_path(batch_docx)
    wb = load_workbook(str(xlsx))
    block = parse_blocks(wb[SHEET_NEW_TABLES])[1]["XT01"]
    wb[SHEET_NEW_TABLES].cell(block["grid_rows"][1], block["col0"] + 1).value = 222
    wb.save(str(xlsx))
    results = api.run_batch_writeback([batch_docx])
    check("t04b 批量同步XT", results[0]["status"] == "ok" and any(row[1] == ["乙", "222"] for row in word_tables(batch_docx)), str(results))

    blank_docx = make_doc(d / "空白.docx", [("表1 原表", [["项目", "金额"], ["甲", "1"]])])
    api.run_export(blank_docx)
    doc = Document(str(blank_docx))
    _add_table_at_start(doc, [["项目", "金额"], ["", ""]])
    doc.save(str(blank_docx))
    api.run_incremental_export(blank_docx)
    api.run_blank_tables(blank_docx, pick="1", reversible=True, keep_caption=True)
    check("t04b XT可逆删除", api.list_deleted_blocks(blank_docx)[0]["code"] == "XT01", str(api.list_deleted_blocks(blank_docx)))
    api.run_restore_tables(blank_docx, pick="1")
    check("t04b XT恢复", not api.list_deleted_blocks(blank_docx) and len(word_tables(blank_docx)) == 2, str(word_tables(blank_docx)))

    template = make_doc(d / "模板.docx", [("表1 原表", [["项目", "金额"], ["甲", "1"]])])
    api.run_export(template)
    doc = Document(str(template))
    _add_table_at_start(doc, [["项目", "金额"], ["乙", "2"]])
    doc.save(str(template))
    api.run_incremental_export(template)
    xlsx = api.workbook_path(template)
    wb = load_workbook(str(xlsx))
    block = parse_blocks(wb[SHEET_NEW_TABLES])[1]["XT01"]
    wb[SHEET_NEW_TABLES].cell(block["grid_rows"][1], block["col0"] + 1).value = 333
    wb.save(str(xlsx))
    results = api.run_template_batch([template], [{"excel": str(xlsx), "template": 1}], d / "out", contents=("tables",))
    check("t04b 模板同步XT", results[0]["status"] == "ok", str(results))
    check("t04b 模板XT值", any(row[1] == ["乙", "333"] for row in word_tables(results[0]["output"])), str(results))


def t51_formula_cache_is_sheet_scoped():
    from wesync.commands.writeback_cmd import _new_table_grid
    from wesync.sort_rules import _amount

    d = ws_dir("t04a")
    docx = make_doc(d / "公式.docx", [("表1 原表", [["项目", "金额"], ["甲", "10"], ["乙", "20"], ["合计", "30"]])])
    api.run_export(docx)
    doc = Document(str(docx))
    table = doc.add_table(rows=4, cols=2)
    table.style = "Table Grid"
    for r, row in enumerate([["项目", "金额"], ["新增甲", "7"], ["新增乙", "3"], ["合计", "10"]]):
        for c, text in enumerate(row):
            table.cell(r, c).text = text
    doc.tables[0]._tbl.addprevious(table._tbl)
    doc.save(str(docx))
    api.run_incremental_export(docx)

    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    _, t_blocks = parse_blocks(wb[SHEET_NAME])
    _, xt_blocks = parse_blocks(wb[SHEET_NEW_TABLES])
    t_row = t_blocks["T01"]["grid_rows"][1]
    xt_block = xt_blocks["XT01"]
    xt_row = xt_block["grid_rows"][1]
    wb[SHEET_NAME].cell(t_row, 4).value = 999
    wb[SHEET_NEW_TABLES].cell(xt_row, 4).value = "=5+5"
    wb.save(str(xlsx))

    cached = formula_map(xlsx)
    check("t04a 缓存键包含工作表", all(len(key) == 3 for key in cached), str(list(cached)[:3]))
    check("t04a T缓存保留", cached.get((SHEET_NAME, t_row, 4)) == 999, str(cached.get((SHEET_NAME, t_row, 4))))
    check("t04a XT同坐标不串表", (SHEET_NEW_TABLES, xt_row, 4) not in cached, str(cached.get((SHEET_NEW_TABLES, xt_row, 4))))
    xt_ws = wb[SHEET_NEW_TABLES]
    check("t04a XT排序公式独立求值", _amount(xt_ws, xt_row, 4, cached) == 10.0, str(_amount(xt_ws, xt_row, 4, cached)))
    check("t04a XT新表网格独立求值", _new_table_grid(xt_ws, xt_block, cached)[1][1] == "10", str(_new_table_grid(xt_ws, xt_block, cached)))
    wb.close()

    api.run_writeback(docx)
    tables = word_tables(docx)
    check("t04a XT回写不取T缓存", tables[0][1][1] == "10", str(tables))
    check("t04a T回写仍取本表值", tables[1][1][1] == "999", str(tables))


def t52_incremental_requires_summary_sheet():
    d = ws_dir("t04c")
    docx = make_doc(d / "异常工作簿.docx", [("表1", [["项目", "金额"], ["甲", "1"]])])
    api.run_export(docx)
    for name in ("新增一", "新增二"):
        doc = Document(str(docx))
        table = doc.add_table(rows=2, cols=2)
        table.style = "Table Grid"
        table.cell(0, 0).text, table.cell(0, 1).text = "项目", "金额"
        table.cell(1, 0).text, table.cell(1, 1).text = name, "2"
        doc.save(str(docx))
        if name == "新增一":
            api.run_incremental_export(docx)

    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    wb.remove(wb[SHEET_NAME])
    wb.active = wb.sheetnames.index(SHEET_NEW_TABLES)
    wb.save(str(xlsx))
    wb.close()
    docx_before, xlsx_before = docx.read_bytes(), xlsx.read_bytes()
    for dry_run in (True, False):
        try:
            api.run_incremental_export(docx, dry_run=dry_run)
            raise AssertionError("t04c 缺少表格汇总应拒绝增量导出")
        except SyncError as exc:
            check("t04c 明确缺少表格汇总", "表格汇总" in str(exc), str(exc))
        check("t04c Word保持原字节", docx.read_bytes() == docx_before)
        check("t04c Excel保持原字节", xlsx.read_bytes() == xlsx_before)
