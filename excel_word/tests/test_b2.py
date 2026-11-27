"""B2：空白表可逆删除/恢复、保留标题、锚点消失恢复、重导出保护。"""
from helpers import *  # noqa

BLANK = [["项目", "本期", "上期"], ["", "", ""], ["", "", ""]]
FULL1 = [["项目", "金额"], ["A", "1"], ["合计", "1"]]
FULL3 = [["项目", "金额"], ["C", "3"], ["合计", "3"]]


def _doc(d, name="b.docx"):
    return make_doc(d / name, [("表1 数据一", FULL1), ("表2 空白披露", BLANK), ("表3 数据三", FULL3)])


def t05_blank_roundtrip():
    d = ws_dir("t05")
    docx = _doc(d)
    api.run_export(docx)
    cands = api.scan_blank_tables(docx)
    check("t05 扫描只读识别1张", len(cands) == 1 and cands[0]["caption_text"] == "表2 空白披露", str(cands))
    api.run_blank_tables(docx, pick="1", reversible=True)
    check("t05 删后剩2张", len(word_tables(docx)) == 2)
    deleted = api.list_deleted_blocks(docx)
    check("t05 已删除标记+原位置", len(deleted) == 1 and deleted[0]["restore_after"], str(deleted))
    wb, _, blocks = blocks_of(api.workbook_path(docx))
    ws = wb[SHEET_NAME]
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 36
    ws.column_dimensions["E"].width = 24
    wb.save(str(api.workbook_path(docx)))
    ar = blocks["T02"]["anchor_row"]
    check(
        "t05 锚点表名/列数保留且链接清除",
        str(ws.cell(ar, 3).value).startswith("【表")
        and str(ws.cell(ar, 6).value).startswith("列数：")
        and ws.cell(ar, 5).value is None
        and ws.cell(ar, 5).hyperlink is None,
        str([ws.cell(ar, c).value for c in range(3, 9)]),
    )
    api.run_writeback(docx)
    check("t05 同步跳过已删表不重建", len(word_tables(docx)) == 2)
    api.run_restore_tables(docx, pick="1")
    ts = word_tables(docx)
    check("t05 恢复后3张且原位", len(ts) == 3 and ts[1][0] == ["项目", "本期", "上期"], str([t[0] for t in ts]))
    doc = Document(str(docx))
    check("t05 标题一起恢复", any(p.text.strip() == "表2 空白披露" for p in doc.paragraphs))
    from docx.oxml.ns import qn
    from wesync.docx_newtable import available_width_dxa, col_widths_from_excel
    restored_table = next(t for t in doc.tables if [c.text for c in t.rows[0].cells] == BLANK[0])
    actual_widths = [int(col.get(qn("w:w"))) for col in restored_table._tbl.tblGrid.gridCol_lst]
    expected_widths = col_widths_from_excel(ws, 3, available_width_dxa(doc), col0=3)
    check("t05 恢复列宽按CDE", actual_widths == expected_widths, f"{actual_widths} != {expected_widths}")
    check("t05 恢复后无已删除块", api.list_deleted_blocks(docx) == [])
    api.run_writeback(docx)
    check("t05 恢复表可正常同步", len(word_tables(docx)) == 3)


def t06_keep_caption_and_reexport():
    d = ws_dir("t06")
    docx = _doc(d)
    api.run_export(docx)
    api.run_blank_tables(docx, pick="1", keep_caption_for="1", reversible=True)
    doc = Document(str(docx))
    check("t06 保留标题", any(p.text.strip() == "表2 空白披露" for p in doc.paragraphs))
    wb, _, blocks = blocks_of(api.workbook_path(docx))
    old_block = blocks["T02"]
    old_rows = list(old_block["grid_rows"])
    wb[SHEET_NAME].cell(old_rows[1], 5).value = f"=SUM(D{old_rows[1]}:D{old_rows[2]})"
    wb.save(str(api.workbook_path(docx)))
    api.run_export(docx)
    deleted = api.list_deleted_blocks(docx)
    check("t06 重导出保留已删块", len(deleted) == 1, str(deleted))
    wb, _, blocks = blocks_of(api.workbook_path(docx))
    new_block = blocks[deleted[0]["code"]]
    expected_formula = f"=SUM(D{new_block['grid_rows'][1]}:D{new_block['grid_rows'][2]})"
    actual_formula = wb[SHEET_NAME].cell(new_block["grid_rows"][1], 5).value
    check("t06 删除块搬迁公式改行号", new_block["grid_rows"] != old_rows and actual_formula == expected_formula,
          f"{actual_formula} != {expected_formula}")
    api.run_restore_tables(docx, pick="1")
    check("t06 重导出后仍可恢复", len(word_tables(docx)) == 3)


def t09_restore_all_uses_each_block_sheet():
    d = ws_dir("t06c")
    docx = make_doc(d / "跨表恢复.docx", [("表1 空", BLANK)])
    api.run_export(docx)
    doc = Document(str(docx))
    table = doc.add_table(rows=len(BLANK), cols=len(BLANK[0]))
    table.style = "Table Grid"
    for r, row in enumerate(BLANK):
        for c, text in enumerate(row):
            table.cell(r, c).text = text
    doc.save(str(docx))
    api.run_incremental_export(docx)
    api.run_blank_tables(docx, pick="all", reversible=True)
    check("t06c T和XT均可逆删除", [b["code"] for b in api.list_deleted_blocks(docx)] == ["T01", "XT01"], str(api.list_deleted_blocks(docx)))
    api.run_restore_tables(docx, pick="all")

    check("t06c 跨表全部恢复", len(word_tables(docx)) == 2 and api.list_deleted_blocks(docx) == [], str(word_tables(docx)))
    wb = load_workbook(str(api.workbook_path(docx)))
    _, t_blocks = parse_blocks(wb[SHEET_NAME])
    _, xt_blocks = parse_blocks(wb["新增表格"])
    names = {t_blocks["T01"]["bookmark"], xt_blocks["XT01"]["bookmark"]}
    check("t06c 两表元数据各写本表", len(names) == 2 and all(name.startswith("Table_") for name in names), str(names))


def t10_reversible_transaction_rollback(monkeypatch):
    import wesync.fsops as fsops
    from wesync.errors import SyncError

    d = ws_dir("t05b")
    docx = _doc(d)
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    docx_before, xlsx_before = docx.read_bytes(), xlsx.read_bytes()
    real_replace = fsops.os.replace

    def fail_excel(src, dst):
        if Path(dst) == xlsx:
            raise OSError("模拟可逆删除 Excel 替换失败")
        return real_replace(src, dst)

    monkeypatch.setattr(fsops.os, "replace", fail_excel)
    try:
        api.run_blank_tables(docx, pick="1", reversible=True)
        raise AssertionError("t05b 可逆删除应模拟保存失败")
    except SyncError as exc:
        check("t05b 删除失败明确报错", "保存" in str(exc), str(exc))
    check("t05b 删除失败Word字节不变", docx.read_bytes() == docx_before)
    check("t05b 删除失败Excel字节不变", xlsx.read_bytes() == xlsx_before)
    check("t05b 删除失败表格仍在", len(word_tables(docx)) == 3 and api.list_deleted_blocks(docx) == [])

    monkeypatch.setattr(fsops.os, "replace", real_replace)
    api.run_blank_tables(docx, pick="1", reversible=True)
    deleted_docx, deleted_xlsx = docx.read_bytes(), xlsx.read_bytes()

    def fail_word(src, dst):
        if Path(dst) == docx:
            raise OSError("模拟恢复 Word 最终替换失败")
        return real_replace(src, dst)

    monkeypatch.setattr(fsops.os, "replace", fail_word)
    try:
        api.run_restore_tables(docx, pick="1")
        raise AssertionError("t05b 恢复应模拟保存失败")
    except SyncError as exc:
        check("t05b 恢复失败已回滚Excel", "已恢复" in str(exc), str(exc))
    check("t05b 恢复失败Word字节不变", docx.read_bytes() == deleted_docx)
    check("t05b 恢复失败Excel字节不变", xlsx.read_bytes() == deleted_xlsx)
    check("t05b 恢复失败状态不变", len(word_tables(docx)) == 2 and len(api.list_deleted_blocks(docx)) == 1)
    check("t05b 无残留暂存文件", not list(d.glob(".b.*")))

    monkeypatch.setattr(fsops.os, "replace", real_replace)
    api.run_restore_tables(docx, pick="1")
    check("t05b 重试恢复成功", len(word_tables(docx)) == 3 and api.list_deleted_blocks(docx) == [])


def t07_restore_anchor_gone():
    d = ws_dir("t07")
    docx = _doc(d)
    api.run_export(docx)
    api.run_blank_tables(docx, pick="1", reversible=True)
    doc = Document(str(docx))
    t0 = doc.tables[0]._element
    t0.getparent().remove(t0)
    doc.save(str(docx))
    LOGS.clear()
    api.run_restore_tables(docx, pick="1")
    ts = word_tables(docx)
    check("t07 锚点消失恢复到末尾", len(ts) == 2 and ts[-1][0] == ["项目", "本期", "上期"], str([t[0] for t in ts]))
    check("t07 有提醒日志", any("末尾" in m or "原位置" in m for m in LOGS), str(LOGS[-4:]))


def t08_plain_delete():
    d = ws_dir("t08")
    docx = _doc(d)
    api.run_export(docx)
    api.run_blank_tables(docx, pick="1", reversible=False)
    check("t08 普通删除剩2张", len(word_tables(docx)) == 2)
    check("t08 无已删除块", api.list_deleted_blocks(docx) == [])
    api.run_export(docx)
    wb, order, blocks = blocks_of(api.workbook_path(docx))
    check("t08 重导出后2块", len(order) == 2, str(order))