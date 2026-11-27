"""F：改动留痕清单导出 + 只读差异比对 + 备份列出与可逆回滚。"""
from helpers import *  # noqa
from wesync.config import SHEET_NEW_TABLES
from wesync.valuefmt import to_excel_value

ROWS = [["项目", "金额"], ["A", "1"], ["合计", "1"]]


def _doc_with_wb(d, name="报告.docx", value=888):
    docx = make_doc(d / name, [("表1", ROWS)])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, order, blocks = blocks_of(xlsx)
    wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 4).value = value
    wb.save(str(xlsx))
    return docx, xlsx


def t20_changelog_export():
    d = ws_dir("t20")
    docx, xlsx = _doc_with_wb(d)
    api.run_writeback(docx, changelog=True, backup=True)
    lists = list(d.glob("*_改动清单_*.xlsx"))
    check("t20 清单已生成", len(lists) == 1, str(lists))
    wb = load_workbook(str(lists[0]))
    ws = wb["改动清单"]
    check("t20 表头", [ws.cell(2, c).value for c in range(1, 9)] ==
          ["表号", "表名", "行", "列", "项目", "旧值", "新值", "类型"], str([ws.cell(2, c).value for c in range(1, 9)]))
    rows = [[ws.cell(r, c).value for c in range(1, 9)] for r in range(3, ws.max_row + 1)]
    cell_rows = [r for r in rows if r[7] == "单元格"]
    check("t20 含旧1新888", any(r[5] == "1" and r[6] == "888" for r in cell_rows), str(rows))
    check("t20 报告数据已同步", word_tables(docx)[0][1][1] == "888")


def t21_diff_readonly():
    d = ws_dir("t21")
    docx, xlsx = _doc_with_wb(d)
    docx_bytes, xlsx_bytes = docx.read_bytes(), xlsx.read_bytes()
    records = api.scan_writeback_diff(docx)
    cells = [r for r in records if r.get("kind") == "单元格"]
    check("t21 有差异记录", any(r["old"] == "1" and r["new"] == "888" for r in cells), str(records))
    check("t21 不写Word", docx.read_bytes() == docx_bytes)
    check("t21 不写Excel", xlsx.read_bytes() == xlsx_bytes)
    check("t21 无清单产出", not list(d.glob("*_改动清单_*.xlsx")))


def t22_backup_restore():
    d = ws_dir("t22")
    docx, xlsx = _doc_with_wb(d)
    api.run_writeback(docx, backup=True)
    check("t22 同步后为888", word_tables(docx)[0][1][1] == "888")
    info = api.scan_backups(docx)
    check("t22 Word有备份", len(info["word"]) >= 1, str(info["word"]))
    check("t22 分列含Excel键", "excel" in info and "word_path" in info)
    oldest = info["word"][-1]
    n_before = len(info["word"])
    api.run_restore_backup(oldest["path"], info["word_path"])
    check("t22 回滚后回到1", word_tables(docx)[0][1][1] == "1", str(word_tables(docx)))
    n_after = len(api.scan_backups(docx)["word"])
    check("t22 恢复前现状已再备份", n_after == n_before + 1, f"{n_before} -> {n_after}")


def t23_word_save_failure_rolls_back_excel(monkeypatch):
    import wesync.fsops as fsops
    from wesync.errors import SyncError

    d = ws_dir("t23")
    docx = make_doc(d / "报告.docx", [("表1", [
        ["项目", "金额"], ["A", "1"], ["B", "2"], ["合计", "3"],
    ])])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, _, blocks = blocks_of(xlsx)
    block = blocks["T01"]
    row_a = next(r for r in block["grid_rows"] if str(wb[SHEET_NAME].cell(r, block["col0"]).value) == "A")
    wb[SHEET_NAME].cell(row_a, 2).value = "✕"
    wb.save(str(xlsx))
    docx_before = docx.read_bytes()
    xlsx_before = xlsx.read_bytes()

    real_replace = fsops.os.replace
    failed = False

    def fail_word_once(src, dst):
        nonlocal failed
        if Path(dst) == docx and not failed:
            failed = True
            raise OSError("模拟 Word 最终替换失败")
        return real_replace(src, dst)

    monkeypatch.setattr(fsops.os, "replace", fail_word_once)
    try:
        api.run_writeback(docx, backup=False, delete_empty_rows=True)
        raise AssertionError("t23 应模拟保存失败")
    except SyncError as exc:
        check("t23 明确提示已恢复", "已恢复" in str(exc), str(exc))

    check("t23 Word保持原字节", docx.read_bytes() == docx_before)
    check("t23 Excel保持原字节", xlsx.read_bytes() == xlsx_before)
    wb_after = load_workbook(str(xlsx))
    check("t23 Excel未留下已删墓碑", wb_after[SHEET_NAME].cell(row_a, 2).value == "✕",
          str(wb_after[SHEET_NAME].cell(row_a, 2).value))


def t24_export_word_failure_rolls_back_excel(monkeypatch):
    import wesync.fsops as fsops
    from wesync.errors import SyncError

    d = ws_dir("t24")
    docx = make_doc(d / "报告.docx", [("表1", ROWS)])
    api.run_export(docx)
    api.run_clear_bookmarks(docx)
    xlsx = api.workbook_path(docx)
    wb, _, blocks = blocks_of(xlsx)
    block = blocks["T01"]
    row_a = next(r for r in block["grid_rows"] if str(wb[SHEET_NAME].cell(r, block["col0"]).value) == "A")
    wb[SHEET_NAME].cell(row_a, 2).value = "999"
    wb.save(str(xlsx))
    docx_before = docx.read_bytes()
    xlsx_before = xlsx.read_bytes()

    real_replace = fsops.os.replace

    def fail_word_once(src, dst):
        if Path(dst) == docx:
            raise OSError("模拟导出 Word 最终替换失败")
        return real_replace(src, dst)

    monkeypatch.setattr(fsops.os, "replace", fail_word_once)
    try:
        api.run_export(docx, backup=False)
        raise AssertionError("t24 应模拟保存失败")
    except SyncError as exc:
        check("t24 明确提示已恢复", "已恢复" in str(exc), str(exc))

    check("t24 Word保持原字节", docx.read_bytes() == docx_before)
    check("t24 Excel保持原字节", xlsx.read_bytes() == xlsx_before)
    check("t24 无残留暂存文件", not list(d.glob(".报告.*")))


def t25_first_export_failure_removes_new_excel(monkeypatch):
    import wesync.fsops as fsops
    from wesync.errors import SyncError

    d = ws_dir("t25")
    docx = make_doc(d / "报告.docx", [("表1", ROWS)])
    xlsx = api.workbook_path(docx)
    docx_before = docx.read_bytes()
    real_replace = fsops.os.replace

    def fail_word_once(src, dst):
        if Path(dst) == docx:
            raise OSError("模拟首次导出 Word 最终替换失败")
        return real_replace(src, dst)

    monkeypatch.setattr(fsops.os, "replace", fail_word_once)
    try:
        api.run_export(docx, backup=False)
        raise AssertionError("t25 应模拟保存失败")
    except SyncError as exc:
        check("t25 明确提示已恢复", "已恢复" in str(exc), str(exc))

    check("t25 Word保持原字节", docx.read_bytes() == docx_before)
    check("t25 不留下半成品Excel", not xlsx.exists())
    check("t25 无残留暂存文件", not list(d.glob(".报告.*")))


def t26_number_text_conversion():
    cases = [
        ("1,2", ("1,2", None)),
        ("12,34,567", ("12,34,567", None)),
        ("1,2345", ("1,2345", None)),
        ("0,123", ("0,123", None)),
        ("00,000", ("00,000", None)),
        ("00123", ("00123", None)),
        ("-00123", ("-00123", None)),
        ("1234567890123456", ("1234567890123456", None)),
        ("12345678901234567890", ("12345678901234567890", None)),
        ("123456789012345", (123456789012345, "#,##0")),
        ("1,234", (1234, "#,##0")),
        ("12,345,678.90", (12345678.9, "#,##0.00")),
        ("0.50", (0.5, "#,##0.00")),
    ]
    for text, expected in cases:
        check(f"t26 数字文本 {text}", to_excel_value(text) == expected, str(to_excel_value(text)))


def t27_number_identifiers_full_incremental_roundtrip():
    rows = [
        ["项目", "编号", "金额"],
        ["甲", "00123", "1,234"],
        ["乙", "12345678901234567890", "1,2"],
        ["丙", "00077", "12,345,678.90"],
        ["合计", "", "12,346,912.90"],
    ]
    d = ws_dir("t27")
    docx = make_doc(d / "数字文本.docx", [("表1", rows)])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, _, blocks = blocks_of(xlsx)
    block = blocks["T01"]
    gr = block["grid_rows"]
    ws = wb[SHEET_NAME]
    check("t27 全量前导零保留文本", ws.cell(gr[1], 4).value == "00123" and ws.cell(gr[1], 4).number_format == "@")
    check("t27 全量长整数保留文本", ws.cell(gr[2], 4).value == rows[2][1] and ws.cell(gr[2], 4).number_format == "@")
    check("t27 全量错误千分位保留文本", ws.cell(gr[2], 5).value == "1,2" and ws.cell(gr[2], 5).number_format == "@")
    check("t27 全量合法千分位转数值", ws.cell(gr[1], 5).value == 1234)
    check("t27 全量合计仍为公式", str(ws.cell(gr[-1], 5).value).startswith("=SUM("), str(ws.cell(gr[-1], 5).value))
    api.run_writeback(docx)
    check("t27 全量导出回写原文", word_tables(docx)[0] == rows, str(word_tables(docx)[0]))

    xt_rows = [
        ["项目", "编号", "金额"],
        ["丁", "00088", "9,876"],
        ["戊", "98765432109876543210", "2,3"],
        ["己", "00123", "6,543.21"],
        ["合计", "", "16,419.21"],
    ]
    doc = Document(str(docx))
    table = doc.add_table(rows=len(xt_rows), cols=3)
    table.style = "Table Grid"
    for r, row in enumerate(xt_rows):
        for c, text in enumerate(row):
            table.cell(r, c).text = text
    doc.save(str(docx))
    api.run_incremental_export(docx)
    wb = load_workbook(str(xlsx))
    _, xt_blocks = parse_blocks(wb[SHEET_NEW_TABLES])
    xt_block = xt_blocks["XT01"]
    xgr = xt_block["grid_rows"]
    ws = wb[SHEET_NEW_TABLES]
    check("t27 增量前导零保留文本", ws.cell(xgr[1], 4).value == "00088" and ws.cell(xgr[1], 4).number_format == "@")
    check("t27 增量长整数保留文本", ws.cell(xgr[2], 4).value == xt_rows[2][1] and ws.cell(xgr[2], 4).number_format == "@")
    check("t27 增量错误千分位保留文本", ws.cell(xgr[2], 5).value == "2,3" and ws.cell(xgr[2], 5).number_format == "@")
    api.run_writeback(docx)
    restored = next(t for t in word_tables(docx) if t[0] == xt_rows[0] and t[1] == xt_rows[1])
    check("t27 增量导出回写原文", restored == xt_rows, str(restored))
