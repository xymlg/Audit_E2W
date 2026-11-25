"""F：改动留痕清单导出 + 只读差异比对 + 备份列出与可逆回滚。"""
from helpers import *  # noqa

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
