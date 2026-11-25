"""D1：多报告批量同步 + 批量删空表。"""
from helpers import *  # noqa

ROWS = [["项目", "金额"], ["A", "1"], ["合计", "1"]]
BLANK = [["项目", "本期", "上期"], ["", "", ""], ["", "", ""]]


def t12_batch_writeback():
    d = ws_dir("t12")
    a = make_doc(d / "报告A.docx", [("表1", ROWS)])
    b = make_doc(d / "报告B.docx", [("表1", ROWS)])
    api.run_export(a)
    xlsx = api.workbook_path(a)
    wb, order, blocks = blocks_of(xlsx)
    wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 4).value = 999
    wb.save(str(xlsx))
    results = api.run_batch_writeback([a, b])
    check("t12 A成功B跳过", results[0]["status"] == "ok" and results[1]["status"] == "skipped", str(results))
    check("t12 A已写回999", word_tables(a)[0][1][1] == "999", str(word_tables(a)))
    check("t12 B未被改动", word_tables(b)[0][1][1] == "1")
    results = api.run_batch_writeback([a, d / "不存在.docx"], dry_run=True)
    check("t12 不存在文件跳过", results[1]["status"] == "skipped", str(results[1]))


def t13_blank_batch():
    d = ws_dir("t13")
    a = make_doc(d / "甲.docx", [("表1", ROWS), ("表2 空", BLANK)])
    b = make_doc(d / "乙.docx", [("表1", ROWS), ("表2 空", BLANK), ("表3 空", BLANK)])
    api.run_export(a)
    api.run_export(b)
    scans = api.scan_blank_batch([a, b])
    check("t13 扫描分组", len(scans[0]["candidates"]) == 1 and len(scans[1]["candidates"]) == 2, str(scans))
    items = [
        {"docx": str(a), "pick": "1", "keep_caption_for": None, "reversible": True},
        {"docx": str(b), "pick": "1,2", "keep_caption_for": "2", "reversible": True},
    ]
    results = api.run_blank_batch(items)
    check("t13 两份都成功", all(r["status"] == "ok" for r in results), str(results))
    check("t13 甲剩1张", len(word_tables(a)) == 1)
    check("t13 乙剩1张", len(word_tables(b)) == 1)
    docb = Document(str(b))
    check("t13 乙保留候选2标题", any(p.text.strip() == "表3 空" for p in docb.paragraphs))
    check("t13 可逆块保留", len(api.list_deleted_blocks(a)) == 1 and len(api.list_deleted_blocks(b)) == 2)
    api.run_restore_tables(b, pick="all")
    check("t13 乙全部恢复", len(word_tables(b)) == 3)