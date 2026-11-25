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
    api.run_export(docx)
    deleted = api.list_deleted_blocks(docx)
    check("t06 重导出保留已删块", len(deleted) == 1, str(deleted))
    api.run_restore_tables(docx, pick="1")
    check("t06 重导出后仍可恢复", len(word_tables(docx)) == 3)


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