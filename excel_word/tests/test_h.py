"""H：B 列行标记（表头/✓/✕）、重新判定、删除✕行（可恢复）与旧布局兼容。"""
from helpers import *  # noqa
from wesync.config import MARK_DELETED, MARK_HEADER, MARK_NO, MARK_YES  # noqa

ROWS = [["项目", "金额", "备注"], ["甲", "100", ""], ["乙", "", ""],
        ["丙", "—", ""], ["合计", "100", ""]]


def _marks(ws, gr):
    return [str(ws.cell(r, 2).value or "").strip() for r in gr]


def t27_marks_and_rejudge():
    d = ws_dir("t27")
    docx = make_doc(d / "h.docx", [("表1 明细", ROWS)])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, order, blocks = blocks_of(xlsx)
    ws = wb[SHEET_NAME]
    gr = blocks["T01"]["grid_rows"]
    check("t27 导出标记", _marks(ws, gr) == [MARK_HEADER, MARK_YES, MARK_NO, MARK_YES, MARK_YES],
          str(_marks(ws, gr)))
    check("t27 新布局col0=3", blocks["T01"]["col0"] == 3, str(blocks["T01"]["col0"]))
    # 改内容后重新判定：乙填数→✓，甲清空→✕
    ws.cell(gr[2], 4).value = 5
    ws.cell(gr[1], 4).value = None
    wb.save(str(xlsx))
    before = xlsx.read_bytes()
    api.run_rejudge(docx, dry_run=True)
    check("t27 预演不写盘", xlsx.read_bytes() == before)
    api.run_rejudge(docx)
    wb, order, blocks = blocks_of(xlsx)
    ws = wb[SHEET_NAME]
    gr = blocks["T01"]["grid_rows"]
    check("t27 重判后标记", _marks(ws, gr) == [MARK_HEADER, MARK_NO, MARK_YES, MARK_YES, MARK_YES],
          str(_marks(ws, gr)))
    grid = grid_text(ws, blocks["T01"])
    check("t27 重判不动数据", grid[2][1] == "5" and grid[1][1] == "", str(grid))
    api.run_rejudge(docx)
    check("t27 重复重判幂等", any("无需写入" in m for m in LOGS), str(LOGS[-3:]))


def t28_delete_cross_rows():
    d = ws_dir("t28")
    rows = [["项目", "金额"], ["甲", "100"], ["乙", ""], ["丙", "200"], ["合计", "300"]]
    docx = make_doc(d / "h.docx", [("表1 明细", rows)])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, order, blocks = blocks_of(xlsx)
    gr = blocks["T01"]["grid_rows"]
    check("t28 乙为✕", str(wb[SHEET_NAME].cell(gr[2], 2).value) == MARK_NO)
    docx_bytes, xlsx_bytes = docx.read_bytes(), xlsx.read_bytes()
    api.run_writeback(docx, delete_empty_rows=True, dry_run=True)
    check("t28 预演不写Word", docx.read_bytes() == docx_bytes)
    check("t28 预演不写Excel", xlsx.read_bytes() == xlsx_bytes)
    api.run_writeback(docx, delete_empty_rows=True)
    t = word_tables(docx)[0]
    check("t28 Word删✕行", [r[0] for r in t] == ["项目", "甲", "丙", "合计"], str(t))
    wb, order, blocks = blocks_of(xlsx)
    ws = wb[SHEET_NAME]
    check("t28 Excel保留已删行", str(ws.cell(gr[2], 2).value) == MARK_DELETED
          and str(ws.cell(gr[2], 3).value) == "乙", str(ws.cell(gr[2], 2).value))
    check("t28 已删行不进网格", gr[2] not in blocks["T01"]["grid_rows"], str(blocks["T01"]["grid_rows"]))
    # 不勾选选项时已删行仍不同步（保持删除状态）
    api.run_writeback(docx)
    check("t28 默认不恢复", len(word_tables(docx)[0]) == 4)
    # 改回✓恢复
    ws.cell(gr[2], 2).value = MARK_YES
    wb.save(str(xlsx))
    api.run_writeback(docx)
    t = word_tables(docx)[0]
    check("t28 改回✓恢复行", [r[0] for r in t] == ["项目", "甲", "乙", "丙", "合计"], str(t))
    # 合计行豁免：整行空也不删
    d2 = ws_dir("t28b")
    docx2 = make_doc(d2 / "h.docx", [("表1", [["项目", "金额"], ["甲", ""], ["合计", ""]])])
    api.run_export(docx2)
    api.run_writeback(docx2, delete_empty_rows=True)
    t2 = word_tables(docx2)[0]
    check("t28 合计行豁免", [r[0] for r in t2] == ["项目", "合计"], str(t2))
    wb2, _, blocks2 = blocks_of(api.workbook_path(docx2))
    gr2 = blocks2["T01"]["grid_rows"]
    check("t28 合计标记仍✕", str(wb2[SHEET_NAME].cell(gr2[-1], 2).value) == MARK_NO,
          str(wb2[SHEET_NAME].cell(gr2[-1], 2).value))


def t30_delete_all_body_warn():
    d = ws_dir("t30")
    docx = make_doc(d / "h.docx", [("表1", [["项目", "金额"], ["甲", ""], ["乙", ""]])])
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    doc = Document(str(docx))
    doc.tables[0].rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    doc.save(str(docx))
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, order, blocks = blocks_of(xlsx)
    gr = blocks["T01"]["grid_rows"]
    LOGS.clear()
    api.run_writeback(docx, delete_empty_rows=True)
    check("t30 全✕表告警", any("覆盖全部数据行" in m for m in LOGS), str(LOGS[-6:]))
    check("t30 只剩表头", len(word_tables(docx)[0]) == 1, str(word_tables(docx)))
    wb, order, blocks = blocks_of(xlsx)
    ws = wb[SHEET_NAME]
    for r in gr[1:]:
        ws.cell(r, 2).value = MARK_YES
    wb.save(str(xlsx))
    LOGS.clear()
    api.run_writeback(docx)
    check("t30 恢复跳过有提示", any("跳过" in m and "复制格式" in m for m in LOGS), str(LOGS[-4:]))
    check("t30 跳过后不写坏", len(word_tables(docx)[0]) == 1)


def t29_old_layout_compat():
    d = ws_dir("t29")
    docx = make_doc(d / "h.docx", [("表1", [["项目", "金额"], ["甲", "1"], ["合计", "1"]])])
    api.run_export(docx, totals_formula=False)
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    wb[SHEET_NAME].delete_cols(2)  # 去掉标记列 → 旧布局
    wb.save(str(xlsx))
    wb, order, blocks = blocks_of(xlsx)
    check("t29 识别旧布局", blocks["T01"]["col0"] == 2, str(blocks["T01"]["col0"]))
    wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 3).value = 777
    wb.save(str(xlsx))
    api.run_writeback(docx, delete_empty_rows=True)
    check("t29 旧布局正常同步", word_tables(docx)[0][1][1] == "777", str(word_tables(docx)))
    before = xlsx.read_bytes()
    LOGS.clear()
    api.run_rejudge(docx)
    check("t29 旧布局重判提示", any("旧布局" in m for m in LOGS), str(LOGS[-3:]))
    check("t29 旧布局重判不写盘", xlsx.read_bytes() == before)