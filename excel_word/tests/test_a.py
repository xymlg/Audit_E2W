"""A组：导出、改值回写、增删行列。"""
from helpers import *  # noqa
from wesync.errors import SyncError


def t01_export():
    d = ws_dir("t01")
    docx = copy_sample(d)
    n_tables = len(Document(str(docx)).tables)
    api.run_export(docx, dry_run=True)
    check("t01 预演不生成工作簿", not api.workbook_path(docx).exists())
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    check("t01 工作簿生成", xlsx.exists())
    wb, order, blocks = blocks_of(xlsx)
    check("t01 表块数=Word表数", len(order) == n_tables, f"{len(order)} vs {n_tables}")
    check("t01 每块有书签", all(blocks[c]["bookmark"].startswith(("Table_", "CPAH_")) for c in order))


def t02_writeback_value():
    d = ws_dir("t02")
    docx = copy_sample(d)
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, order, blocks = blocks_of(xlsx)
    target = None
    for code in order:
        b = blocks[code]
        for r in b["grid_rows"][1:]:
            for c in range(b["col0"], b["col0"] + b["ncols"]):
                v = wb[SHEET_NAME].cell(r, c).value
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    target = (code, r, c, v)
                    break
            if target:
                break
        if target:
            break
    check("t02 找到数值单元格", target is not None)
    if not target:
        return
    code, r, c, old = target
    ws = wb[SHEET_NAME]
    old_text = display_text(old, ws.cell(r, c).number_format).strip()
    ws.cell(r, c).value = 12345.67
    new_text = display_text(12345.67, ws.cell(r, c).number_format).strip()
    wb.save(str(xlsx))
    before = docx.read_bytes()
    api.run_writeback(docx, dry_run=True)
    check("t02 预演不改Word", docx.read_bytes() == before)
    api.run_writeback(docx)
    texts = {c for t in word_tables(docx) for row in t for c in row}
    check("t02 新值写入Word", new_text in texts, f"期望 {new_text!r}")
    check("t02 旧值已被替换", old_text not in texts or old_text == new_text, f"旧值 {old_text!r} 仍在")


def t03_rows_cols():
    d = ws_dir("t03")
    docx = make_doc(d / "r.docx", [("表1 明细", [
        ["项目", "本期", "上期"], ["甲", "1", "2"], ["乙", "3", "4"], ["合计", "4", "6"]])])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    # 增行：在合计行前插一行并填数
    wb, order, blocks = blocks_of(xlsx)
    ws = wb[SHEET_NAME]
    b = blocks["T01"]
    ws.insert_rows(b["grid_rows"][3])
    for j, v in enumerate(["丙", "5", "6"]):
        ws.cell(b["grid_rows"][3], b["col0"] + j).value = v
    wb.save(str(xlsx))
    api.run_writeback(docx)
    t = word_tables(docx)[0]
    check("t03 增行生效", len(t) == 5 and t[3][0] == "丙" and t[4][0] == "合计", str(t))
    # 删行：删掉 乙
    wb, order, blocks = blocks_of(xlsx)
    ws = wb[SHEET_NAME]
    b = blocks["T01"]
    row_yi = next(r for r in b["grid_rows"] if str(ws.cell(r, b["col0"]).value) == "乙")
    ws.delete_rows(row_yi)
    wb.save(str(xlsx))
    api.run_writeback(docx)
    t = word_tables(docx)[0]
    check("t03 删行生效", len(t) == 4 and all(row[0] != "乙" for row in t), str(t))
    # 增列：末尾加「备注」列
    wb, order, blocks = blocks_of(xlsx)
    ws = wb[SHEET_NAME]
    b = blocks["T01"]
    newcol = b["col0"] + b["ncols"]
    for i, r in enumerate(b["grid_rows"]):
        ws.cell(r, newcol).value = "备注" if i == 0 else "x"
    wb.save(str(xlsx))
    api.run_writeback(docx)
    t = word_tables(docx)[0]
    check("t03 增列生效", len(t[0]) == 4 and t[0][3] == "备注", str(t[0]))
    # 整列空白但带格式的新增列：只忽略、不写 Word；改动留痕不报错（曾因 len(int) 崩溃）
    from openpyxl.styles import Border, Side
    wb, order, blocks = blocks_of(xlsx)
    ws = wb[SHEET_NAME]
    b = blocks["T01"]
    blank_col = b["col0"] + b["ncols"]
    side = Side(style="thin")
    for r in b["grid_rows"]:
        ws.cell(r, blank_col).border = Border(left=side, right=side, top=side, bottom=side)
    wb.save(str(xlsx))
    records = api.scan_writeback_diff(docx)
    check("t03 空白新增列留痕", any(r.get("kind") == "列结构" for r in records), str(records))
    api.run_writeback(docx)
    t = word_tables(docx)[0]
    check("t03 空白新增列不写入", len(t[0]) == 4, str(t[0]))


def t04_mixed_bookmark_styles():
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from wesync.docx_bookmarks import next_bookmark_id, remove_bookmark

    d = ws_dir("t04")
    docx = make_doc(d / "r.docx", [
        ("表1 明细", [["项目", "金额"], ["甲", "1"]]),
        ("表2 明细", [["项目", "金额"], ["乙", "2"]]),
    ])
    api.run_export(docx)
    _, _, blocks = blocks_of(api.workbook_path(docx))
    name = blocks["T02"]["bookmark"]

    doc = Document(str(docx))
    remove_bookmark(doc, name)
    caption = next(p for p in doc.paragraphs if p.text == "表2 明细")
    bookmark_id = next_bookmark_id(doc)
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    caption._p.insert(0, start)
    caption._p.append(end)
    doc.save(str(docx))

    wb = load_workbook(str(api.workbook_path(docx)))
    ws = wb[SHEET_NAME]
    ws.cell(blocks["T01"]["grid_rows"][1], blocks["T01"]["col0"] + 1).value = 101
    ws.cell(blocks["T02"]["grid_rows"][1], blocks["T02"]["col0"] + 1).value = 202
    wb.save(str(api.workbook_path(docx)))

    api.run_writeback(docx)
    tables = word_tables(docx)
    check("t04 新式书签表更新", tables[0][1][1] == "101", str(tables[0]))
    check("t04 旧式书签表也更新", tables[1][1][1] == "202", str(tables[1]))


def t05_failed_table_rolls_back_partial_changes():
    d = ws_dir("t05")
    docx = make_doc(d / "r.docx", [
        ("表1 明细", [["项目", "金额"], ["甲", "1"]]),
        ("表2 明细", [["项目", "金额"], ["乙", "2"]]),
    ])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, _, blocks = blocks_of(xlsx)
    ws = wb[SHEET_NAME]
    first = blocks["T01"]
    new_col = first["col0"] + first["ncols"]
    for index, row in enumerate(first["grid_rows"]):
        ws.cell(row, new_col).value = "备注" if index == 0 else "x"
    ws.cell(first["grid_rows"][1], first["col0"] + 1).value = "=[1]Sheet1!$C$5"
    second = blocks["T02"]
    ws.cell(second["grid_rows"][1], second["col0"] + 1).value = 202
    wb.save(str(xlsx))

    before_xml = Document(str(docx)).tables[0]._tbl.xml
    LOGS.clear()
    api.run_writeback(docx)
    after_xml = Document(str(docx)).tables[0]._tbl.xml
    check("t05 失败表XML完全恢复", after_xml == before_xml)
    check("t05 另一张表正常同步", word_tables(docx)[1][1][1] == "202", str(word_tables(docx)[1]))
    check("t05 日志说明单表回滚", any("该表已回滚" in message for message in LOGS), str(LOGS[-6:]))
