"""B1：追加新表块 → 建表（含合并/公式）→ 重复同步不重复建表。"""
from helpers import *  # noqa


def t04_new_table():
    d = ws_dir("t04")
    docx = make_doc(d / "n.docx", [
        ("表1 甲", [["项目", "金额"], ["A", "1"], ["合计", "1"]]),
        ("表2 乙", [["项目", "金额"], ["B", "2"], ["合计", "2"]]),
    ])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    code = api.run_append_table_block(docx, name="表新增 披露", nrows=4, ncols=4, after="T01")
    wb, order, blocks = blocks_of(xlsx)
    check("t04 新块插在T01后", order[1] == code and blocks[code]["is_new"], str(order))
    ws = wb[SHEET_NAME]
    b = blocks[code]
    rows = b["grid_rows"]
    data = [["项目", "本期", "上期", "备注"], ["甲", 100, 200, ""],
            ["乙", 300, 400, ""], ["合计", "=SUM(D{0}:D{1})", "=SUM(E{0}:E{1})", ""]]
    fmt = data[3]
    data[3] = ["合计", fmt[1].format(rows[1], rows[2]), fmt[2].format(rows[1], rows[2]), ""]
    for i, row in enumerate(data):
        for j, v in enumerate(row):
            ws.cell(rows[i], b["col0"] + j).value = v
    ws.merge_cells(start_row=rows[1], start_column=b["col0"] + 3, end_row=rows[2], end_column=b["col0"] + 3)
    wb.save(str(xlsx))
    api.run_writeback(docx, dry_run=True)
    check("t04 预演不建表", len(word_tables(docx)) == 2)
    api.run_writeback(docx)
    ts = word_tables(docx)
    check("t04 建表后共3张", len(ts) == 3, str(len(ts)))
    check("t04 位置在第2张", ts[1][0][0] == "项目" and ts[1][1][0] == "甲", str(ts[1]))
    check("t04 公式求值", ts[1][3][1] == "400" and ts[1][3][2] == "600", str(ts[1][3]))
    doc = Document(str(docx))
    check("t04 标题段落存在", any(p.text.strip() == "表新增 披露" for p in doc.paragraphs))
    wb, order, blocks = blocks_of(xlsx)
    check("t04 书签已回填", not blocks[order[1]]["is_new"] and blocks[order[1]]["bookmark"].startswith("Table_"))
    api.run_writeback(docx)
    check("t04 重复同步不重复建表", len(word_tables(docx)) == 3)