"""回归测试共用工具：工作目录、日志捕获、构造文档、断言（失败即报错）。"""
import atexit
import posixpath
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from docx import Document  # noqa: E402
from openpyxl import load_workbook  # noqa: E402
from wesync import api, console  # noqa: E402
from wesync.config import SHEET_NAME  # noqa: E402
from wesync.excel_blocks import parse_blocks  # noqa: E402
from wesync.fsops import tmp_root  # noqa: E402
from wesync.valuefmt import display_text  # noqa: E402

# 约定：测试工作目录落在项目内 .tmp/（不写系统 Temp），进程结束自动清理
WORK = Path(tempfile.mkdtemp(prefix="wesync_tests_", dir=tmp_root()))
atexit.register(shutil.rmtree, WORK, ignore_errors=True)
LOGS = []
console.set_log_sink(LOGS.append)


def ws_dir(name):
    d = WORK / name
    d.mkdir(parents=True, exist_ok=True)
    return d


def copy_sample(d, name="报告.docx"):
    return make_doc(d / name, [
        ("表1 明细", [["项目", "金额"], ["甲", 100.50], ["乙", 200.00]]),
        ("表2 明细", [["项目", "金额"], ["丙", 300.25], ["丁", 400.75]]),
    ])


def set_formula_cached_value(xlsx, sheet_name, coordinate, value):
    """给已写入的公式补上 Excel 缓存结果，供 data_only 读取测试。"""
    xlsx = Path(xlsx)
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    with zipfile.ZipFile(xlsx) as zf:
        files = {info.filename: zf.read(info.filename) for info in zf.infolist()}
        infos = zf.infolist()
    wb_root = ET.fromstring(files["xl/workbook.xml"])
    sheet = next(s for s in wb_root.findall(f".//{{{main_ns}}}sheet") if s.attrib["name"] == sheet_name)
    rid = sheet.attrib[f"{{{rel_ns}}}id"]
    rel_root = ET.fromstring(files["xl/_rels/workbook.xml.rels"])
    target = next(r.attrib["Target"] for r in rel_root.findall(f"{{{pkg_ns}}}Relationship") if r.attrib["Id"] == rid)
    sheet_path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
    root = ET.fromstring(files[sheet_path])
    cell = next(c for c in root.findall(f".//{{{main_ns}}}c") if c.attrib.get("r") == coordinate)
    if cell.find(f"{{{main_ns}}}f") is None:
        raise ValueError(f"{sheet_name}!{coordinate} 不是公式单元格")
    cached = cell.find(f"{{{main_ns}}}v")
    if cached is None:
        cached = ET.SubElement(cell, f"{{{main_ns}}}v")
    if isinstance(value, str):
        cell.set("t", "str")
        cached.text = value
    elif isinstance(value, bool):
        cell.set("t", "b")
        cached.text = "1" if value else "0"
    else:
        cell.attrib.pop("t", None)
        cached.text = str(value)
    files[sheet_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    tmp = xlsx.with_name(xlsx.stem + ".cache.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for info in infos:
            zf.writestr(info, files[info.filename])
    tmp.replace(xlsx)


def make_doc(path, tables):
    """tables: [(caption, rows)]，rows 首行为表头。"""
    doc = Document()
    doc.add_paragraph("报告书正文开头。")
    for caption, rows in tables:
        if caption:
            doc.add_paragraph(caption)
        t = doc.add_table(rows=len(rows), cols=len(rows[0]))
        t.style = "Table Grid"
        for i, row in enumerate(rows):
            for j, text in enumerate(row):
                t.cell(i, j).text = str(text)
        doc.add_paragraph("")
    doc.save(str(path))
    return path


def blocks_of(xlsx):
    wb = load_workbook(str(xlsx))
    order, blocks = parse_blocks(wb[SHEET_NAME])
    return wb, order, blocks


def grid_text(ws, block):
    col0 = block.get("col0", 2)
    return [
        [display_text(ws.cell(r, col0 + c).value, ws.cell(r, col0 + c).number_format).strip()
         for c in range(block["ncols"])]
        for r in block["grid_rows"]
    ]


def word_tables(path):
    return [[[c.text for c in row.cells] for row in t.rows] for t in Document(str(path)).tables]


def check(name, cond, detail=""):
    if not cond:
        raise AssertionError(f"{name} 失败：{detail}")