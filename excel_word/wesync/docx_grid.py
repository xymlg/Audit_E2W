"""Word 表格网格层：解析 tblGrid/gridSpan/vMerge，得到网格坐标、合并区、边框与对齐、单元格文本。"""

import copy
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from .config import BORDER_MAP, JC_MAP, VA_MAP


class GridPos:
    """表格网格中的一个位置。owner=True 表示该位置是单元格的真实起点。"""

    __slots__ = ("tc", "owner", "vmerge", "span")

    def __init__(self, tc, owner: bool, vmerge: str | None, span: int = 1):
        self.tc = tc
        self.owner = owner
        self.vmerge = vmerge
        self.span = span


def body_items(doc) -> list[tuple[str, object]]:
    """按文档顺序返回正文中的段落与表格。"""
    items: list[tuple[str, object]] = []
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[1]
        if tag == "p":
            items.append(("p", Paragraph(child, doc)))
        elif tag == "tbl":
            items.append(("tbl", Table(child, doc)))
    return items


def parse_tc_props(tc):
    tcPr = tc.find(qn("w:tcPr"))
    span = 1
    vmerge = None
    if tcPr is not None:
        grid_span = tcPr.find(qn("w:gridSpan"))
        if grid_span is not None:
            try:
                span = max(1, int(grid_span.get(qn("w:val"), "1")))
            except (TypeError, ValueError):
                span = 1
        vm = tcPr.find(qn("w:vMerge"))
        if vm is not None:
            vmerge = "restart" if vm.get(qn("w:val")) == "restart" else "continue"
    return span, vmerge


def build_cellmap(tbl):
    """返回 (trs, cellmap, ncols)。cellmap[r][c] 为 GridPos 或 None。"""
    trs = tbl.findall(qn("w:tr"))
    grid = tbl.find(qn("w:tblGrid"))
    ncols = len(grid.findall(qn("w:gridCol"))) if grid is not None else 0
    ncols = max(ncols, 1)
    cellmap: list[list[GridPos | None]] = []
    for tr in trs:
        row: list[GridPos | None] = [None] * ncols
        col = 0
        for tc in tr.findall(qn("w:tc")):
            span, vmerge = parse_tc_props(tc)
            while col < len(row) and row[col] is not None:
                col += 1
            if col + span > len(row):
                row.extend([None] * (col + span - len(row)))
            if vmerge == "continue":
                for k in range(span):
                    row[col + k] = GridPos(None, False, "continue")
            else:
                for k in range(span):
                    row[col + k] = GridPos(tc, k == 0, vmerge, span)
            col += span
        cellmap.append(row)
    ncols = max([len(row) for row in cellmap] + [ncols])
    for row in cellmap:
        row.extend([None] * (ncols - len(row)))
    return trs, cellmap, ncols


def compute_merges(cellmap) -> list[tuple[int, int, int, int]]:
    """返回 (起始行, 起始列, 行跨度, 列跨度)。"""
    merges = []
    for r, row in enumerate(cellmap):
        for c, pos in enumerate(row):
            if pos is None or not pos.owner:
                continue
            colspan = pos.span
            rowspan = 1
            if pos.vmerge == "restart":
                rr = r + 1
                while rr < len(cellmap):
                    nxt = cellmap[rr][c]
                    if nxt is None or nxt.vmerge != "continue":
                        break
                    rowspan += 1
                    rr += 1
            merges.append((r, c, rowspan, colspan))
    return merges


def cell_text(tc) -> str:
    if tc is None:
        return ""
    parts = []
    for p in tc.findall(qn("w:p")):
        parts.append("".join(t.text or "" for t in p.iter(qn("w:t"))))
    return "\n".join(parts)


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def text_grid(tbl, cellmap) -> list[list[str]]:
    grid = []
    for row in cellmap:
        grid.append([cell_text(pos.tc) if (pos is not None and pos.owner) else None for pos in row])
    return grid


def header_rows(trs) -> set[int]:
    result = set()
    for idx, tr in enumerate(trs):
        trPr = tr.find(qn("w:trPr"))
        if trPr is not None and trPr.find(qn("w:tblHeader")) is not None:
            result.add(idx)
    return result


def cell_alignment(tc) -> tuple[str, str, bool]:
    horizontal = "left"
    for p in tc.findall(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            continue
        jc = pPr.find(qn("w:jc"))
        if jc is not None:
            horizontal = JC_MAP.get(jc.get(qn("w:val")), "left")
            break
    vertical = "top"
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is not None:
        va = tcPr.find(qn("w:vAlign"))
        if va is not None:
            vertical = VA_MAP.get(va.get(qn("w:val")), "top")
    multiline = "\n" in cell_text(tc)
    return horizontal, vertical, multiline


def table_borders(tbl) -> dict[str, str | None]:
    """读取表级边框，作为 Excel 边框的默认值。"""
    default = {"top": "thin", "bottom": "thin", "left": None, "right": None, "insideH": "dotted", "insideV": "dotted"}
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        return default
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        return default
    result = dict(default)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{side}"))
        if el is None:
            continue
        val = el.get(qn("w:val"))
        if val in (None, "nil", "none"):
            result[side] = None
        else:
            result[side] = BORDER_MAP.get(val, "thin")
    return result


def explicit_borders(tc) -> dict[str, str | None]:
    result: dict[str, str | None] = {}
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return result
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        return result
    for side in ("top", "left", "bottom", "right"):
        el = borders.find(qn(f"w:{side}"))
        if el is None:
            continue
        val = el.get(qn("w:val"))
        result[side] = None if val in (None, "nil", "none") else BORDER_MAP.get(val, "thin")
    return result


def position_sides(borders: dict, r: int, c: int, nrows: int, ncols: int) -> dict[str, str | None]:
    return {
        "top": borders["top"] if r == 0 else borders["insideH"],
        "bottom": borders["bottom"] if r == nrows - 1 else borders["insideH"],
        "left": borders["left"] if c == 0 else borders["insideV"],
        "right": borders["right"] if c == ncols - 1 else borders["insideV"],
    }


def para_first_run_rpr(p_el):
    """取段落第一个 run 的 rPr（没有 run 时退回段落标记的 rPr）。"""
    if p_el is None:
        return None
    for run in p_el.findall(qn("w:r")):
        rPr = run.find(qn("w:rPr"))
        if rPr is not None:
            return copy.deepcopy(rPr)
    pPr = p_el.find(qn("w:pPr"))
    if pPr is not None:
        rPr = pPr.find(qn("w:rPr"))
        if rPr is not None:
            return copy.deepcopy(rPr)
    return None
