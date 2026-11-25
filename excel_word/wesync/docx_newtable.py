"""新表创建层：按 Excel 网格造标准样式的 Word 表格，插到指定表格之后（带表标题段落）。

样式与导出观感一致：宋体 9、上下细线内部点线、列宽按 Excel 列宽比例分配页面宽度；
首行为表头（加粗 + tblHeader），首列以「合计」开头的末行加粗；支持 gridSpan/vMerge。
"""

from __future__ import annotations

import copy

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl.utils import get_column_letter

from .config import FONT_NAME, FONT_SIZE, NUMBER_RE
from .docx_grid import para_first_run_rpr, para_text

EMU_PER_DXA = 635


def _el(tag: str, **attrs):
    el = OxmlElement(tag)
    for key, value in attrs.items():
        el.set(qn(f"w:{key}"), str(value))
    return el


def _run(text: str, bold: bool):
    r = _el("w:r")
    rpr = _el("w:rPr")
    rpr.append(_el("w:rFonts", ascii="Times New Roman", hAnsi="Times New Roman", eastAsia=FONT_NAME))
    if bold:
        rpr.append(_el("w:b"))
    rpr.append(_el("w:sz", val=FONT_SIZE * 2))
    rpr.append(_el("w:szCs", val=FONT_SIZE * 2))
    r.append(rpr)
    t = _el("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def _cell_paragraphs(text: str, bold: bool, align: str) -> list:
    parts = text.split("\n") if text else [""]
    result = []
    for part in parts:
        p = _el("w:p")
        ppr = _el("w:pPr")
        ppr.append(_el("w:jc", val=align))
        p.append(ppr)
        if part:
            p.append(_run(part, bold))
        result.append(p)
    return result


def _table_borders_el():
    borders = _el("w:tblBorders")
    for side, val in (("top", "single"), ("left", "nil"), ("bottom", "single"),
                      ("right", "nil"), ("insideH", "dotted"), ("insideV", "dotted")):
        borders.append(_el(f"w:{side}", val=val, sz=4, space=0, color="auto"))
    return borders


def available_width_dxa(doc) -> int:
    try:
        sec = doc.sections[0]
        return max(2000, int((sec.page_width - sec.left_margin - sec.right_margin) / EMU_PER_DXA))
    except Exception:
        return 9000


def col_widths_from_excel(ws, ncols: int, total_dxa: int, col0: int = 2) -> list[int]:
    widths = []
    for c in range(ncols):
        letter = get_column_letter(col0 + c)
        dim = ws.column_dimensions.get(letter)
        widths.append(float(dim.width) if dim is not None and dim.width else 8.43)
    total = sum(widths) or 1.0
    result = [max(400, int(total_dxa * w / total)) for w in widths]
    result[-1] += total_dxa - sum(result)
    return result


def build_table(grid: list[list[str]], merges: list[tuple[int, int, int, int]],
                ncols: int, col_widths: list[int]):
    """按网格文本与合并区造 w:tbl。grid 非起点位置的文字被忽略。"""
    nrows = len(grid)
    span_of: dict[tuple[int, int], tuple[int, int]] = {}
    hcovered: set[tuple[int, int]] = set()
    vcont: dict[tuple[int, int], int] = {}
    for (r, c, rs, cs) in merges:
        span_of[(r, c)] = (rs, cs)
        for k in range(1, cs):
            hcovered.add((r, c + k))
        for rr in range(r + 1, r + rs):
            vcont[(rr, c)] = cs
            for k in range(1, cs):
                hcovered.add((rr, c + k))

    total_row = nrows - 1 if nrows and grid[-1] and grid[-1][0].strip().startswith("合计") else -1

    tbl = _el("w:tbl")
    tblpr = _el("w:tblPr")
    tblpr.append(_el("w:tblW", w=sum(col_widths), type="dxa"))
    tblpr.append(_table_borders_el())
    tblpr.append(_el("w:tblLayout", type="fixed"))
    tbl.append(tblpr)
    tblgrid = _el("w:tblGrid")
    for width in col_widths:
        tblgrid.append(_el("w:gridCol", w=width))
    tbl.append(tblgrid)

    for r in range(nrows):
        tr = _el("w:tr")
        if r == 0:
            trpr = _el("w:trPr")
            trpr.append(_el("w:tblHeader"))
            tr.append(trpr)
        bold_row = r == 0 or r == total_row
        c = 0
        while c < ncols:
            if (r, c) in hcovered:
                c += 1
                continue
            if (r, c) in vcont:
                cs = vcont[(r, c)]
                tc = _el("w:tc")
                tcpr = _el("w:tcPr")
                tcpr.append(_el("w:tcW", w=sum(col_widths[c:c + cs]), type="dxa"))
                if cs > 1:
                    tcpr.append(_el("w:gridSpan", val=cs))
                tcpr.append(_el("w:vMerge"))
                tcpr.append(_el("w:vAlign", val="center"))
                tc.append(tcpr)
                tc.append(_el("w:p"))
                tr.append(tc)
                c += cs
                continue
            rs, cs = span_of.get((r, c), (1, 1))
            text = grid[r][c] if c < len(grid[r]) else ""
            tc = _el("w:tc")
            tcpr = _el("w:tcPr")
            tcpr.append(_el("w:tcW", w=sum(col_widths[c:c + cs]), type="dxa"))
            if cs > 1:
                tcpr.append(_el("w:gridSpan", val=cs))
            if rs > 1:
                tcpr.append(_el("w:vMerge", val="restart"))
            tcpr.append(_el("w:vAlign", val="center"))
            tc.append(tcpr)
            if r == 0:
                align = "center"
            elif NUMBER_RE.match(text.strip()):
                align = "right"
            else:
                align = "left"
            for p in _cell_paragraphs(text, bold_row, align):
                tc.append(p)
            tr.append(tc)
            c += cs
        tbl.append(tr)
    return tbl


def _caption_source_p(prev_tbl):
    """前一张表上方最近的非空段落，作为新表标题的格式样板。"""
    if prev_tbl is None:
        return None
    node = prev_tbl.getprevious()
    while node is not None and node.tag == qn("w:p"):
        if para_text(node).strip():
            return node
        node = node.getprevious()
    return None


def _caption_paragraph(text: str, source_p):
    p = _el("w:p")
    if source_p is not None:
        ppr = source_p.find(qn("w:pPr"))
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
    rpr = para_first_run_rpr(source_p)
    r = _el("w:r")
    if rpr is not None:
        r.append(rpr)
    else:
        default_rpr = _el("w:rPr")
        default_rpr.append(_el("w:rFonts", ascii="Times New Roman", hAnsi="Times New Roman", eastAsia=FONT_NAME))
        default_rpr.append(_el("w:sz", val=21))
        r.append(default_rpr)
    t = _el("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p.append(r)
    return p


def insert_new_table(doc, prev_tbl, tbl, caption_text: str) -> None:
    """把新表插到 prev_tbl 之后（无 prev_tbl 时插到文档末尾），前面带标题段落。

    Word 要求两张表之间、以及文档末尾的表之后必须有段落，这里一并补齐。
    """
    caption = _caption_paragraph(caption_text, _caption_source_p(prev_tbl))
    if prev_tbl is not None and prev_tbl.getparent() is not None:
        prev_tbl.addnext(caption)
        caption.addnext(tbl)
    else:
        body = doc.element.body
        sectpr = body.find(qn("w:sectPr"))
        if sectpr is not None:
            sectpr.addprevious(caption)
        else:
            body.append(caption)
        caption.addnext(tbl)
    nxt = tbl.getnext()
    if nxt is None or nxt.tag in (qn("w:tbl"), qn("w:sectPr")):
        tbl.addnext(_el("w:p"))