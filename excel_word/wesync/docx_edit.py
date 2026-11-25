"""Word 单元格写入：替换文字但保留 run 格式、段落数量与单元格属性。"""

import copy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from .docx_grid import cell_text


def _remove_text_nodes(p_el) -> None:
    for t in list(p_el.findall(qn("w:r"))):
        p_el.remove(t)
    for hl in list(p_el.findall(qn("w:hyperlink"))):
        p_el.remove(hl)


def _set_paragraph_text(p_el, text: str, fallback_rpr=None) -> None:
    """写入段落文本，保留第一个 run 的格式。"""
    runs = p_el.findall(qn("w:r"))
    for hl in list(p_el.findall(qn("w:hyperlink"))):
        p_el.remove(hl)
    if runs:
        keep = runs[0]
        for extra in runs[1:]:
            p_el.remove(extra)
        texts = keep.findall(qn("w:t"))
        for idx, node in enumerate(texts):
            if idx == 0:
                node.text = text
                node.set(qn("xml:space"), "preserve")
            else:
                keep.remove(node)
        if not texts:
            node = OxmlElement("w:t")
            node.text = text
            node.set(qn("xml:space"), "preserve")
            keep.append(node)
        return

    rpr = None
    pPr = p_el.find(qn("w:pPr"))
    if pPr is not None:
        rpr = pPr.find(qn("w:rPr"))
    if rpr is None:
        rpr = fallback_rpr
    run = OxmlElement("w:r")
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    node = OxmlElement("w:t")
    node.text = text
    node.set(qn("xml:space"), "preserve")
    run.append(node)
    p_el.append(run)


def set_cell_text(tc, text: str, fallback_rpr=None) -> bool:
    """把单元格文本设为 text（\n 表示多段），保留格式。返回是否发生改动。"""
    if tc is None:
        return False
    if cell_text(tc) == text:
        return False
    segments = text.split("\n") if text != "" else [""]
    paras = tc.findall(qn("w:p"))
    if not paras:
        node = OxmlElement("w:p")
        tc.append(node)
        paras = [node]
    while len(paras) > len(segments):
        extra = paras.pop()
        tc.remove(extra)
    while len(paras) < len(segments):
        clone = copy.deepcopy(paras[-1])
        _remove_text_nodes(clone)
        paras[-1].addnext(clone)
        paras = tc.findall(qn("w:p"))
    for p_el, segment in zip(paras, segments):
        _set_paragraph_text(p_el, segment, fallback_rpr)
    return True
