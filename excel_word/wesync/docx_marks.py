"""高亮标记层：把 Word 里的高亮文字识别成“标记”，并做可定位、可写回的锚点。

约定：
* 高亮（``w:highlight``，任意颜色）视为“待处理位置”的输入标记；
* 每个标记（同一段落内连续的高亮 run 合并为一个）用一个不可见书签 ``MRK_0001`` 定位；
* 书签只做锚点，不修改任何文字；同步写回时才按 Excel 内容替换文字并去掉高亮。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from .docx_bookmarks import bookmark_names, next_bookmark_id

MARK_PREFIX = "MRK_"
MARK_NAME_RE = re.compile(rf"^{MARK_PREFIX}\d{{4}}$")


@dataclass
class Mark:
    """一个高亮标记。"""

    index: int                       # 1 基序号（导出为 M01、M02…）
    p_el: object                     # 所在段落元素
    children: list = field(default_factory=list)   # 组成该标记的段落子元素（run 或 hyperlink）
    runs: list = field(default_factory=list)       # 组成该标记的 run 列表
    bookmark: str = ""               # 锚点书签名（MRK_xxxx）
    text: str = ""                   # 标记当前文字
    context: str = ""                # 所在段落文本（截断，便于在 Excel 里辨认）
    location: str = ""               # 位置描述：正文 第N段 / 表N


def _highlight_val(r_el) -> str | None:
    rPr = r_el.find(qn("w:rPr"))
    if rPr is None:
        return None
    hl = rPr.find(qn("w:highlight"))
    if hl is None:
        return None
    return hl.get(qn("w:val"))


def run_highlighted(r_el) -> bool:
    return _highlight_val(r_el) not in (None, "none")


def _payload_run(r_el) -> bool:
    """单个 run 元素是否是高亮标记的一部分。"""
    return r_el.tag == qn("w:r") and run_highlighted(r_el)


def _payload_of_child(child) -> list:
    """把一个段落子元素（run 或 hyperlink）转成“标记 run 列表”；不是标记则返回空。"""
    if child.tag == qn("w:r"):
        return [child] if _payload_run(child) else []
    if child.tag == qn("w:hyperlink"):
        runs = [r for r in child.iter(qn("w:r"))]
        if runs and all(_payload_run(r) for r in runs):
            return runs
    return []


def _run_texts(runs: list) -> str:
    return "".join("".join(t.text or "" for t in r.iter(qn("w:t"))) for r in runs)


def _bookmark_ranges_in_paragraph(p_el, prefix: str = "") -> list[tuple[str, int, int]]:
    """返回段落内的书签范围 ``[(名字, 起始子元素下标, 结束子元素下标)]``。"""
    children = list(p_el.iterchildren())
    starts: dict[str, tuple[str, int]] = {}
    ranges: list[tuple[str, int, int]] = []
    for index, child in enumerate(children):
        if child.tag == qn("w:bookmarkStart"):
            name = child.get(qn("w:name")) or ""
            if not prefix or name.startswith(prefix):
                starts[child.get(qn("w:id"))] = (name, index)
        elif child.tag == qn("w:bookmarkEnd"):
            key = child.get(qn("w:id"))
            if key in starts:
                name, start = starts.pop(key)
                ranges.append((name, start, index))
    return ranges


def _scan_paragraph(p_el, location: str, marks: list[Mark]) -> None:
    children = list(p_el.iterchildren())
    existing = {start: name for name, start, _end in _bookmark_ranges_in_paragraph(p_el, MARK_PREFIX)}
    paragraph_text = "".join(t.text or "" for t in p_el.iter(qn("w:t"))).strip()

    group: list = []
    group_start = 0
    for index, child in enumerate(children):
        payload = _payload_of_child(child)
        if payload:
            if not group:
                group_start = index
            group.append((child, payload))
            continue
        if group:
            _emit_mark(p_el, children, group, group_start, index - 1, location, paragraph_text, existing, marks)
            group = []
    if group:
        _emit_mark(p_el, children, group, group_start, len(children) - 1, location, paragraph_text, existing, marks)


def _emit_mark(p_el, children, group, start_index, end_index, location, paragraph_text, existing, marks) -> None:
    runs: list = []
    for _child, payload in group:
        runs.extend(payload)
    if not runs:
        return
    mark = Mark(
        index=len(marks) + 1,
        p_el=p_el,
        children=[child for child, _payload in group],
        runs=runs,
        text=_run_texts(runs),
        context=paragraph_text[:80],
        location=location,
    )
    # 复用已经存在的 MRK_ 书签（重复导出时保持稳定）
    name = existing.get(start_index - 1)
    if name:
        mark.bookmark = name
    marks.append(mark)


def find_marks(doc) -> list[Mark]:
    """按文档顺序找出所有高亮标记（正文段落与表格单元格内的段落都算）。"""
    marks: list[Mark] = []
    for p_el, location in iter_paragraph_locations(doc):
        _scan_paragraph(p_el, location, marks)
    return marks


def iter_paragraph_locations(doc):
    """按文档顺序产出 ``(段落元素, 位置描述)``；表格内段落标为“表N”，其余标为“正文 第N段”。"""
    para_no = 0
    table_no = 0
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[1]
        if tag == "p":
            para_no += 1
            yield child, f"正文 第{para_no}段"
        elif tag == "tbl":
            table_no += 1
            for p_el in child.iter(qn("w:p")):
                yield p_el, f"表{table_no}"


def processed_marks(doc, pending: set[str]) -> list[dict]:
    """找出“已处理”的标记：书签仍在文档里，但该处已不再带高亮（通常是同步过）。"""
    result: list[dict] = []
    for p_el, location in iter_paragraph_locations(doc):
        for name, start, end in _bookmark_ranges_in_paragraph(p_el, MARK_PREFIX):
            if not name or name in pending:
                continue
            children = list(p_el.iterchildren())[start + 1 : end]
            text = "".join(
                "".join(t.text or "" for t in child.iter(qn("w:t"))) for child in children
            )
            result.append({"bookmark": name, "text": text.strip(), "location": location})
    return result


def ensure_mark_bookmarks(doc, marks: list[Mark]) -> int:
    """给还没有书签的标记插入 ``MRK_xxxx`` 锚点，返回新建数量。"""
    existing = bookmark_names(doc.element.body)
    numbers = [int(name[len(MARK_PREFIX):]) for name in existing if MARK_NAME_RE.match(name)]
    next_no = (max(numbers) + 1) if numbers else 1
    bookmark_id = next_bookmark_id(doc)
    created = 0
    for mark in marks:
        if mark.bookmark:
            continue
        name = f"{MARK_PREFIX}{next_no:04d}"
        while name in existing:
            next_no += 1
            name = f"{MARK_PREFIX}{next_no:04d}"
        next_no += 1
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(bookmark_id))
        start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(bookmark_id))
        first = mark.children[0]
        last = mark.children[-1]
        first.addprevious(start)
        last.addnext(end)
        mark.bookmark = name
        existing.add(name)
        bookmark_id += 1
        created += 1
    return created


def mark_payload_by_bookmark(doc, name: str):
    """按书签名找标记：返回 ``(段落元素, 标记子元素列表, run 列表)``；找不到返回 None。"""
    for p_el in doc.element.body.iter(qn("w:p")):
        for range_name, start, end in _bookmark_ranges_in_paragraph(p_el):
            if range_name != name:
                continue
            children = list(p_el.iterchildren())[start + 1 : end]
            runs: list = []
            for child in children:
                if child.tag == qn("w:r"):
                    runs.append(child)
                elif child.tag == qn("w:hyperlink"):
                    runs.extend(child.iter(qn("w:r")))
            return p_el, children, runs
    return None


def _remove_highlight(r_el) -> None:
    rPr = r_el.find(qn("w:rPr"))
    if rPr is None:
        return
    hl = rPr.find(qn("w:highlight"))
    if hl is not None:
        rPr.remove(hl)


def apply_mark_change(
    part, p_el, children: list, runs: list, new_text: str, link: str, keep_highlight: bool = False
) -> str:
    """把标记写回：替换文字（保留格式）、按需加超链接、去掉高亮。

    ``keep_highlight=True`` 时保留该处高亮（便于复核），其余行为相同。
    返回 ``changed`` / ``unchanged`` / ``deleted``。
    """
    if not runs:
        return "deleted"
    current_text = _run_texts(runs)
    target_text = new_text if new_text != "" else current_text
    if new_text == "" and link == "":
        return "unchanged"

    # 先看当前是否已经是超链接、指向哪里（用于幂等判断）
    existing_link = ""
    for child in children:
        if child.tag == qn("w:hyperlink"):
            rel_id = child.get(qn("r:id"))
            rel = part.rels.get(rel_id) if rel_id else None
            existing_link = rel.target_ref if rel is not None else ""
            break

    highlight_pending = (not keep_highlight) and any(run_highlighted(r) for r in runs)
    changed = current_text != target_text or (link or "") != (existing_link or "") or highlight_pending
    if not changed:
        return "unchanged"

    # 解开已有超链接，改由第一个 run 承载文字
    for child in list(children):
        if child.tag == qn("w:hyperlink"):
            for r in list(child.iter(qn("w:r"))):
                if r.getparent() is child:
                    child.addprevious(r)
            p_el.remove(child)
    runs = [r for r in runs if r.getparent() is not None]
    if not runs:
        return "deleted"

    keep = runs[0]
    texts = keep.findall(qn("w:t"))
    for index, node in enumerate(texts):
        node.text = target_text if index == 0 else ""
        node.set(qn("xml:space"), "preserve")
    if not texts:
        node = OxmlElement("w:t")
        node.text = target_text
        node.set(qn("xml:space"), "preserve")
        keep.append(node)
    if not keep_highlight:
        _remove_highlight(keep)
    for extra in runs[1:]:
        parent = extra.getparent()
        if parent is not None:
            parent.remove(extra)

    if link:
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), part.relate_to(link, RT.HYPERLINK, is_external=True))
        hyperlink.set(qn("w:history"), "1")
        keep.addprevious(hyperlink)
        hyperlink.append(keep)
    return "changed"
