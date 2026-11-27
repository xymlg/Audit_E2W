"""高亮标记层：识别显式高亮并建立可定位、可写回的书签锚点。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from .docx_bookmarks import bookmark_names, next_bookmark_id
from .errors import SyncError

MARK_PREFIX = "MRK_"
MARK_NAME_RE = re.compile(rf"^{MARK_PREFIX}\d{{4}}$")


@dataclass
class Mark:
    index: int
    p_el: object
    children: list = field(default_factory=list)
    runs: list = field(default_factory=list)
    bookmark: str = ""
    text: str = ""
    context: str = ""
    location: str = ""
    part: object = None
    doc: object = None


def _highlight_val(r_el) -> str | None:
    rPr = r_el.find(qn("w:rPr"))
    if rPr is None:
        return None
    hl = rPr.find(qn("w:highlight"))
    if hl is None:
        return None
    return hl.get(qn("w:val"))


def _style_highlight_val(doc, style_id: str | None) -> str | None:
    if doc is None or not style_id:
        return None
    try:
        styles = doc.styles.element
    except Exception:
        return None
    style = next((s for s in styles.findall(qn("w:style")) if s.get(qn("w:styleId")) == style_id), None)
    if style is None or style.get(qn("w:type")) != "character":
        return None
    seen = set()
    while style is not None and style.get(qn("w:type")) == "character":
        sid = style.get(qn("w:styleId"))
        if sid in seen:
            return None
        seen.add(sid)
        rPr = style.find(qn("w:rPr"))
        hl = rPr.find(qn("w:highlight")) if rPr is not None else None
        if hl is not None:
            return hl.get(qn("w:val"))
        based = style.find(qn("w:basedOn"))
        parent_id = based.get(qn("w:val")) if based is not None else None
        style = next((s for s in styles.findall(qn("w:style")) if s.get(qn("w:styleId")) == parent_id), None)
    return None


def run_highlighted(r_el, doc=None) -> bool:
    value = _highlight_val(r_el)
    if value is not None:
        return value != "none"
    rPr = r_el.find(qn("w:rPr"))
    style_id = rPr.find(qn("w:rStyle")).get(qn("w:val")) if rPr is not None and rPr.find(qn("w:rStyle")) is not None else None
    return _style_highlight_val(doc, style_id) not in (None, "none")


def _style_highlighted(r_el, doc) -> bool:
    rPr = r_el.find(qn("w:rPr"))
    style_id = rPr.find(qn("w:rStyle")).get(qn("w:val")) if rPr is not None and rPr.find(qn("w:rStyle")) is not None else None
    return _highlight_val(r_el) is None and _style_highlight_val(doc, style_id) not in (None, "none")


def _paragraph_mark_highlighted(p_el) -> bool:
    pPr = p_el.find(qn("w:pPr"))
    rPr = pPr.find(qn("w:rPr")) if pPr is not None else None
    hl = rPr.find(qn("w:highlight")) if rPr is not None else None
    return hl is not None and hl.get(qn("w:val")) != "none"


def _inside_table(p_el) -> bool:
    parent = p_el.getparent()
    while parent is not None:
        if parent.tag == qn("w:tc"):
            return True
        parent = parent.getparent()
    return False


def _run_texts(runs: list) -> str:
    return "".join("".join(t.text or "" for t in r.iter(qn("w:t"))) for r in runs)


def _containers(doc):
    seen: set[int] = set()
    yield doc.element.body, "body"
    seen.add(id(doc.element.body))
    for section_index, section in enumerate(doc.sections, 1):
        for label, part in (
            ("页眉", section.header),
            ("页脚", section.footer),
            ("首页页眉", section.first_page_header),
            ("首页页脚", section.first_page_footer),
            ("偶数页页眉", section.even_page_header),
            ("偶数页页脚", section.even_page_footer),
        ):
            if part is not None and id(part._element) not in seen:
                seen.add(id(part._element))
                yield part._element, f"第{section_index}节{label}"


def _bookmark_ranges_in_paragraph(p_el, prefix: str = "") -> list[tuple[str, int, int]]:
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


def _duplicate_mark_bookmarks(doc) -> dict[str, int]:
    counts: dict[str, int] = {}
    for root, _label in _containers(doc):
        for start in root.iter(qn("w:bookmarkStart")):
            name = start.get(qn("w:name")) or ""
            if MARK_NAME_RE.match(name):
                counts[name] = counts.get(name, 0) + 1
    duplicates = {name: count for name, count in counts.items() if count > 1}
    if duplicates:
        detail = "、".join(f"{name}（{count} 处）" for name, count in sorted(duplicates.items()))
        raise SyncError(f"Word 中存在重复的标注书签：{detail}。请先清理重复书签，再导出或同步。")
    return counts


def _payload_of_child(child, doc):
    if child.tag == qn("w:r"):
        return [child] if run_highlighted(child, doc) else []
    if child.tag == qn("w:hyperlink"):
        runs = [r for r in child.iter(qn("w:r"))]
        if runs and all(run_highlighted(r, doc) for r in runs):
            return runs
    return []


def _scan_paragraph(p_el, part, doc, location: str, marks: list[Mark]) -> None:
    if _inside_table(p_el):
        return
    children = list(p_el.iterchildren())
    existing = {start: name for name, start, _end in _bookmark_ranges_in_paragraph(p_el, MARK_PREFIX)}
    paragraph_text = "".join(t.text or "" for t in p_el.iter(qn("w:t"))).strip()

    group: list = []
    group_start = 0
    paragraph_mark = _paragraph_mark_highlighted(p_el)
    for index, child in enumerate(children):
        payload = _payload_of_child(child, doc)
        if paragraph_mark and child.tag in (qn("w:r"), qn("w:hyperlink")):
            payload = list(child.iter(qn("w:r")))
        if payload:
            if not group:
                group_start = index
            group.append((child, payload))
            continue
        if group and not paragraph_mark:
            _emit_mark(p_el, part, doc, children, group, group_start, index - 1, location, paragraph_text, existing, marks)
            group = []
    if group:
        _emit_mark(p_el, part, doc, children, group, group_start, len(children) - 1, location, paragraph_text, existing, marks)


def _emit_mark(p_el, part, doc, children, group, start_index, end_index, location, paragraph_text, existing, marks) -> None:
    runs: list = []
    for _child, payload in group:
        runs.extend(payload)
    text = _run_texts(runs)
    if not text.strip():
        return
    mark = Mark(
        index=len(marks) + 1,
        p_el=p_el,
        children=[child for child, _payload in group],
        runs=runs,
        text=text,
        context=paragraph_text[:80],
        location=location,
        part=part,
        doc=doc,
    )
    name = existing.get(start_index - 1)
    if name:
        mark.bookmark = name
    marks.append(mark)


def find_marks(doc) -> list[Mark]:
    """按文档顺序找出正文、页眉页脚里的高亮；表格内高亮不参与。"""
    _duplicate_mark_bookmarks(doc)
    marks: list[Mark] = []
    para_no = 0
    for root, label in _containers(doc):
        for child in root.iterchildren():
            if child.tag != qn("w:p"):
                continue
            if label == "body":
                para_no += 1
                location = f"正文 第{para_no}段"
            else:
                location = label
            _scan_paragraph(child, _part_for_root(doc, root), doc, location, marks)
    return marks


def _part_for_root(doc, root):
    if root is doc.element.body:
        return doc.part
    for section in doc.sections:
        for part in (section.header, section.footer, section.first_page_header,
                     section.first_page_footer, section.even_page_header, section.even_page_footer):
            if part is not None and part._element is root:
                return part.part
    return doc.part


def iter_paragraph_locations(doc):
    """按文档顺序产出正文与页眉页脚段落，表格内段落不产出。"""
    para_no = 0
    for root, label in _containers(doc):
        for child in root.iterchildren():
            if child.tag != qn("w:p"):
                continue
            if label == "body":
                para_no += 1
                yield child, f"正文 第{para_no}段"
            else:
                yield child, label


def processed_marks(doc, pending: set[str]) -> list[dict]:
    """找出仍有 MRK 书签但已不高亮的标记。"""
    result: list[dict] = []
    for p_el, location in iter_paragraph_locations(doc):
        for name, start, end in _bookmark_ranges_in_paragraph(p_el, MARK_PREFIX):
            if not name or name in pending:
                continue
            children = list(p_el.iterchildren())[start + 1:end]
            text = "".join("".join(t.text or "" for t in child.iter(qn("w:t"))) for child in children)
            result.append({"bookmark": name, "text": text.strip(), "location": location})
    return result


def ensure_mark_bookmarks(doc, marks: list[Mark]) -> int:
    _duplicate_mark_bookmarks(doc)
    existing = set()
    for root, _label in _containers(doc):
        existing |= bookmark_names(root)
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
    _duplicate_mark_bookmarks(doc)
    for p_el, _location in iter_paragraph_locations(doc):
        for range_name, start, end in _bookmark_ranges_in_paragraph(p_el):
            if range_name != name:
                continue
            children = list(p_el.iterchildren())[start + 1:end]
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


def _add_no_highlight(r_el) -> None:
    rPr = r_el.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r_el.insert(0, rPr)
    hl = rPr.find(qn("w:highlight"))
    if hl is None:
        hl = OxmlElement("w:highlight")
        rPr.append(hl)
    hl.set(qn("w:val"), "none")


def apply_mark_change(part, p_el, children: list, runs: list, new_text: str, link: str,
                      keep_highlight: bool = False, doc=None) -> str:
    """替换标记文字，按需加链接与去色；返回 changed/unchanged/deleted。"""
    if not runs:
        return "deleted"
    current_text = _run_texts(runs)
    target_text = new_text if new_text != "" else current_text
    if new_text == "" and link == "":
        return "unchanged"

    existing_link = ""
    for child in children:
        if child.tag == qn("w:hyperlink"):
            rel_id = child.get(qn("r:id"))
            rel = part.rels.get(rel_id) if rel_id else None
            existing_link = rel.target_ref if rel is not None else ""
            break

    style_highlight = any(_style_highlighted(r, doc) for r in runs)
    highlight_pending = (not keep_highlight) and any(run_highlighted(r, doc) for r in runs)
    changed = current_text != target_text or (link or "") != (existing_link or "") or highlight_pending
    if not changed:
        return "unchanged"

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
        if style_highlight:
            _add_no_highlight(keep)
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
