"""Word 书签层：书签读写、Table_ 随机命名与旧书签迁移、遗留〔→Excel〕文字与超链接域清理。"""

import secrets
import string
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from .config import BOOKMARK_PREFIX, EXCEL_LINK_TEXT, LEGACY_BOOKMARK_RE


def bookmark_names(root) -> set[str]:
    names = set()
    for el in root.iter(qn("w:bookmarkStart")):
        value = el.get(qn("w:name"))
        if value:
            names.add(value)
    return names


def used_bookmark_ids(doc) -> set[int]:
    ids = set()
    for el in doc.element.body.iter():
        if el.tag in (qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
            value = el.get(qn("w:id"))
            if value and value.lstrip("-").isdigit():
                ids.add(int(value))
    return ids


def first_bookmark_in_table(tbl) -> str | None:
    for el in tbl.iter(qn("w:bookmarkStart")):
        name = el.get(qn("w:name"))
        if name:
            return name
    return None


def pair_anchor_bookmarks_with_tables(doc, candidates: set[str]) -> dict[str, int]:
    """把“锚点段落里的书签”与“紧随其后的表格”配对。

    只认 ``candidates`` 里的书签名（通常是工作簿 C 列记录的那些）；段落里其它书签
    （例如 ``_Toc…``）忽略、不打断配对。返回 ``{书签名: 表格序号}``（序号即 doc.tables 顺序）：
    没出现在结果里的候选书签，即“对应的表格已从 Word 中删除”。
    """
    if not candidates:
        return {}
    pairing: dict[str, int] = {}
    pending: str | None = None
    table_index = -1
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[1]
        if tag == "p":
            hits = [
                name
                for name in (
                    el.get(qn("w:name")) for el in child.iter(qn("w:bookmarkStart"))
                )
                if name in candidates
            ]
            if hits:
                pending = hits[-1]  # 同一段落有多个候选书签时，取最后插入的那个
        elif tag == "tbl":
            table_index += 1
            if pending is not None:
                pairing[pending] = table_index
                pending = None
    return pairing


def unclaimed_table_indices(doc, pairing: dict[str, int]) -> list[int]:
    """返回没有匹配到任何书签的 Word 表格序号列表。"""
    claimed = set(pairing.values())
    return [index for index in range(len(doc.tables)) if index not in claimed]


def table_index_map_by_bookmark(
    doc, candidates: set[str], *, include_anchors: bool = False
) -> dict[str, int]:
    """书签名 → 表格序号；可兼容表内书签与旧式标题段落书签。"""
    if not candidates:
        return {}
    result: dict[str, int] = {}
    for index, tab in enumerate(doc.tables):
        for el in tab._tbl.iter(qn("w:bookmarkStart")):
            name = el.get(qn("w:name"))
            if name in candidates:
                result.setdefault(name, index)
    if include_anchors:
        for name, index in pair_anchor_bookmarks_with_tables(doc, candidates).items():
            result.setdefault(name, index)
    return result


def _ensure_cell_paragraph(tc, last: bool):
    """取单元格里的第一个/最后一个段落，没有就补一个。"""
    paragraphs = tc.findall(qn("w:p"))
    if paragraphs:
        return paragraphs[-1] if last else paragraphs[0]
    paragraph = OxmlElement("w:p")
    tc.append(paragraph)
    return paragraph


def insert_table_bookmark(tbl, name: str, bookmark_id: int) -> bool:
    """把书签**整表锚定**：起点在首行首格首段开头，终点在末行末格末段末尾。

    这样从 Excel 点“→定位Word表”会直接选中整张表格，而不是只停在标题行。
    """
    rows = tbl.findall(qn("w:tr"))
    if not rows:
        return False
    first_tc = rows[0].find(qn("w:tc"))
    last_row_cells = rows[-1].findall(qn("w:tc"))
    if first_tc is None or not last_row_cells:
        return False
    last_tc = last_row_cells[-1]

    start_paragraph = _ensure_cell_paragraph(first_tc, last=False)
    end_paragraph = _ensure_cell_paragraph(last_tc, last=True)
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    pPr = start_paragraph.find(qn("w:pPr"))
    if pPr is not None:
        pPr.addnext(start)
    else:
        start_paragraph.insert(0, start)
    end_paragraph.append(end)
    return True


def next_bookmark_id(doc) -> int:
    ids = used_bookmark_ids(doc)
    return (max(ids) + 1) if ids else 1


def new_cpah_name(existing: set[str]) -> str:
    """生成 Table_ + 4 位大写字母/数字随机书签名（如 Table_3K9Q）。"""
    alphabet = string.ascii_uppercase + string.digits
    while True:
        name = BOOKMARK_PREFIX + "".join(secrets.choice(alphabet) for _ in range(4))
        if name not in existing:
            existing.add(name)
            return name


def remove_bookmark(doc, name: str) -> bool:
    body = doc.element.body
    removed = False
    for start in list(body.iter(qn("w:bookmarkStart"))):
        if start.get(qn("w:name")) != name:
            continue
        bookmark_id = start.get(qn("w:id"))
        start.getparent().remove(start)
        removed = True
        if bookmark_id is not None:
            for end in list(body.iter(qn("w:bookmarkEnd"))):
                if end.get(qn("w:id")) == bookmark_id:
                    end.getparent().remove(end)
    return removed


def is_tool_bookmark(name: str) -> bool:
    """判断是否本工具写入的书签（Table_ / 旧 CPAH_ / 旧 TBL_ / 标注 MRK_ 前缀）。

    Word 自带书签（如 ``_Toc…`` 目录书签）不在此列，清理时保留不动。
    """
    return (
        name.startswith(BOOKMARK_PREFIX)
        or name.startswith("CPAH_")
        or bool(LEGACY_BOOKMARK_RE.match(name))
        or name.startswith("MRK_")
    )


def clear_tool_bookmarks(doc) -> dict:
    """删除文档里所有本工具书签（start 与配对的 end 一起移除）。

    返回 ``{"removed": int, "kept": int, "names": [...]}``：removed 为本工具书签数，
    kept 为保留的 Word 自带书签数。
    """
    body = doc.element.body
    names = bookmark_names(body)
    tool_names = sorted(n for n in names if is_tool_bookmark(n))
    kept = len(names) - len(tool_names)
    removed = 0
    for name in tool_names:
        if remove_bookmark(doc, name):
            removed += 1
    return {"removed": removed, "kept": kept, "names": tool_names}


def bookmark_starts_in(p_el) -> list[str]:
    return [b.get(qn("w:name")) for b in p_el.iter(qn("w:bookmarkStart")) if b.get(qn("w:name"))]


def _remove_link_fields(p_el) -> int:
    """删除段落里由本工具写入的 HYPERLINK 域（Word 会把超链接规范化为域）。"""
    removed = 0
    runs = list(p_el.findall(qn("w:r")))
    index = 0
    while index < len(runs):
        fld = runs[index].find(qn("w:fldChar"))
        if fld is None or fld.get(qn("w:fldCharType")) != "begin":
            index += 1
            continue
        depth = 0
        end = index
        instr_parts: list[str] = []
        display_parts: list[str] = []
        separated = False
        while end < len(runs):
            inner = runs[end].find(qn("w:fldChar"))
            if inner is not None:
                fld_type = inner.get(qn("w:fldCharType"))
                if fld_type == "begin":
                    depth += 1
                elif fld_type == "end":
                    depth -= 1
                elif fld_type == "separate":
                    separated = True
            for node in runs[end].findall(qn("w:instrText")):
                instr_parts.append(node.text or "")
            if separated:
                for node in runs[end].findall(qn("w:t")):
                    display_parts.append(node.text or "")
            if depth == 0 and end > index:
                break
            end += 1
        instr = "".join(instr_parts)
        display = "".join(display_parts)
        ours = (
            EXCEL_LINK_TEXT in display
            or "_表格数据.xlsx" in instr
            or '\\l "EX_T' in instr
        )
        if ours and "HYPERLINK" in instr.upper():
            for run in runs[index : end + 1]:
                p_el.remove(run)
            removed += 1
            runs = list(p_el.findall(qn("w:r")))
            continue
        index = end + 1
    return removed


def _strip_marker_from_paragraph(p_el) -> int:
    """删除段落文本中的〔→Excel〕（可能被 Word 拆成多个 run），其余文字与格式不动。"""
    removed = 0
    while True:
        runs = list(p_el.findall(qn("w:r")))
        full = "".join("".join(t.text or "" for t in run.findall(qn("w:t"))) for run in runs)
        start = full.find(EXCEL_LINK_TEXT)
        if start < 0:
            return removed
        end = start + len(EXCEL_LINK_TEXT)
        removed += 1
        offset = 0
        for run in runs:
            texts = run.findall(qn("w:t"))
            run_text = "".join(t.text or "" for t in texts)
            run_start, run_end = offset, offset + len(run_text)
            offset = run_end
            if run_end <= start or run_start >= end:
                continue
            new_text = run_text[: max(0, start - run_start)] + run_text[max(0, end - run_start) :]
            if new_text == "" and run.find(qn("w:drawing")) is None:
                p_el.remove(run)
                continue
            for idx, node in enumerate(texts):
                node.text = new_text if idx == 0 else ""
                node.set(qn("xml:space"), "preserve")


def cleanup_excel_markers(doc) -> int:
    """删除 Word 里遗留的〔→Excel〕文字与超链接域，段落其余内容与格式不动。"""
    removed = 0
    for p_el in doc.element.body.iter(qn("w:p")):
        for hyperlink in list(p_el.findall(qn("w:hyperlink"))):
            text = "".join(t.text or "" for t in hyperlink.iter(qn("w:t")))
            if EXCEL_LINK_TEXT in text:
                p_el.remove(hyperlink)
                removed += 1
        removed += _remove_link_fields(p_el)
        removed += _strip_marker_from_paragraph(p_el)
    return removed
