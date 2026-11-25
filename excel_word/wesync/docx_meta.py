"""Word 文档语义层：标题层级（outlineLvl/样式）、表名与章节识别、表块信息汇总、同名表去重。"""

from docx.oxml.ns import qn

from .config import BOOKMARK_PREFIX, DATE_RE, EXCEL_LINK_TEXT
from .docx_bookmarks import bookmark_starts_in, first_bookmark_in_table
from .docx_grid import build_cellmap, compute_merges, header_rows, para_text, table_borders, text_grid


def clean_paragraph_text(text: str) -> str:
    text = text.replace(EXCEL_LINK_TEXT, "")
    return text.strip()


def style_name_of(paragraph) -> str:
    try:
        return paragraph.style.name or ""
    except Exception:
        return ""


def is_heading_style(name: str) -> bool:
    return ("标题" in name) or ("Heading" in name) or name.lower().startswith("heading")


def paragraph_outline_level(p_el) -> int | None:
    """审计软件生成的文档常直接用 w:outlineLvl 标记标题层级。"""
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return None
    lvl = pPr.find(qn("w:outlineLvl"))
    if lvl is None:
        return None
    value = lvl.get(qn("w:val"))
    if value is not None and value.lstrip("-").isdigit():
        return int(value)
    return None


def is_heading_paragraph(p) -> bool:
    return paragraph_outline_level(p._p) is not None or is_heading_style(style_name_of(p))


def heading_level(p) -> int | None:
    """标题层级；仅样式标为标题时按最低优先级处理。"""
    level = paragraph_outline_level(p._p)
    if level is not None:
        return level
    if is_heading_style(style_name_of(p)):
        return 99
    return None


def table_context(doc, items, table_index: int) -> tuple[object, str, str, str]:
    """返回 (书签锚点段落, 表名, 章节名, 日期提示)。"""
    paragraphs_before = []
    seen_tables = 0
    for kind, obj in items:
        if kind == "tbl":
            if seen_tables == table_index:
                break
            seen_tables += 1
        elif kind == "p":
            paragraphs_before.append(obj)

    name = ""
    name_index = None
    for idx in range(len(paragraphs_before) - 1, -1, -1):
        text = clean_paragraph_text(para_text(paragraphs_before[idx]._p))
        if text == "" or DATE_RE.match(text):
            continue
        name = text
        name_index = idx
        break

    # 书签与“回到 Excel”链接锚定在表格前最近的非空段落（多为表格标题行或日期行），
    # 这样每张表各有一个独占锚点，避免同一段落被多张表共用。
    anchor = None
    for p in reversed(paragraphs_before):
        if clean_paragraph_text(para_text(p._p)) != "":
            anchor = p
            break

    # 章节取“层级高于表名所在标题”的最近标题，找不到时退化为最近的标题段落。
    name_level = paragraph_outline_level(paragraphs_before[name_index]._p) if name_index is not None else None
    section = ""
    start = (name_index - 1) if name_index is not None else (len(paragraphs_before) - 1)
    nearest_heading = ""
    for idx in range(start, -1, -1):
        p = paragraphs_before[idx]
        text = clean_paragraph_text(para_text(p._p))
        if text == "" or DATE_RE.match(text):
            continue
        level = heading_level(p)
        if level is None:
            continue
        if not nearest_heading:
            nearest_heading = text
        if name_level is not None and level < name_level:
            section = text
            break
    if not section:
        section = nearest_heading

    date_hint = ""
    for p in reversed(paragraphs_before):
        text = clean_paragraph_text(para_text(p._p))
        if DATE_RE.match(text):
            date_hint = text
            break
    return anchor, name, section, date_hint


def collect_table_info(doc, items, index: int, tab) -> dict:
    tbl = tab._tbl
    trs, cellmap, ncols = build_cellmap(tbl)
    merges = compute_merges(cellmap)
    borders = table_borders(tbl)
    grid = text_grid(tbl, cellmap)
    anchors, name, section, date_hint = table_context(doc, items, index)
    code = f"T{index + 1:02d}"
    # 锚点段落上已有的书签（导出/书签命令会据此复用；这里不做臆测命名）
    anchor_names = bookmark_starts_in(anchors._p) if anchors is not None else []
    own_bookmark = next((n for n in anchor_names if n.startswith(BOOKMARK_PREFIX)), "")
    return {
        "index": index,
        "code": code,
        "tbl": tbl,
        "nrows": len(trs),
        "ncols": ncols,
        "cellmap": cellmap,
        "text": grid,
        "merges": merges,
        "borders": borders,
        "header_rows": header_rows(trs),
        "bookmark": own_bookmark,
        "old_bookmark": first_bookmark_in_table(tbl),
        "name": name or f"表{index + 1}",
        "section": section or "（未识别章节）",
        "date_hint": date_hint,
        "anchor_p": anchors._p if anchors is not None else None,
        "anchor_paragraph": anchors,
    }


def dedupe_table_names(infos: list[dict]) -> None:
    """同章节下多张表名称重复时，用日期或表号区分，便于在 Excel 中辨认。"""
    seen: set[str] = set()
    for info in infos:
        name = info["name"]
        if name in seen:
            hint = info["date_hint"]
            candidate = f"{name}（{hint}）" if hint else f"{name}（表{info['index'] + 1}）"
            if candidate in seen:
                candidate = f"{name}（表{info['index'] + 1}）"
            info["name"] = candidate
        seen.add(info["name"])
