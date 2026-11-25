"""空白表识别与删除：整表全空 / 仅表头有字的表格列为候选，确认后可连表标题一起删除。

判定口径（保守，宁可漏报不误报）：
- 首列（多为行次/项目名称）不参与判断；单列表除外，首列即全部内容；
- 整表全空：除首列外所有单元格都没有文字（空白字符不算文字）；
- 仅表头有字：有文字的行全部是表头行（w:tblHeader 标记，未标记时按首行），
  且表头之外至少还有一行数据区、数据区全空。
数据区只有 —、0 等占位内容的表**不算**空白。
"""

from __future__ import annotations

import re

from docx.oxml.ns import qn

from .config import DATE_RE
from .docx_bookmarks import is_tool_bookmark
from .docx_grid import body_items, build_cellmap, header_rows, para_text, text_grid
from .docx_meta import clean_paragraph_text

# 章节大标题（一、/第X节 等）不是表标题，删表时不应连带删除
SECTION_RE = re.compile(r"^(第[一二三四五六七八九十百\d]+[节章部分]|[一二三四五六七八九十]+、)")

CAPTION_MAX_LEN = 60
SENTENCE_ENDS = "。；！？;!?"


def _looks_like_caption(text: str) -> bool:
    if not text or len(text) > CAPTION_MAX_LEN:
        return False
    if SECTION_RE.match(text):
        return False
    return text[-1] not in SENTENCE_ENDS


def _caption_before(items: list, pos: int):
    """从表格位置向前找表标题：跳过空段与日期行，遇到正文句/另一张表即止。

    返回 (标题段落元素或 None, 标题文字, 标题与表格之间的空段/日期段列表)。
    """
    extras: list = []
    idx = pos - 1
    while idx >= 0:
        kind, obj = items[idx]
        if kind != "p":
            break
        text = clean_paragraph_text(para_text(obj._p))
        if text == "" or DATE_RE.match(text):
            extras.append(obj._p)
            idx -= 1
            continue
        if _looks_like_caption(text):
            return obj._p, text, extras
        break
    return None, "", []


def _blank_kind(trs, grid) -> str | None:
    nrows = len(trs)
    ncols = max((len(row) for row in grid), default=0)
    # 首列多为行次/项目名称，不参与空白判断；单列表除外（首列即全部内容）
    start_col = 1 if ncols >= 2 else 0
    filled = [
        (r, c)
        for r, row in enumerate(grid)
        for c, text in enumerate(row)
        if c >= start_col and text is not None and text.strip()
    ]
    if not filled:
        return "整表全空"
    heads = header_rows(trs) or {0}
    if nrows > max(heads) + 1 and all(r in heads for r, _ in filled):
        return "仅表头有字"
    return None


def find_blank_tables(doc) -> list[dict]:
    """扫描全文，返回空白表候选清单（只读，不改文档）。"""
    items = body_items(doc)
    candidates: list[dict] = []
    table_index = -1
    for pos, (kind, obj) in enumerate(items):
        if kind != "tbl":
            continue
        table_index += 1
        tbl = obj._tbl
        trs, cellmap, ncols = build_cellmap(tbl)
        grid = text_grid(tbl, cellmap)
        blank_kind = _blank_kind(trs, grid)
        if blank_kind is None:
            continue
        caption_p, caption_text, extra_ps = _caption_before(items, pos)
        candidates.append(
            {
                "no": len(candidates) + 1,
                "index": table_index,
                "code": f"T{table_index + 1:02d}",
                "kind": blank_kind,
                "nrows": len(trs),
                "ncols": ncols,
                "tbl": tbl,
                "caption_p": caption_p,
                "caption_text": caption_text,
                "extra_ps": extra_ps,
            }
        )
    return candidates


def _tool_bookmark_ids(doc) -> set[str]:
    ids = set()
    for el in doc.element.body.iter(qn("w:bookmarkStart")):
        name = el.get(qn("w:name")) or ""
        if is_tool_bookmark(name):
            ids.add(el.get(qn("w:id")))
    return ids


def _cleanup_orphan_bookmarks(doc, tool_ids: set[str]) -> int:
    """删表/删标题后，清掉失去配对的本工具书签端点（Word 自带书签不动）。"""
    body = doc.element.body
    start_ids = {el.get(qn("w:id")) for el in body.iter(qn("w:bookmarkStart"))}
    end_ids = {el.get(qn("w:id")) for el in body.iter(qn("w:bookmarkEnd"))}
    removed = 0
    for el in list(body.iter(qn("w:bookmarkStart"))):
        if el.get(qn("w:id")) in tool_ids and el.get(qn("w:id")) not in end_ids:
            el.getparent().remove(el)
            removed += 1
    for el in list(body.iter(qn("w:bookmarkEnd"))):
        if el.get(qn("w:id")) in tool_ids and el.get(qn("w:id")) not in start_ids:
            el.getparent().remove(el)
            removed += 1
    return removed


def delete_blank_tables(doc, selected: list[tuple[dict, bool]]) -> dict:
    """按选择删除空白表。selected 为 (候选, 是否连表标题一起删) 列表。

    返回 {"tables": 删表数, "captions": 删标题数, "bookmarks": 清理书签端点数}。
    """
    tool_ids = _tool_bookmark_ids(doc)
    removed_tables = 0
    removed_captions = 0
    for cand, with_caption in selected:
        tbl = cand["tbl"]
        if tbl.getparent() is not None:
            tbl.getparent().remove(tbl)
            removed_tables += 1
        if with_caption and cand["caption_p"] is not None:
            for p_el in [cand["caption_p"], *cand["extra_ps"]]:
                if p_el.getparent() is not None:
                    p_el.getparent().remove(p_el)
            removed_captions += 1
    removed_bookmarks = _cleanup_orphan_bookmarks(doc, tool_ids)
    return {"tables": removed_tables, "captions": removed_captions, "bookmarks": removed_bookmarks}