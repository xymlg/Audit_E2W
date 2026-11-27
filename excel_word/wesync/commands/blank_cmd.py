"""blank 子命令：删除 Word 里的空白表格（整表全空 / 仅表头有字），可连表标题一起删。

默认只扫描列出候选，不删任何内容；用 --pick 1,3 或 --all 明确选择后才删除。
正式删除前自动备份原文件到 _backup。
"""

from __future__ import annotations

from copy import copy

import docx

from ..config import DELETED_BOOKMARK, NEW_BOOKMARK, SHEET_NAME, SHEET_NEW_TABLES
from ..console import log
from ..docx_blank import delete_blank_tables, find_blank_tables
from ..docx_bookmarks import table_index_map_by_bookmark
from ..errors import SyncError
from ..excel_blocks import parse_blocks
from ..fsops import backup_file, check_file_free, load_workbook_for_save, resolve_docx, resolve_xlsx, save_docx_xlsx_transaction


def _parse_pick(text: str, count: int) -> list[int]:
    nos: list[int] = []
    for part in text.split(","):
        part = part.strip()
        if not part:
            continue
        if not part.isdigit() or not 1 <= int(part) <= count:
            raise SyncError(f"候选编号无效：{part}（应为 1～{count} 的数字，逗号分隔）")
        no = int(part)
        if no not in nos:
            nos.append(no)
    if not nos:
        raise SyncError("--pick 未给出有效编号，例如 --pick 1,3")
    return nos


def _load_all_blocks(wb):
    order, blocks = parse_blocks(wb[SHEET_NAME])
    base_codes = set(order)
    for sheet_name in (SHEET_NEW_TABLES,):
        if sheet_name in wb.sheetnames:
            extra_order, extra_blocks = parse_blocks(wb[sheet_name])
            order.extend(extra_order)
            blocks.update(extra_blocks)
    for code, block in blocks.items():
        block["sheet_name"] = SHEET_NAME if code in base_codes else SHEET_NEW_TABLES
        block["_ws"] = wb[block["sheet_name"]]
    return order, blocks


def _plan_reversible(doc, xlsx_path, selected, force=False):
    """可逆删除前的准备：找到每个候选对应的工作簿表块与原位置锚点。"""
    if not xlsx_path.exists():
        raise SyncError("可逆删除需要工作簿来保留表块：请先执行「导出表格到 Excel」生成工作簿。")
    check_file_free(xlsx_path, force)
    wb = load_workbook_for_save(xlsx_path)
    if SHEET_NAME not in wb.sheetnames:
        raise SyncError(f"工作簿中找不到工作表“{SHEET_NAME}”，请先重新执行「导出表格到 Excel」。")
    _, blocks = _load_all_blocks(wb)
    name_to_code = {
        block["bookmark"]: code
        for code, block in blocks.items()
        if block.get("bookmark") and block["bookmark"] not in (NEW_BOOKMARK, DELETED_BOOKMARK)
    }
    index_map = table_index_map_by_bookmark(doc, set(name_to_code))
    index_to_name = {index: name for name, index in index_map.items()}
    sel_indices = {cand["index"] for cand, _ in selected}
    plans = []
    for cand, with_caption in selected:
        code = name_to_code.get(index_to_name.get(cand["index"], ""))
        if not code:
            raise SyncError(
                f"候选{cand['no']} 在工作簿中找不到对应表块（可能未导出或已重建），请先重新执行「导出表格到 Excel」。"
            )
        prev_name = ""
        for i in range(cand["index"] - 1, -1, -1):
            if i not in sel_indices and i in index_to_name:
                prev_name = index_to_name[i]
                break
        plans.append((cand, with_caption, code, prev_name))
    return wb, blocks, plans


def _mark_deleted_blocks(blocks, plans) -> None:
    """Word 删表后，把工作簿对应表块标记为「已删除」并记录原位置与原标题。"""
    for cand, with_caption, code, prev_name in plans:
        block = blocks[code]
        ws = block["_ws"]
        row = block["anchor_row"]
        col0 = block.get("col0", 2)
        ws.cell(row, col0 + 1, f"书签：{DELETED_BOOKMARK}")
        link = ws.cell(row, col0 + 2)
        link.value = None
        link.hyperlink = None
        ws.cell(row, col0 + 4, f"原位置：{prev_name}之后" if prev_name else "原位置：文档末尾")
        caption = cand["caption_text"] if with_caption and cand["caption_p"] is not None else ""
        ws.cell(row, col0 + 5, f"原标题：{caption}" if caption else None)
        anchor_style = ws.cell(row, col0 + 1)._style
        for col in (col0 + 4, col0 + 5):
            ws.cell(row, col)._style = copy(anchor_style)


def cmd_blank_tables(args) -> int:
    if getattr(args, "restore", None):
        return cmd_restore_tables(args)
    docx_path = resolve_docx(args.docx)
    doc = docx.Document(str(docx_path))
    candidates = find_blank_tables(doc)

    log(f"读取 Word：{docx_path.name}")
    if not candidates:
        log("[空白表] 未发现空白表格（整表全空或仅表头有字），无需处理")
        return 0
    log(f"[空白表] 发现候选空白表：{len(candidates)} 张")
    for cand in candidates:
        caption = f"，表标题：{cand['caption_text']}" if cand["caption_text"] else "，未识别到表标题"
        log(f"  [候选{cand['no']}] {cand['code']}（{cand['kind']}，{cand['nrows']}行×{cand['ncols']}列）{caption}")

    if args.dry_run:
        log("预演模式（--dry-run）：未写入任何文件。删除请用 --pick 编号 或 --all。")
        return 0

    pick_text = getattr(args, "pick", None)
    if getattr(args, "all", False):
        nos = [cand["no"] for cand in candidates]
    elif pick_text:
        nos = _parse_pick(pick_text, len(candidates))
    else:
        log("[空白表] 未指定要删除的候选，本次未改动文件。")
        log("  删除示例：--pick 1,3（按候选编号）或 --all（全部）；只删表格保留标题加 --keep-caption")
        return 0

    keep_all = bool(getattr(args, "keep_caption", False))
    keep_nos = set()
    keep_for = getattr(args, "keep_caption_for", None)
    if keep_for:
        keep_nos = set(_parse_pick(keep_for, len(candidates)))
    selected = [
        (cand, not keep_all and cand["no"] not in keep_nos and cand["caption_p"] is not None)
        for cand in candidates
        if cand["no"] in nos
    ]

    check_file_free(docx_path, args.force)
    reversible = bool(getattr(args, "reversible", False))
    xlsx_path = resolve_xlsx(docx_path, getattr(args, "excel", None))
    prepared = _plan_reversible(doc, xlsx_path, selected, args.force) if reversible else None
    result = delete_blank_tables(doc, selected)
    if reversible:
        wb, blocks, plans = prepared
        _mark_deleted_blocks(blocks, plans)
        backups = save_docx_xlsx_transaction(
            docx_path=docx_path, doc=doc, xlsx_path=xlsx_path, wb=wb,
            backup=getattr(args, "backup", True),
        )
        docx_backup = backups.get(docx_path)
        xlsx_backup = backups.get(xlsx_path)
    else:
        docx_backup = backup_file(docx_path)
        doc.save(str(docx_path))
        xlsx_backup = None
    log(
        f"[空白表] 已删除 {result['tables']} 张空白表、{result['captions']} 个表标题"
        + (f"（原文件已备份：{docx_backup.name}）" if docx_backup else "")
    )
    if result["bookmarks"]:
        log(f"[空白表] 顺带清理失去配对的定位书签端点 {result['bookmarks']} 个")
    if reversible:
        log(
            f"[可逆] 工作簿已标记 {len(plans)} 个表块为「已删除」（内容保留，可随时恢复）"
            + (f"（原工作簿已备份：{xlsx_backup.name}）" if xlsx_backup else "")
        )
        log("[可逆] Excel 表块已保留；要恢复时在「删除空白表格」页勾选恢复，或命令行 --restore 编号。")
    else:
        log("[空白表] 表格有删减，请重新执行「1-导出表格到Excel」重建工作簿，再做同步。")
    return 0


def cmd_restore_tables(args) -> int:
    """恢复已可逆删除的表：按工作簿保留的表块在原位置重建 Word 表格。"""
    from ..docx_bookmarks import bookmark_names, insert_table_bookmark, new_cpah_name, next_bookmark_id
    from ..docx_grid import body_items
    from ..docx_meta import collect_table_info
    from ..docx_newtable import available_width_dxa, build_table, col_widths_from_excel, insert_new_table
    from ..excel_blocks import block_merges, refresh_anchor_metadata, refresh_defined_names
    from ..valuefmt import formula_map
    from .writeback_cmd import _block_label, _new_table_grid

    docx_path = resolve_docx(args.docx)
    xlsx_path = resolve_xlsx(docx_path, getattr(args, "excel", None))
    if not xlsx_path.exists():
        raise SyncError(f"找不到工作簿：{xlsx_path}\n恢复需要工作簿里保留的表块。")
    doc = docx.Document(str(docx_path))
    wb = load_workbook_for_save(xlsx_path)
    if SHEET_NAME not in wb.sheetnames:
        raise SyncError(f"工作簿中找不到工作表“{SHEET_NAME}”，请确认该文件由本工具导出")
    order, blocks = _load_all_blocks(wb)
    deleted = [code for code in order if blocks[code].get("bookmark") == DELETED_BOOKMARK]
    log(f"读取 Word：{docx_path.name}")
    if not deleted:
        log("[恢复] 工作簿中没有「已删除」标记的表块，无需恢复")
        return 0
    log(f"[恢复] 可恢复的表块：{len(deleted)} 个")
    for no, code in enumerate(deleted, 1):
        block = blocks[code]
        meta = block.get("meta", {})
        pos = meta.get("restore_after") or "末尾"
        log(f"  [{no}] {code} {_block_label(block)}（{len(block['grid_rows'])}行×{block['ncols']}列，原位置：{pos}）")

    restore_arg = str(args.restore).strip()
    if restore_arg.lower() == "all":
        nos = list(range(1, len(deleted) + 1))
    else:
        nos = _parse_pick(restore_arg, len(deleted))
    if args.dry_run:
        log(f"预演模式（--dry-run）：将恢复 {len(nos)} 张表，未写入任何文件。")
        return 0

    check_file_free(docx_path, args.force)
    check_file_free(xlsx_path, args.force)
    cached = formula_map(xlsx_path)
    existing_names = bookmark_names(doc.element.body)
    bookmark_id = next_bookmark_id(doc)
    total_dxa = available_width_dxa(doc)
    real_names = {
        blocks[c]["bookmark"] for c in order
        if blocks[c].get("bookmark") and blocks[c]["bookmark"] not in (NEW_BOOKMARK, DELETED_BOOKMARK)
    }
    index_map = table_index_map_by_bookmark(doc, real_names)
    restored: dict[str, tuple] = {}
    for no in nos:
        code = deleted[no - 1]
        block = blocks[code]
        if not block["grid_rows"]:
            raise SyncError(f"{code}：表块没有网格行，无法恢复")
        meta = block.get("meta", {})
        grid = _new_table_grid(block["_ws"], block, cached)
        merges = block_merges(block["_ws"], block)
        widths = col_widths_from_excel(block["_ws"], block["ncols"], total_dxa, col0=block.get("col0", 2))
        tbl = build_table(grid, merges, block["ncols"], widths)
        anchor = meta.get("restore_after") or ""
        prev_tbl = None
        if anchor and anchor != "末尾" and anchor in index_map:
            prev_tbl = doc.tables[index_map[anchor]]._tbl
        elif anchor and anchor != "末尾":
            log(f"[提醒] {code} 原位置的锚点表已不在 Word 中，恢复到文档末尾")
        caption = meta.get("caption") or _block_label(block, limit=60) or code
        insert_new_table(doc, prev_tbl, tbl, caption)
        bookmark = new_cpah_name(existing_names)
        existing_names.add(bookmark)
        insert_table_bookmark(tbl, bookmark, bookmark_id)
        bookmark_id += 1
        restored[code] = (bookmark, tbl)
        log(f"[恢复] {code} {caption}：已重建（{len(grid)}行×{block['ncols']}列，书签 {bookmark}）")

    items = body_items(doc)
    infos = [collect_table_info(doc, items, idx, tab) for idx, tab in enumerate(doc.tables)]
    for code, (bookmark, tbl) in restored.items():
        index = next(i for i, tab in enumerate(doc.tables) if tab._tbl is tbl)
        info = {**infos[index], "bookmark": bookmark}
        refresh_anchor_metadata(
            blocks[code]["_ws"], blocks[code], info, docx_path.name,
            with_link=True, ncols=blocks[code]["ncols"]
        )
    refresh_defined_names(wb, {code: blocks[code] for code in restored})

    backups = save_docx_xlsx_transaction(
        docx_path=docx_path, doc=doc, xlsx_path=xlsx_path, wb=wb,
        backup=getattr(args, "backup", True),
    )
    docx_backup = backups.get(docx_path)
    xlsx_backup = backups.get(xlsx_path)
    log(f"[恢复] 已保存 {docx_path.name}" + (f"（原文件已备份：{docx_backup.name}）" if docx_backup else ""))
    log(f"[恢复] 已保存 {xlsx_path.name}（表块标记已改回真实书签）" + (f"（原工作簿已备份：{xlsx_backup.name}）" if xlsx_backup else ""))
    return 0