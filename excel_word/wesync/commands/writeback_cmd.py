"""writeback 子命令：按书签把工作簿的表块回写到 Word 表格。

匹配方式：工作簿每个表块 C 列记录的书签名 ↔ Word 里的表格（锚点段落中的书签 + 紧随其后的表格）。
Word 中已删除的表格（书签找不到对应表格）会被**跳过**并记日志，其余表格照常同步。
"""

import re
from collections import defaultdict
from copy import deepcopy

import docx
from docx.oxml.ns import qn
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

from ..config import DELETED_BOOKMARK, EXCEL_LINK_TEXT, FILL_CODE_COLUMN, FONT_NAME, FONT_SIZE, MARK_DELETED, MARK_NO, NEW_BOOKMARK, SHEET_NAME
from ..console import flush_log_summary, log, log_change, warn
from ..align import align_blocks_to_tables
from ..changelog import export_changelog
from ..docx_bookmarks import (
    bookmark_names,
    cleanup_excel_markers,
    insert_table_bookmark,
    new_cpah_name,
    next_bookmark_id,
    table_index_map_by_bookmark,
    unclaimed_table_indices,
)
from ..docx_columns import sync_columns
from ..docx_edit import set_cell_text
from ..docx_grid import body_items, build_cellmap, para_first_run_rpr, text_grid
from ..docx_meta import collect_table_info, dedupe_table_names
from ..docx_newtable import available_width_dxa, build_table, col_widths_from_excel, insert_new_table
from ..docx_rows import align_rows
from ..errors import SyncError
from ..excel_blocks import block_merges, parse_blocks, refresh_anchor_metadata, refresh_defined_names
from ..fsops import check_file_free, load_workbook_for_save, resolve_docx, resolve_xlsx, save_docx_xlsx_transaction
from ..sort_rules import apply_block_sort
from ..valuefmt import display_text, evaluate_formula, formula_map


def _block_label(block: dict, limit: int = 30) -> str:
    """取表块的简短名称（去掉【表N】前缀与章节后缀并截断），用于日志。"""
    text = (block.get("meta", {}).get("name") or "").strip()
    if not text:
        return ""
    text = re.sub(r"^【(表\d+|新表)】", "", text)
    text = re.sub(r"\s+", " ", text.split(" · 章节：")[0]).strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


def _new_table_grid(ws, block: dict, cached) -> list[list[str]]:
    """新表块的网格显示文本（公式取缓存或轻量求值，与回写口径一致）。"""
    grid = []
    col0 = block.get("col0", 2)
    for xlsx_row in block["grid_rows"]:
        row_texts = []
        for c in range(block["ncols"]):
            cell = ws.cell(xlsx_row, col0 + c)
            raw = cell.value
            if isinstance(raw, str) and raw.startswith("="):
                value = cached.get((xlsx_row, col0 + c))
                if value is None:
                    value = evaluate_formula(ws, xlsx_row, col0 + c, cached)
                if value is None:
                    raise SyncError(
                        f"{block['code']} 新表第{len(grid) + 1}行第{c + 1}列是公式且无法计算结果"
                        f"（支持 SUM 与简单四则运算）。请先在 Excel 中打开并保存该文件后再回写"
                    )
                raw = value
            row_texts.append(display_text(raw, cell.number_format))
        grid.append(row_texts)
    return grid


def _prev_anchor_tbl(order, code, recorded, table_index, doc, created):
    """新表插入点：工作簿中它之前最近的、能定位到 Word 表格的块。"""
    for prior in reversed(order[: order.index(code)]):
        if prior in created:
            return created[prior], prior
        bookmark = recorded.get(prior, "")
        if not bookmark or bookmark in (NEW_BOOKMARK, DELETED_BOOKMARK):
            continue
        index = table_index.get(bookmark)
        if index is not None:
            return doc.tables[index]._tbl, prior
    return None, "文档末尾"


def _new_block_anchor(order, code, block, recorded, table_index, doc, created):
    """新表插入点：优先用表块记录的指定位置（F 列「位置：T0x之后」），否则按工作簿顺序。"""
    after = (block.get("meta", {}).get("after") or "").strip()
    if after and after != "末尾":
        if after in created:
            return created[after], after
        bookmark = recorded.get(after, "")
        if bookmark and bookmark not in (NEW_BOOKMARK, DELETED_BOOKMARK):
            index = table_index.get(bookmark)
            if index is not None:
                return doc.tables[index]._tbl, after
        log(f"[提醒] {code} 指定的插入位置 {after} 在 Word 中找不到对应表格，改按工作簿顺序插入")
    return _prev_anchor_tbl(order, code, recorded, table_index, doc, created)


def _create_new_tables(doc, ws, order, blocks, recorded, table_index, cached) -> int:
    """按「书签：新增」标记在 Word 里建表并登记书签，返回创建数。"""
    new_codes = [code for code in order if recorded.get(code) == NEW_BOOKMARK]
    if not new_codes:
        return 0
    existing_names = bookmark_names(doc.element.body)
    bookmark_id = next_bookmark_id(doc)
    total_dxa = available_width_dxa(doc)
    created: dict[str, object] = {}
    for code in new_codes:
        block = blocks[code]
        if not block["grid_rows"]:
            raise SyncError(f"{code}：新表块没有网格行，请在 Excel 里填好表格内容后再同步")
        grid = _new_table_grid(ws, block, cached)
        merges = block_merges(ws, block)
        widths = col_widths_from_excel(ws, block["ncols"], total_dxa, block.get("col0", 2))
        tbl = build_table(grid, merges, block["ncols"], widths)
        prev_tbl, prev_label = _new_block_anchor(order, code, block, recorded, table_index, doc, created)
        name = _block_label(block, limit=60) or "新表"
        insert_new_table(doc, prev_tbl, tbl, name)
        bookmark = new_cpah_name(existing_names)
        insert_table_bookmark(tbl, bookmark, bookmark_id)
        bookmark_id += 1
        recorded[code] = bookmark
        created[code] = tbl
        log(
            f"[新增] {code} {name}：已在 {prev_label} 之后创建"
            f"（{len(grid)}行×{block['ncols']}列，合并区 {len(merges)} 处，书签 {bookmark}）"
        )
    return len(created)


def _sync_table(code: str, block: dict, tbl, ws, cached, recorder=None, name: str = "") -> dict:
    """同步一张表（列结构 → 行结构 → 单元格内容），返回改动统计。

    recorder 不为 None 时，把逐格改动与结构改动收集为留痕记录（供导出清单/差异表）。
    """
    _, cellmap_before, _ = build_cellmap(tbl)
    old_grid = text_grid(tbl, cellmap_before)

    include, inserted_cols, deleted_cols = sync_columns(tbl, ws, block, code, cached)
    column_changed = include != list(range(block["ncols"]))

    delta, deleted_rows = align_rows(tbl, len(block["grid_rows"]))
    if delta:
        log(f"[结构] {code}：{'新增' if delta > 0 else '删除'} {abs(delta)} 行（合计行保持在末行）")
    if deleted_rows:
        log(f"[结构] {code}：被删除行的首列内容 → {'、'.join(deleted_rows)}")
    if recorder is not None:
        if delta:
            recorder.append({"code": code, "name": name, "kind": "增行" if delta > 0 else "删行",
                             "item": f"{'新增' if delta > 0 else '删除'} {abs(delta)} 行", "old": "", "new": ""})
        if deleted_rows:
            recorder.append({"code": code, "name": name, "kind": "删行明细",
                             "item": "被删行首列：" + "、".join(deleted_rows), "old": "", "new": ""})
        if column_changed:
            skipped_cols = block["ncols"] - len(include)
            recorder.append({"code": code, "name": name, "kind": "列结构",
                             "item": f"新增 {inserted_cols} 列、删除 {deleted_cols} 列"
                             + (f"、忽略整列空白 {skipped_cols} 个" if skipped_cols else ""),
                             "old": "", "new": ""})

    trs, cellmap, ncols = build_cellmap(tbl)
    if len(trs) != len(block["grid_rows"]):
        raise SyncError(f"{code}：行数对齐失败，Word 现有 {len(trs)} 行，Excel 为 {len(block['grid_rows'])} 行")
    if ncols != len(include):
        raise SyncError(f"{code}：列数对齐失败，Word 现有 {ncols} 列，Excel 有效列为 {len(include)} 列")

    col_rpr = {}
    for c in range(ncols):
        for r in range(len(trs)):
            pos = cellmap[r][c]
            if pos is not None and pos.owner:
                rpr = para_first_run_rpr(pos.tc.find(qn("w:p")))
                if rpr is not None:
                    col_rpr[c] = rpr
                    break

    changed_cells = 0
    col0 = block.get("col0", 2)
    for r, xlsx_row in enumerate(block["grid_rows"]):
        for word_col, excel_col in enumerate(include):
            pos = cellmap[r][word_col]
            cell = ws.cell(xlsx_row, col0 + excel_col)
            if cell.hyperlink is not None:
                warn(
                    f"{code} 第{r + 1}行第{excel_col + 1}列带有超链接元数据"
                    f"（常见于用脚本增删行导致错位），已跳过该单元格"
                )
                continue
            raw = cell.value
            if isinstance(raw, str) and raw.startswith("="):
                cached_value = cached.get((xlsx_row, col0 + excel_col))
                if cached_value is None:
                    cached_value = evaluate_formula(ws, xlsx_row, col0 + excel_col, cached)
                if cached_value is None:
                    raise SyncError(
                        f"{code} 第{r + 1}行第{excel_col + 1}列是公式且无法计算结果"
                        f"（支持 SUM 与简单四则运算）。请先在 Excel 中打开并保存该文件后再回写"
                    )
                raw = cached_value
            text = display_text(raw, cell.number_format)
            if pos is None or not pos.owner:
                if text.strip():
                    warn(
                        f"{code} 第{r + 1}行第{excel_col + 1}列在 Excel 中有值，"
                        f"但该位置在 Word 中属合并单元格，已跳过"
                    )
                continue
            old_value = old_grid[r][excel_col] if (r < len(old_grid) and excel_col < len(old_grid[r])) else None
            if old_value is None and r >= len(old_grid):
                old_value = "（新增行）"
            if set_cell_text(pos.tc, text, col_rpr.get(word_col)):
                changed_cells += 1
                old_display = old_value if old_value not in (None, "") else "（空）"
                new_display = text if text != "" else "（空）"
                if not column_changed:
                    log_change(f"[回写] {code} 第{r + 1}行第{excel_col + 1}列：{old_display} → {new_display}")
                if recorder is not None:
                    item = old_grid[r][0] if (r < len(old_grid) and old_grid[r]) else ""
                    recorder.append({"code": code, "name": name, "row": r + 1, "col": excel_col + 1,
                                     "item": item, "old": old_display, "new": new_display, "kind": "单元格"})

    if changed_cells == 0:
        log(f"[回写] {code}：无变化")
    elif column_changed:
        log(f"[回写] {code}：列结构变化，共更新 {changed_cells} 个单元格（逐格明细省略）")
    return {"cells": changed_cells, "rows": delta, "cols": inserted_cols + deleted_cols}


def cmd_writeback(args) -> int:
    docx_path = resolve_docx(args.docx)
    xlsx_path = resolve_xlsx(docx_path, args.excel)
    if not xlsx_path.exists():
        raise SyncError(f"找不到 Excel 文件：{xlsx_path}\n请先执行 export 导出，或检查 --excel 参数。")
    if not args.dry_run:
        check_file_free(docx_path, args.force)
        check_file_free(xlsx_path, args.force)

    doc = docx.Document(str(docx_path))
    wb = load_workbook_for_save(xlsx_path)
    if SHEET_NAME not in wb.sheetnames:
        raise SyncError(f"工作簿中找不到工作表“{SHEET_NAME}”，请确认该文件由本工具导出")
    ws = wb[SHEET_NAME]
    cached = formula_map(xlsx_path)
    recorder: list[dict] = []

    order, blocks = parse_blocks(ws)
    if not order:
        raise SyncError("工作簿里没有识别到任何表块，请确认该文件由本工具导出")

    recorded = {code: (blocks[code].get("bookmark") or "").strip() for code in order}
    if not any(recorded.values()):
        raise SyncError(
            "该工作簿没有书记录，无法按书签匹配 Word 表格，已中止以免写错位置。\n"
            "请先用「导出表格到 Excel」重建工作簿（或确认 --excel 指向的是本工具导出的文件）。"
        )

    sort_changed = 0
    for code in order:
        if not recorded.get(code) or recorded[code] in (NEW_BOOKMARK, DELETED_BOOKMARK):
            continue
        stats = apply_block_sort(ws, blocks[code], cached, dry_run=args.dry_run)
        if stats and not args.dry_run:
            sort_changed += 1

    delete_cross = bool(getattr(args, "delete_empty_rows", False))
    if delete_cross and all(blocks[c].get("col0", 2) == 2 for c in order):
        log("[提示] 该工作簿为旧布局（无 B 列标记），「删除✕行」本次不生效；重新导出后即可使用。")

    new_codes = [code for code in order if recorded.get(code) == NEW_BOOKMARK]
    created_count = 0
    if new_codes:
        real_names = {n for n in recorded.values() if n and n not in (NEW_BOOKMARK, DELETED_BOOKMARK)}
        pre_index = table_index_map_by_bookmark(doc, real_names, include_anchors=True)
        if args.dry_run:
            for code in new_codes:
                block = blocks[code]
                _, prev_label = _new_block_anchor(order, code, block, recorded, pre_index, doc, {})
                log(
                    f"[新增·预演] {code} {_block_label(block)}：将在 {prev_label} 之后创建"
                    f"（{len(block['grid_rows'])}行×{block['ncols']}列）"
                )
        else:
            created_count = _create_new_tables(doc, ws, order, blocks, recorded, pre_index, cached)
            for code in new_codes:
                recorder.append({"code": code, "name": _block_label(blocks[code]), "kind": "新表",
                                 "item": f"新建 {len(blocks[code]['grid_rows'])}行×{blocks[code]['ncols']}列",
                                 "old": "", "new": ""})

    items = body_items(doc)
    infos = [collect_table_info(doc, items, idx, tab) for idx, tab in enumerate(doc.tables)]
    dedupe_table_names(infos)

    candidates = {name for name in recorded.values() if name and name not in (NEW_BOOKMARK, DELETED_BOOKMARK)}
    table_index = table_index_map_by_bookmark(doc, candidates, include_anchors=True)
    unclaimed = unclaimed_table_indices(doc, table_index)

    # 书签名重复时（旧版工具曾把 C 列写脏）：只有内容能对上该表的那一块可以同步，其余跳过
    shares: dict[str, list[str]] = defaultdict(list)
    for code in order:
        if recorded[code] and recorded[code] not in (NEW_BOOKMARK, DELETED_BOOKMARK):
            shares[recorded[code]].append(code)
    align_map, _, _ = align_blocks_to_tables(doc, ws, blocks)
    owner: dict[str, str] = {}
    for name, codes in shares.items():
        if len(codes) == 1:
            owner[name] = codes[0]
            continue
        target = table_index.get(name)
        matched = [code for code in codes if code in align_map and align_map[code] == target]
        if not matched:
            # 内容对不上（例如用户在 Excel 里改过行列表头）：让候选按顺序尝试，
            # 第一个成功同步的认领该表，失败的不认领，避免误写
            continue
        chosen = matched[0]
        owner[name] = chosen
        others = "、".join(code for code in codes if code != chosen)
        log(f"[提示] 书签名 {name} 被 {len(codes)} 个表块共用，按内容归属 {chosen}，其余跳过（{others}）")

    marker_removed = cleanup_excel_markers(doc)
    if marker_removed:
        log(f"[清理] 已删除 Word 正文中遗留的 {marker_removed} 处 {EXCEL_LINK_TEXT} 文字/断链")

    log(f"读取 Word：{docx_path.name}（{len(doc.tables)} 张表）")
    log(f"读取 Excel：{xlsx_path.name}（工作表：{SHEET_NAME}，{len(order)} 个表块）")
    if args.dry_run:
        log("预演模式（--dry-run）：只显示将发生的改动，不写入文件。")

    total_cells = 0
    total_rows = 0
    total_columns = 0
    total_cross = 0
    mark_changed = 0
    synced_codes: list[str] = []
    skipped_codes: list[str] = []
    structure_skips = 0
    claimed_indices: dict[int, str] = {}

    for code in order:
        block = blocks[code]
        bookmark = recorded[code]
        label = _block_label(block)
        if bookmark == NEW_BOOKMARK:
            continue  # 新增块的预演已在前面单独列出
        if bookmark == DELETED_BOOKMARK:
            skipped_codes.append(code)
            log(f"[跳过] {code} {label}：已可逆删除（Excel 保留表块，可在「删除空白表格」页恢复）")
            continue
        if not bookmark:
            skipped_codes.append(code)
            log(f"[跳过] {code} {label}：工作簿中该表块没有书签记录，无法匹配 Word 表格")
            continue
        if owner.get(bookmark, code) != code:
            skipped_codes.append(code)
            log(f"[跳过] {code} {label}：书签名 {bookmark} 已被 {owner[bookmark]} 使用（C 列重复），已跳过")
            continue
        index = table_index.get(bookmark)
        if index is None:
            skipped_codes.append(code)
            log(f"[跳过] {code} {label}：Word 中已找不到该表格（书签 {bookmark}，可能已被删除）")
            continue
        if index in claimed_indices:
            skipped_codes.append(code)
            log(f"[跳过] {code} {label}：第 {index + 1} 张表已被 {claimed_indices[index]} 认领，跳过以免重复写入")
            continue
        tbl = doc.tables[index]._tbl
        tbl_snapshot = deepcopy(tbl)
        recorder_len = len(recorder)
        try:
            sync_block = block
            cross_rows: list[int] = []
            if delete_cross and block.get("col0", 2) == 3:
                cross_rows = [
                    r for r in block["grid_rows"][1:]
                    if str(ws.cell(r, 2).value or "").strip() == MARK_NO
                    and not display_text(ws.cell(r, 3).value, ws.cell(r, 3).number_format).strip().startswith("合计")
                ]
                if cross_rows:
                    sync_block = {**block, "grid_rows": [r for r in block["grid_rows"] if r not in cross_rows]}
                    remaining_body = [
                        r for r in block["grid_rows"][1:]
                        if r not in cross_rows
                        and not display_text(ws.cell(r, 3).value, ws.cell(r, 3).number_format).strip().startswith("合计")
                    ]
                    if not remaining_body:
                        warn(
                            f"{code}：✕行覆盖全部数据行，删除后该表只剩表头/合计；"
                            f"之后改回✓恢复时因无数据行可复制格式会跳过该表，需先在 Word 手工补一行或从备份回滚。"
                        )
                    items_text = "、".join(
                        (display_text(ws.cell(r, 3).value, ws.cell(r, 3).number_format).strip() or "（空行）")
                        for r in cross_rows
                    )
                    log(
                        f"[删除{'·预演' if args.dry_run else ''}] {code}：{len(cross_rows)} 个✕行"
                        f"{'将' if args.dry_run else '已'}从 Word 删除（Excel 保留为「已删」，改回✓可恢复）：{items_text}"
                    )
                    recorder.append({"code": code, "name": label, "kind": "删行(✕)",
                                     "item": f"删除 {len(cross_rows)} 个✕行：" + items_text, "old": "", "new": ""})
            stats = _sync_table(code, sync_block, tbl, ws, cached, recorder=recorder, name=label)
        except SyncError as exc:
            tbl.getparent().replace(tbl, tbl_snapshot)
            del recorder[recorder_len:]
            skipped_codes.append(code)
            structure_skips += 1
            log(f"[跳过] {code} {label}：结构不符，该表已回滚、未留下部分修改（{exc}）")
            continue
        if cross_rows and not args.dry_run:
            for r in cross_rows:
                cell = ws.cell(r, 2)
                cell.value = MARK_DELETED
                cell.font = Font(name=FONT_NAME, size=FONT_SIZE, color="808080")
                cell.fill = FILL_CODE_COLUMN
                cell.alignment = Alignment(horizontal="center", vertical="center")
            mark_changed += len(cross_rows)
            total_cross += len(cross_rows)
        synced_codes.append(code)
        claimed_indices[index] = code
        total_cells += stats["cells"]
        total_rows += abs(stats["rows"])
        total_columns += stats["cols"]

    if getattr(args, "recorder", None) is not None:
        args.recorder.extend(recorder)

    if unclaimed:
        log(f"[提示] Word 中有 {len(unclaimed)} 张表格未在工作簿中找到对应记录，已保持原样")

    if not synced_codes and not created_count and not new_codes:
        flush_log_summary()
        log(f"没有任何表块匹配到 Word 表格，未做任何写入（共 {len(order)} 个表块，全部跳过）。")
        return 0

    meta_changed = 0
    for code in synced_codes:
        block = blocks[code]
        index = table_index.get(recorded[code])
        if index is None or index >= len(infos):
            continue
        # 元数据里的书签名必须沿用工作簿里记录的那个（配对依据），不能被占位名覆盖
        info = {**infos[index], "bookmark": recorded[code]}
        row = block["anchor_row"]
        before = [ws.cell(row, col).value for col in range(1, 6)]
        refresh_anchor_metadata(
            ws,
            block,
            info,
            docx_path.name,
            with_link=block.get("meta", {}).get("link_col") is not None or code in new_codes,
            ncols=block["ncols"],
        )
        after = [ws.cell(row, col).value for col in range(1, 6)]
        if before != after:
            meta_changed += 1

    name_changes = refresh_defined_names(wb, {code: blocks[code] for code in synced_codes})
    for item in name_changes:
        log(f"[定位] 定义名称已更新：{item}")

    flush_log_summary()
    reasons = []
    if structure_skips:
        reasons.append(f"结构不符 {structure_skips} 张")
    other_skips = len(skipped_codes) - structure_skips
    if other_skips:
        reasons.append(f"已删除/缺少书签 {other_skips} 张")
    skip_note = f"，跳过 {len(skipped_codes)} 张表（{'、'.join(reasons)}）" if skipped_codes else ""
    new_note = f"，新增 {created_count} 张表" if created_count else ""
    sort_note = f"，排序 {sort_changed} 张表" if sort_changed else ""
    cross_note = f"，删除✕行 {total_cross} 行" if total_cross else ""
    log(
        f"合计：同步 {len(synced_codes)} 张表（改动 {total_cells} 个单元格，"
        f"{total_rows} 行结构，{total_columns} 列结构）{new_note}{sort_note}{cross_note}{skip_note}。"
    )

    if args.dry_run:
        log("预演结束，未写入任何文件。")
        return 0

    docx_changed = bool(total_cells or total_rows or total_columns or marker_removed or created_count)
    xlsx_changed = bool(meta_changed or name_changes or sort_changed or mark_changed)
    if not docx_changed and not xlsx_changed:
        log("Word 与 Excel 数据一致，未做任何写入。")
        return 0

    backups = save_docx_xlsx_transaction(
        docx_path=docx_path if docx_changed else None,
        doc=doc if docx_changed else None,
        xlsx_path=xlsx_path if xlsx_changed else None,
        wb=wb if xlsx_changed else None,
        backup=getattr(args, "backup", True),
    )
    xlsx_backup = backups.get(xlsx_path) if xlsx_changed else None
    docx_backup = backups.get(docx_path) if docx_changed else None

    if xlsx_changed:
        log(
            f"[定位] 已保存 {xlsx_path.name}（元数据 {meta_changed} 处、定义名称 {len(name_changes)} 处"
            + (f"、✕行标记已删 {mark_changed} 行" if mark_changed else "")
            + f"）"
            + (f"（原文件已备份：{xlsx_backup.name}）" if xlsx_backup else "")
        )

    if docx_changed:
        log(f"[回写] 已保存 {docx_path.name}" + (f"（原文件已备份：{docx_backup.name}）" if docx_backup else ""))
    if getattr(args, "changelog", False):
        export_changelog(recorder, xlsx_path, docx_path.name)
    return 0
