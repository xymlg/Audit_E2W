"""export 子命令：按当前 Word 重新生成工作簿，并把书签整表锚定。

导出时会先按内容把现有工作簿的表块与当前 Word 表格对齐，沿用原有的书签名
（对不上的块/表会自动报告），然后把书签重建到表格内部（覆盖整张表）。
"""

import docx
from openpyxl import load_workbook

from ..align import apply_table_bookmarks, plan_table_bookmarks, split_bookmark_updates
from ..config import DELETED_BOOKMARK, EXCEL_LINK_TEXT, NEW_BOOKMARK, SHEET_NAME
from ..console import log
from ..docx_bookmarks import cleanup_excel_markers
from ..docx_grid import body_items
from ..docx_marks import ensure_mark_bookmarks, find_marks, processed_marks
from ..docx_meta import collect_table_info, dedupe_table_names
from ..docx_placeholders import VALUES_SHEET, init_placeholder_sheet, read_placeholder_values, scan_placeholders
from ..errors import SyncError
from ..excel_blocks import carry_over_deleted_blocks, parse_blocks, read_recorded_bookmarks, rejudge_marks
from ..excel_export import export_workbook
from ..excel_marks import SHEET_MARKS, build_marks_sheet, read_previous_marks
from ..fsops import backup_file, check_file_free, load_workbook_for_save, resolve_docx, resolve_xlsx


def _code_number(code: str) -> int:
    try:
        return int(code[1:])
    except ValueError:
        return 0


def _removed_block_hints(xlsx_path, recorded: dict[str, str], blocks_old: dict, plan: dict) -> list[str]:
    """导出前提醒：现有工作簿里有块在 Word 中已找不到对应表格（按实际对齐结果判断）。"""
    if not recorded:
        if xlsx_path.exists():
            return ["[提醒] 现有工作簿没有书签记录，本次导出将按当前 Word 重建并重新登记书签。"]
        return []
    missing = sorted(
        (
            code for code in plan["unmatched_codes"]
            if (blocks_old.get(code) or {}).get("bookmark") != DELETED_BOOKMARK
        ),
        key=_code_number,
    )
    unnamed = sorted((code for code, name in recorded.items() if not name), key=_code_number)
    hints: list[str] = []
    if missing:
        shown = "、".join(missing[:8]) + ("…" if len(missing) > 8 else "")
        hints.append(
            f"[提醒] 现有工作簿中有 {len(missing)} 个表块在 Word 中已找不到对应表格"
            f"（可能是已删除的表：{shown}）；导出将按当前 Word 重建，这些块会消失。"
        )
    if unnamed:
        hints.append(f"[提醒] 现有工作簿中有 {len(unnamed)} 个表块没有书签记录，导出后将重新登记书签。")
    return hints


def _load_previous_workbook(xlsx_path):
    """读取现有工作簿（表块 + 工作表），用于沿用书签名；不存在或读不了时返回 (None, {})。"""
    if not xlsx_path.exists():
        return None, {}
    try:
        workbook = load_workbook(str(xlsx_path))
    except Exception as exc:  # 文件被占用或损坏时不阻断导出
        log(f"[提醒] 读取现有工作簿失败（{exc}），本次不沿用旧书签名。")
        return None, {}
    if SHEET_NAME not in workbook.sheetnames:
        return None, {}
    ws = workbook[SHEET_NAME]
    _, blocks = parse_blocks(ws)
    return ws, blocks


def cmd_export(args) -> int:
    docx_path = resolve_docx(args.docx)
    xlsx_path = resolve_xlsx(docx_path, args.excel)
    if not args.dry_run:
        check_file_free(docx_path, args.force)
        check_file_free(xlsx_path, args.force)

    doc = docx.Document(str(docx_path))
    items = body_items(doc)
    infos = [collect_table_info(doc, items, idx, tab) for idx, tab in enumerate(doc.tables)]
    dedupe_table_names(infos)
    recorded = {
        code: name
        for code, name in read_recorded_bookmarks(xlsx_path).items()
        if name not in (NEW_BOOKMARK, DELETED_BOOKMARK)
    }
    ws_old, blocks_old = _load_previous_workbook(xlsx_path)
    previous_marks = read_previous_marks(xlsx_path)
    marks = find_marks(doc)
    processed = processed_marks(doc, {mark.bookmark for mark in marks if mark.bookmark})
    ph_fields = scan_placeholders(doc)
    ph_values_old = {}
    if ph_fields and xlsx_path.exists():
        try:
            ph_values_old = read_placeholder_values(xlsx_path, resolve_formulas=False)
        except Exception as exc:  # 旧工作簿读不了时不影响导出，只是带不过来已填值
            log(f"[提醒] 读取原「{VALUES_SHEET}」表失败（{exc}），本次无法保留已填值。")

    plan = plan_table_bookmarks(doc, ws_old, blocks_old, recorded, name_unmatched=True)
    name_by_index = {index: name for index, name in plan["assignments"]}
    for info in infos:
        info["bookmark"] = name_by_index.get(info["index"], "")

    log(f"读取 Word：{docx_path.name}")
    log(f"发现表格：{len(infos)} 张" + ("（无表格，工作表「" + SHEET_NAME + "」将为空）" if not infos else ""))
    with_links = not getattr(args, "no_excel_links", False)
    for info in infos:
        log(
            f"  {info['code']} {info['name']}（{info['nrows']} 行 × {info['ncols']} 列，"
            f"合并区 {sum(1 for m in info['merges'] if m[2] > 1 or m[3] > 1)} 处，书签 {info['bookmark']}）"
        )
    log(f"目标工作簿：{xlsx_path.name}（工作表：{SHEET_NAME}）")
    log(f"高亮标记：{len(marks)} 处（将写入工作表「{SHEET_MARKS}」）")
    if ph_fields:
        log(f"正文占位符：{len(ph_fields)} 个（将写入工作表「{VALUES_SHEET}」）")
    else:
        log(f"正文占位符：未发现（工作表「{VALUES_SHEET}」将为空）")
    for hint in _removed_block_hints(xlsx_path, recorded, blocks_old, plan):
        log(hint)

    if args.dry_run:
        log("预演模式（--dry-run）：以上为将导出的内容，未写入任何文件。")
        return 0

    assignments, remove_names = split_bookmark_updates(doc, plan["assignments"], recorded)
    result = apply_table_bookmarks(doc, assignments, remove_names)
    if result["inserted"]:
        log(f"[书签] 已把 {result['inserted']} 个书签整表锚定（点击可选中整张表格）")
    created_marks = ensure_mark_bookmarks(doc, marks)
    if created_marks:
        log(f"[标记] 已为 {created_marks} 处高亮标记插入定位书签（MRK_xxxx，正文外观不变）")
    removed_markers = cleanup_excel_markers(doc)
    if removed_markers:
        log(f"[清理] 已删除 Word 正文中遗留的 {removed_markers} 处 {EXCEL_LINK_TEXT} 文字/断链")

    workbook = export_workbook(
        infos, docx_path.name, with_links=with_links,
        totals_formula=not getattr(args, "no_total_formulas", False),
    )
    if ws_old is not None:
        carried = carry_over_deleted_blocks(workbook, ws_old, blocks_old)
        if carried:
            log(f"[保留] 已可逆删除的 {len(carried)} 个表块原样保留（{'、'.join(carried)}），可在「删除空白表格」页恢复。")
        from ..sort_rules import carry_over_sort_settings

        sort_carried = carry_over_sort_settings(workbook, ws_old, blocks_old)
        if sort_carried:
            log(f"[保留] {len(sort_carried)} 个表块的排序设置已按书签沿用（{'、'.join(sort_carried)}）。")
    mark_stats = build_marks_sheet(workbook, docx_path.name, marks, previous_marks, processed)
    log(
        f"[标记] 工作表「{SHEET_MARKS}」已写入 {mark_stats['marks']} 行"
        + (f"，{mark_stats['processed']} 行已处理" if mark_stats.get("processed") else "")
        + (f"，另有 {mark_stats['orphans']} 行已删除记录" if mark_stats["orphans"] else "")
    )
    if ph_fields:
        ws_ph = init_placeholder_sheet(workbook)
        for field in ph_fields:
            ws_ph.append([field, ph_values_old.get(field, "")])
        kept = sum(1 for f in ph_fields if ph_values_old.get(f))
        log(
            f"[占位符] 工作表「{VALUES_SHEET}」已写入 {len(ph_fields)} 个字段"
            + (f"，保留已填值 {kept} 个" if kept else "（待填写）")
        )
    else:
        init_placeholder_sheet(workbook)
        log(f"[占位符] 正文无 {{字段名}}，工作表「{VALUES_SHEET}」已创建为空表。")
    backup = getattr(args, "backup", True)
    if not backup:
        log("[导出] 按选项不备份原文件（--no-backup / 界面取消勾选），直接覆盖。")
    xlsx_backup = backup_file(xlsx_path) if backup else None
    workbook.save(str(xlsx_path))
    log(f"[导出] 已写入 {xlsx_path.name}" + (f"（原文件已备份：{xlsx_backup.name}）" if xlsx_backup else ""))

    if result["removed"] or result["inserted"] or created_marks or removed_markers:
        docx_backup = backup_file(docx_path) if backup else None
        doc.save(str(docx_path))
        log(f"[导出] 已更新 {docx_path.name}" + (f"（原文件已备份：{docx_backup.name}）" if docx_backup else ""))
    else:
        log("[导出] Word 中的书签已是目标状态，无需改动。")
    return 0


def cmd_rejudge(args) -> int:
    """按工作簿当前内容重新判定 B 列表头/✓/✕（只改标记，不改数据、不改 Word）。"""
    docx_path = resolve_docx(args.docx)
    xlsx_path = resolve_xlsx(docx_path, args.excel)
    if not xlsx_path.exists():
        raise SyncError(f"找不到工作簿：{xlsx_path}\n请先执行「导出表格到 Excel」。")
    if not args.dry_run:
        check_file_free(xlsx_path, args.force)
    workbook = load_workbook_for_save(xlsx_path)
    if SHEET_NAME not in workbook.sheetnames:
        raise SyncError(f"工作簿中找不到工作表“{SHEET_NAME}”，请确认该文件由本工具导出")
    ws = workbook[SHEET_NAME]
    _, blocks = parse_blocks(ws)
    if not blocks:
        raise SyncError("工作簿中没有表块，请先执行「导出表格到 Excel」")
    log(f"读取工作簿：{xlsx_path.name}（表块 {len(blocks)} 个）")
    stats = rejudge_marks(ws, blocks)
    if stats["old"]:
        log(
            f"[提醒] {'、'.join(stats['old'])} 为旧布局（没有标记列），无法重新判定；"
            f"请执行一次「导出表格到 Excel」升级为新布局。"
        )
    prefix = "[判定·预演]" if args.dry_run else "[判定]"
    log(f"{prefix} 重新判定 {stats['rows']} 行，标记变化 {stats['changed']} 行（表头/✓/✕；「已删」行保持原样）")
    if args.dry_run:
        log("预演模式（--dry-run）：未写入任何文件。")
        return 0
    if not stats["changed"]:
        log("标记与当前内容一致，工作簿无需写入。")
        return 0
    backup = getattr(args, "backup", True)
    xlsx_backup = backup_file(xlsx_path) if backup else None
    workbook.save(str(xlsx_path))
    log(f"[判定] 已保存 {xlsx_path.name}" + (f"（原文件已备份：{xlsx_backup.name}）" if xlsx_backup else ""))
    return 0
