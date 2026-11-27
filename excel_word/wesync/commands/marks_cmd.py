"""marks / marks-sync 子命令：高亮标记导出到「链接标注」工作表，并把填写内容同步回 Word。"""

import docx
from openpyxl import load_workbook

from ..console import log, warn
from ..docx_marks import (
    apply_mark_change,
    ensure_mark_bookmarks,
    find_marks,
    mark_payload_by_bookmark,
    processed_marks,
)
from ..errors import SyncError
from ..excel_marks import SHEET_MARKS, build_marks_sheet, parse_marks_rows, read_previous_marks
from ..fsops import backup_file, check_file_free, load_workbook_for_save, resolve_docx, resolve_xlsx


def _mark_label(bookmark: str, fallback: str) -> str:
    return bookmark or fallback


def cmd_marks(args) -> int:
    """刷新工作簿里的「链接标注」工作表（不动表格数据，不改 Word 文字）。"""
    docx_path = resolve_docx(args.docx)
    xlsx_path = resolve_xlsx(docx_path, args.excel)
    if not xlsx_path.exists():
        raise SyncError(
            f"找不到工作簿：{xlsx_path}\n请先执行「导出表格到 Excel」（或界面上的“导出/刷新标注”会自动先导出）。"
        )
    if not args.dry_run:
        check_file_free(docx_path, args.force)
        check_file_free(xlsx_path, args.force)

    doc = docx.Document(str(docx_path))
    marks = find_marks(doc)
    previous = read_previous_marks(xlsx_path)
    processed = processed_marks(doc, {mark.bookmark for mark in marks if mark.bookmark})

    log(f"读取 Word：{docx_path.name}")
    log(f"发现高亮标记（待处理）：{len(marks)} 处" + (f"，已处理：{len(processed)} 处" if processed else ""))
    for mark in marks:
        filled = "已填" if any(row["bookmark"] == mark.bookmark and (row["new"] or row["link"]) for row in previous) else "未填"
        log(f"  M{mark.index:02d} {mark.location}（{filled}）：{mark.text[:40]}")
    known = {m.bookmark for m in marks if m.bookmark} | {item["bookmark"] for item in processed}
    stale = [row for row in previous if row["bookmark"] and row["bookmark"] not in known]
    if stale:
        log(f"  （另有 {len(stale)} 处工作簿里记录过的标记在 Word 中已删除，将保留为记录行）")

    if args.dry_run:
        log("预演模式（--dry-run）：以上为将写入「链接标注」工作表的标记，未写入任何文件。")
        return 0

    created = ensure_mark_bookmarks(doc, marks)
    if created:
        backup = backup_file(docx_path)
        doc.save(str(docx_path))
        log(f"[标记] 已为 {created} 处标记插入定位书签并保存 Word（备份：{backup.name if backup else '无'}）")
    else:
        log("[标记] 所有标记都已有定位书签，Word 无需改动")

    workbook = load_workbook_for_save(xlsx_path)
    stats = build_marks_sheet(workbook, docx_path.name, marks, previous, processed)
    xlsx_backup = backup_file(xlsx_path)
    workbook.save(str(xlsx_path))
    log(
        f"[标记] 已刷新工作表「{SHEET_MARKS}」：{stats['marks']} 行标记"
        + (f"，{stats['processed']} 行已处理" if stats.get("processed") else "")
        + (f"，{stats['orphans']} 行已删除记录" if stats["orphans"] else "")
        + (f"（原文件已备份：{xlsx_backup.name}）" if xlsx_backup else "")
    )
    return 0


def cmd_marks_sync(args) -> int:
    """把「链接标注」工作表里填写的内容同步回 Word。"""
    docx_path = resolve_docx(args.docx)
    xlsx_path = resolve_xlsx(docx_path, args.excel)
    if not xlsx_path.exists():
        raise SyncError(f"找不到工作簿：{xlsx_path}\n请先执行「导出表格到 Excel」或「导出/刷新标注」。")
    if not args.dry_run:
        check_file_free(docx_path, args.force)
        check_file_free(xlsx_path, args.force)

    workbook = load_workbook(str(xlsx_path))
    if SHEET_MARKS not in workbook.sheetnames:
        raise SyncError(f"工作簿里没有「{SHEET_MARKS}」工作表，请先在界面上执行“导出/刷新标注”。")
    rows = getattr(args, "rows", None)
    if rows is None:
        cached_workbook = load_workbook(str(xlsx_path), data_only=True)
        rows = parse_marks_rows(workbook[SHEET_MARKS], cached_workbook[SHEET_MARKS])
    if not rows:
        raise SyncError(f"「{SHEET_MARKS}」工作表里没有数据行，请先执行“导出/刷新标注”。")

    doc = docx.Document(str(docx_path))
    keep_highlight = bool(getattr(args, "keep_highlight", False))
    mark_parts = {mark.bookmark: mark.part for mark in find_marks(doc) if mark.bookmark}
    log(f"读取 Word：{docx_path.name}")
    log(f"读取 Excel：{xlsx_path.name}（工作表：{SHEET_MARKS}，{len(rows)} 行）")
    if keep_highlight:
        log("[标记] 本次同步保留高亮（不去色），便于人工复核")
    if args.dry_run:
        log("预演模式（--dry-run）：只显示将发生的改动，不写入文件。")

    changed = 0
    deleted = 0
    pending = 0
    for row in rows:
        label = row["label"]
        bookmark = row["bookmark"]
        if not bookmark:
            deleted += 1
            log(f"[跳过] {label}：该行没有书签，无法定位")
            continue
        payload = mark_payload_by_bookmark(doc, bookmark)
        if payload is None:
            deleted += 1
            log(f"[跳过] {label} {row['original'][:24]}：Word 中该标记已删除")
            continue
        if not row["new"] and not row["link"]:
            pending += 1
            warn(f"{label} {row['position']}：新内容与链接地址均未填，保留原样未同步。")
            continue
        p_el, children, runs = payload
        part = mark_parts.get(bookmark, doc.part)
        result = apply_mark_change(
            part, p_el, children, runs, row["new"], row["link"], keep_highlight=keep_highlight, doc=doc
        )
        if result == "deleted":
            deleted += 1
            log(f"[跳过] {label}：标记内容在 Word 中已为空")
        elif result == "changed":
            changed += 1
            detail = row["new"] if row["new"] else "（文字不变）"
            extra = f"，链接 → {row['link']}" if row["link"] else ""
            log(f"[标记] {label} {row['position']}：{row['original'][:20]} → {detail}{extra}")

    log(
        f"合计：同步 {changed} 处标记"
        + (f"，跳过 {deleted} 处（已删除/无书签）" if deleted else "")
        + (f"，待填写 {pending} 处（新内容与链接地址均为空）" if pending else "")
        + "。"
    )

    if args.dry_run:
        log("预演结束，未写入任何文件。")
        return 0
    if not changed:
        log("标注表与 Word 内容一致，未做任何写入。")
        return 0
    backup = backup_file(docx_path)
    doc.save(str(docx_path))
    log(f"[标记] 已保存 {docx_path.name}" + (f"（原文件已备份：{backup.name}）" if backup else ""))
    return 0
