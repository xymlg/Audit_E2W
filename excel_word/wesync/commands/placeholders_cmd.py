"""placeholders / placeholders-sync 子命令：{字段名} 骨架导出到「正文信息」表、填值写回 Word。"""

import docx

from ..console import log, warn
from ..docx_placeholders import (
    VALUES_SHEET,
    apply_placeholders,
    read_placeholder_values,
    scan_placeholders,
    write_placeholder_sheet,
)
from ..errors import SyncError
from ..fsops import backup_file, check_file_free, resolve_docx, resolve_xlsx


def cmd_placeholders(args) -> int:
    """按当前 Word 的 {字段名} 给工作簿「正文信息」表补齐骨架（保留已填值）。"""
    docx_path = resolve_docx(args.docx)
    xlsx_path = resolve_xlsx(docx_path, args.excel)
    if not xlsx_path.exists():
        raise SyncError(
            f"找不到工作簿：{xlsx_path}\n请先执行「导出表格到 Excel」（界面上“导出/刷新占位符”会确认后自动先导出）。"
        )
    if not args.dry_run:
        check_file_free(xlsx_path, args.force)
    doc = docx.Document(str(docx_path))
    fields = scan_placeholders(doc)
    log(f"读取 Word：{docx_path.name}")
    if not fields:
        log("Word 中没有 {字段名} 占位符，无需生成「正文信息」表。")
        return 0
    log(f"发现占位符 {len(fields)} 个字段：{'、'.join(fields)}")
    if args.dry_run:
        log("预演模式（--dry-run）：以上字段将加入/补齐「正文信息」表，未写入任何文件。")
        return 0
    added = write_placeholder_sheet(xlsx_path, fields)
    log(
        f"[占位符] 「{VALUES_SHEET}」表已就绪（{len(fields)} 个字段"
        + (f"，本次新增 {len(added)} 个" if added else "，无新增字段")
        + "），已填值保留；请在 B 列填值后执行「同步占位符」。"
    )
    return 0


def cmd_placeholders_sync(args) -> int:
    """把「正文信息」表填好的值写回当前 Word 的占位符；未填字段保留原样并逐条警告。"""
    docx_path = resolve_docx(args.docx)
    xlsx_path = resolve_xlsx(docx_path, args.excel)
    if not xlsx_path.exists():
        raise SyncError(f"找不到工作簿：{xlsx_path}\n请先执行「导出表格到 Excel」或「导出/刷新占位符」。")
    if not args.dry_run:
        check_file_free(docx_path, args.force)
        check_file_free(xlsx_path, args.force)
    values = getattr(args, "values", None)
    if values is None:
        values = read_placeholder_values(xlsx_path)
    doc = docx.Document(str(docx_path))
    fields = scan_placeholders(doc)
    log(f"读取 Word：{docx_path.name}")
    if not fields:
        log("Word 中没有 {字段名} 占位符，无需同步。")
        return 0
    if not values:
        raise SyncError(
            f"工作簿没有「{VALUES_SHEET}」表或表里没填任何值，请先执行「导出/刷新占位符」并在 B 列填值。"
        )
    result = apply_placeholders(doc, values, dry_run=args.dry_run)
    for field, new in result["replaced"]:
        log(f"[占位符] {{{field}}} → {new}")
    for field in result["missing"]:
        warn(f"占位符 {{{field}}} 在「{VALUES_SHEET}」表未填值，保留原样未同步。")
    log(
        f"合计：替换 {len(result['replaced'])} 处占位符"
        + (f"，未填保留原样 {len(result['missing'])} 处" if result["missing"] else "")
        + "。"
    )
    if args.dry_run:
        log("预演结束，未写入任何文件。")
        return 0
    if not result["replaced"]:
        log("没有可替换的占位符，未做任何写入。")
        return 0
    backup = backup_file(docx_path)
    doc.save(str(docx_path))
    log(f"[占位符] 已保存 {docx_path.name}" + (f"（原文件已备份：{backup.name}）" if backup else ""))
    return 0