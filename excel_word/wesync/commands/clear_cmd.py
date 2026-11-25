"""clear-bookmarks 子命令：删除 Word 里所有本工具写入的书签（CPAH_/TBL_/MRK_）。

只清本工具自己创建的定位/标注书签；Word 自带书签（如 _Toc 目录）一律保留。
正式执行前自动备份原文件到 _backup。
"""

from __future__ import annotations

import docx

from ..console import log
from ..docx_bookmarks import clear_tool_bookmarks
from ..fsops import backup_file, check_file_free, resolve_docx


def cmd_clear_bookmarks(args) -> int:
    docx_path = resolve_docx(args.docx)
    if not args.dry_run:
        check_file_free(docx_path, args.force)

    doc = docx.Document(str(docx_path))
    result = clear_tool_bookmarks(doc)

    log(f"读取 Word：{docx_path.name}")
    log(f"[清理] 发现本工具书签（CPAH_/TBL_/MRK_）：{result['removed']} 个")
    if result["names"]:
        shown = "、".join(result["names"][:12])
        if len(result["names"]) > 12:
            shown += f" …（共 {len(result['names'])} 个）"
        log(f"[清理] 将删除：{shown}")
    if result["kept"]:
        log(f"[清理] 保留 Word 自带书签 {result['kept']} 个（如 _Toc 目录书签），未改动")

    if not result["removed"]:
        log("[清理] 没有发现本工具书签，无需处理")
        return 0

    if args.dry_run:
        log("预演模式（--dry-run）：未写入任何文件。")
        return 0

    backup = backup_file(docx_path)
    doc.save(str(docx_path))
    log(
        f"[清理] 已从 {docx_path.name} 删除 {result['removed']} 个书签"
        + (f"（原文件已备份：{backup.name}）" if backup else "")
    )
    return 0
