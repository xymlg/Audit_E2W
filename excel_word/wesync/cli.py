"""命令行入口：定义 export / bookmark / writeback 三个子命令与 main()。

直接用包运行：``python -m wesync.cli <子命令> "<Word文件.docx>"``；
中文入口脚本（1-导出表格到Excel.py 等）通过 ``base_dir`` 启用“自动识别同目录唯一 docx”。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .commands.blank_cmd import cmd_blank_tables
from .commands.bookmark_cmd import cmd_bookmark
from .commands.clear_cmd import cmd_clear_bookmarks
from .commands.export_cmd import cmd_export, cmd_rejudge
from .commands.incremental_cmd import cmd_incremental_export
from .commands.marks_cmd import cmd_marks, cmd_marks_sync
from .commands.placeholders_cmd import cmd_placeholders, cmd_placeholders_sync
from .commands.writeback_cmd import cmd_writeback
from .console import USAGE_HINT
from .errors import SyncError
from .fsops import fill_docx_argument


def cmd_writeback_content(args) -> int:
    """writeback 子命令：勾选了标注/占位符时走内容编排，否则保持原回写命令。"""
    if not (args.with_marks or args.with_placeholders):
        return cmd_writeback(args)
    from . import api

    api.run_sync_content(
        args.docx,
        args.excel,
        tables=True,
        marks=args.with_marks,
        placeholders=args.with_placeholders,
        keep_highlight=args.keep_highlight,
        dry_run=args.dry_run,
        force=args.force,
        backup=True,
        delete_empty_rows=args.delete_empty_rows,
    )
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="word_excel_sync.py",
        description="Word 表格 ⇄ Excel 双向同步工具（导出、书签定位、回写）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    export_parser = sub.add_parser("export", help="导出 Word 全部表格到 Excel，并创建/归一化书签")
    export_parser.add_argument("docx", help="Word 文件路径（.docx）")
    export_parser.add_argument("--excel", help="输出工作簿路径，默认同目录 <主名>_表格数据.xlsx")
    export_parser.add_argument("--no-excel-links", action="store_true", help="不在 Excel 中写“→定位Word表”跳转链接")
    export_parser.add_argument("--no-total-formulas", action="store_true", help="合计行不写求和公式（默认写 =SUM，改明细后合计自动更新）")
    export_parser.add_argument("--no-backup", dest="backup", action="store_false", default=True, help="覆盖前不备份原工作簿/Word 到 _backup（默认备份）")
    export_parser.add_argument("--dry-run", action="store_true", help="只显示将导出的内容，不写文件")
    export_parser.add_argument("--force", action="store_true", help="跳过文件占用检查，强制尝试写入")
    export_parser.set_defaults(func=cmd_export)

    incremental_parser = sub.add_parser("incremental-export", help="只登记 Word 中新增的表格、标注与占位符，不改既有工作簿行")
    incremental_parser.add_argument("docx", help="Word 文件路径（.docx）")
    incremental_parser.add_argument("--excel", help="现有工作簿路径，默认同目录 <主名>_表格数据.xlsx")
    incremental_parser.add_argument("--no-excel-links", action="store_true", help="不在 Excel 中写 Word 定位链接")
    incremental_parser.add_argument("--no-total-formulas", action="store_true", help="新增表格的合计行不写求和公式")
    incremental_parser.add_argument("--no-backup", dest="backup", action="store_false", default=True, help="覆盖前不备份原文件到 _backup")
    incremental_parser.add_argument("--dry-run", action="store_true", help="只显示将登记的内容，不写文件")
    incremental_parser.add_argument("--force", action="store_true", help="跳过文件占用检查，强制尝试写入")
    incremental_parser.set_defaults(func=cmd_incremental_export)

    rejudge_parser = sub.add_parser("rejudge", help="按工作簿当前内容重新判定 B 列表头/✓/✕（只改标记，不改数据）")
    rejudge_parser.add_argument("docx", help="Word 文件路径（.docx，用于定位默认工作簿）")
    rejudge_parser.add_argument("--excel", help="工作簿路径，默认同目录 <主名>_表格数据.xlsx")
    rejudge_parser.add_argument("--no-backup", dest="backup", action="store_false", default=True, help="覆盖前不备份原工作簿到 _backup（默认备份）")
    rejudge_parser.add_argument("--dry-run", action="store_true", help="只显示将变化的标记行数，不写文件")
    rejudge_parser.add_argument("--force", action="store_true", help="跳过文件占用检查，强制尝试写入")
    rejudge_parser.set_defaults(func=cmd_rejudge)

    bookmark_parser = sub.add_parser("bookmark", help="只重建/迁移 Word 书签并清理遗留文字，不动数据")
    bookmark_parser.add_argument("docx", help="Word 文件路径（.docx）")
    bookmark_parser.add_argument("--excel", help="工作簿路径，默认同目录 <主名>_表格数据.xlsx")
    bookmark_parser.add_argument("--no-excel-links", action="store_true", help="不在 Excel 中写“→定位Word表”跳转链接")
    bookmark_parser.add_argument("--dry-run", action="store_true", help="只显示将处理的内容，不写文件")
    bookmark_parser.add_argument("--force", action="store_true", help="跳过文件占用检查，强制尝试写入")
    bookmark_parser.set_defaults(func=cmd_bookmark)

    writeback_parser = sub.add_parser("writeback", help="把 Excel 的改动（含增删行、增删列）回写到 Word")
    writeback_parser.add_argument("docx", help="Word 文件路径（.docx）")
    writeback_parser.add_argument("--excel", help="工作簿路径，默认同目录 <主名>_表格数据.xlsx")
    writeback_parser.add_argument("--dry-run", action="store_true", help="只显示将回写的改动，不写文件")
    writeback_parser.add_argument("--force", action="store_true", help="跳过文件占用检查，强制尝试写入")
    writeback_parser.add_argument("--delete-empty-rows", action="store_true", help="把 B 列标记为✕的行连 Word 行一起删除（Excel 保留为「已删」，改回✓可恢复）")
    writeback_parser.add_argument("--with-marks", action="store_true", help="同时同步「链接标注」内容（默认只同步表格）")
    writeback_parser.add_argument("--with-placeholders", action="store_true", help="同时按「正文信息」表替换 Word 的 {字段名} 占位符（默认只同步表格）")
    writeback_parser.add_argument("--keep-highlight", action="store_true", help="同步标注后保留高亮（默认去色；仅 --with-marks 时有意义）")
    writeback_parser.set_defaults(func=cmd_writeback_content)

    marks_parser = sub.add_parser("marks", help="把 Word 里的高亮标记导出/刷新到工作表的「链接标注」页")
    marks_parser.add_argument("docx", help="Word 文件路径（.docx）")
    marks_parser.add_argument("--excel", help="工作簿路径，默认同目录 <主名>_表格数据.xlsx")
    marks_parser.add_argument("--dry-run", action="store_true", help="只显示将导出的标记，不写文件")
    marks_parser.add_argument("--force", action="store_true", help="跳过文件占用检查，强制尝试写入")
    marks_parser.set_defaults(func=cmd_marks)

    marks_sync_parser = sub.add_parser("marks-sync", help="把「链接标注」页里填写的内容同步回 Word")
    marks_sync_parser.add_argument("docx", help="Word 文件路径（.docx）")
    marks_sync_parser.add_argument("--excel", help="工作簿路径，默认同目录 <主名>_表格数据.xlsx")
    marks_sync_parser.add_argument(
        "--keep-highlight", action="store_true", help="同步后保留该处高亮（默认同步后自动去色）"
    )
    marks_sync_parser.add_argument("--dry-run", action="store_true", help="只显示将同步的改动，不写文件")
    marks_sync_parser.add_argument("--force", action="store_true", help="跳过文件占用检查，强制尝试写入")
    marks_sync_parser.set_defaults(func=cmd_marks_sync)

    ph_parser = sub.add_parser("placeholders", help="按 Word 的 {字段名} 给工作簿「正文信息」表补齐填写骨架（保留已填值）")
    ph_parser.add_argument("docx", help="Word 文件路径（.docx）")
    ph_parser.add_argument("--excel", help="工作簿路径，默认同目录 <主名>_表格数据.xlsx")
    ph_parser.add_argument("--dry-run", action="store_true", help="只显示将补齐的字段，不写文件")
    ph_parser.add_argument("--force", action="store_true", help="跳过文件占用检查，强制尝试写入")
    ph_parser.set_defaults(func=cmd_placeholders)

    ph_sync_parser = sub.add_parser("placeholders-sync", help="把「正文信息」表填好的值写回 Word 占位符（未填保留原样并警告）")
    ph_sync_parser.add_argument("docx", help="Word 文件路径（.docx）")
    ph_sync_parser.add_argument("--excel", help="工作簿路径，默认同目录 <主名>_表格数据.xlsx")
    ph_sync_parser.add_argument("--dry-run", action="store_true", help="只显示将替换的占位符，不写文件")
    ph_sync_parser.add_argument("--force", action="store_true", help="跳过文件占用检查，强制尝试写入")
    ph_sync_parser.set_defaults(func=cmd_placeholders_sync)

    blank_parser = sub.add_parser("blank", help="删除 Word 里的空白表格（整表全空/仅表头有字），可连表标题一起删")
    blank_parser.add_argument("docx", help="Word 文件路径（.docx）")
    blank_parser.add_argument("--pick", help="要删除的候选编号，逗号分隔，如 1,3（编号见扫描清单）")
    blank_parser.add_argument("--all", action="store_true", help="删除全部候选空白表")
    blank_parser.add_argument("--keep-caption", action="store_true", help="只删表格，保留表标题段落")
    blank_parser.add_argument("--reversible", action="store_true", help="可逆删除：Excel 表块保留并标记，可随时恢复")
    blank_parser.add_argument("--restore", help="恢复已可逆删除的表：删除清单编号（逗号分隔）或 all")
    blank_parser.add_argument("--excel", help="工作簿路径（默认同目录 <主名>_表格数据.xlsx）")
    blank_parser.add_argument("--dry-run", action="store_true", help="只列出候选空白表，不写文件")
    blank_parser.add_argument("--force", action="store_true", help="跳过文件占用检查，强制尝试写入")
    blank_parser.set_defaults(func=cmd_blank_tables)

    clear_parser = sub.add_parser("clear-bookmarks", help="删除 Word 里所有本工具书签（CPAH_/TBL_/MRK_），保留自带书签")
    clear_parser.add_argument("docx", help="Word 文件路径（.docx）")
    clear_parser.add_argument("--dry-run", action="store_true", help="只显示将删除哪些书签，不写文件")
    clear_parser.add_argument("--force", action="store_true", help="跳过文件占用检查，强制尝试写入")
    clear_parser.set_defaults(func=cmd_clear_bookmarks)

    return parser


def main(argv: list[str] | None = None, base_dir: Path | None = None) -> int:
    """运行入口。

    base_dir 不为 None 时（中文入口脚本使用），若没有显式给出 Word 文件，
    会自动使用该目录下唯一的 .docx；给出多个或没有时打印中文提示并退出。
    """
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    arguments = list(sys.argv[1:] if argv is None else argv)
    if not arguments:
        print(USAGE_HINT, flush=True)
        return 2
    subcommand, rest = arguments[0], arguments[1:]
    try:
        if base_dir is not None and not any(item in ("-h", "--help") for item in rest):
            rest = fill_docx_argument(rest, base_dir)
        parser = build_parser()
        args = parser.parse_args([subcommand, *rest])
        return args.func(args)
    except SyncError as exc:
        print(f"错误：{exc}", file=sys.stderr, flush=True)
        return 2
    except FileNotFoundError as exc:
        print(f"错误：找不到文件或路径不可用（{exc}）", file=sys.stderr, flush=True)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
