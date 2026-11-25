"""bookmark 子命令：按内容把书签重新对齐到正确的表格，并把书签整表锚定。

用法要点：
* 默认**修复**：用「列数+行数+首行+首列」内容签名把工作簿表块与 Word 表格 1:1 配对，
  然后把书签重建到对应表格内部（覆盖整张表），Excel 里点“→定位Word表”会选中整张表。
* 只接受完全一致的签名，不做模糊猜测；对不上的表会列为“未对齐”。
* 只删除“本工具自己的书签”（名字记录在工作簿 C 列里的，以及早期的 TBL_*），
  审计软件的书签与 `_Toc…` 一律不动。
"""

import docx
from openpyxl import load_workbook

from ..align import apply_table_bookmarks, plan_table_bookmarks, split_bookmark_updates
from ..config import EXCEL_LINK_TEXT, SHEET_NAME
from ..console import log
from ..docx_bookmarks import cleanup_excel_markers
from ..docx_grid import body_items
from ..docx_meta import collect_table_info, dedupe_table_names
from ..errors import SyncError
from ..excel_blocks import parse_blocks, read_recorded_bookmarks, refresh_workbook_metadata
from ..fsops import backup_file, check_file_free, resolve_docx, resolve_xlsx


def _short_list(codes: list[str], limit: int = 12) -> str:
    shown = "、".join(codes[:limit])
    return shown + ("…" if len(codes) > limit else "")


def cmd_bookmark(args) -> int:
    docx_path = resolve_docx(args.docx)
    xlsx_path = resolve_xlsx(docx_path, args.excel)
    if not args.dry_run:
        check_file_free(docx_path, args.force)
        if xlsx_path.exists():
            check_file_free(xlsx_path, args.force)

    doc = docx.Document(str(docx_path))
    if not doc.tables:
        raise SyncError("Word 文档中没有任何表格，无需处理")
    items = body_items(doc)
    infos = [collect_table_info(doc, items, idx, tab) for idx, tab in enumerate(doc.tables)]
    dedupe_table_names(infos)
    recorded = read_recorded_bookmarks(xlsx_path)

    ws = None
    blocks: dict = {}
    if xlsx_path.exists():
        workbook = load_workbook(str(xlsx_path))
        if SHEET_NAME in workbook.sheetnames:
            ws = workbook[SHEET_NAME]
            _, blocks = parse_blocks(ws)

    plan = plan_table_bookmarks(doc, ws, blocks, recorded, name_unmatched=not blocks)
    code_by_index = {index: code for code, index in plan["mapping"].items()}

    log(f"读取 Word：{docx_path.name}（{len(doc.tables)} 张表）")
    if blocks:
        log(f"读取 Excel：{xlsx_path.name}（工作表：{SHEET_NAME}，{len(blocks)} 个表块）")
    else:
        log("未找到可用工作簿：将为每张表新建书签。")

    for index, name in sorted(plan["assignments"]):
        code = code_by_index.get(index, "（新表）")
        log(f"[对齐] 第 {index + 1} 张表 ↔ {code}（书签 {name}）")
    for index in plan["unmatched_tables"]:
        if code_by_index:  # 有工作簿时才提示“未对齐”
            log(f"[未对齐] 第 {index + 1} 张表：内容与任何表块都不匹配（建议重新导出工作簿）")
    if plan["unmatched_codes"]:
        log(
            f"[提示] {len(plan['unmatched_codes'])} 个表块在 Word 中已删除，不会被同步："
            f"{_short_list(plan['unmatched_codes'])}"
        )

    if args.dry_run:
        log("预演模式（--dry-run）：以上为将执行的书签对齐，未写入任何文件。")
        return 0

    # 已经正确锚定在目标表内的书签不必重建（保证重复运行幂等）
    assignments, remove_names = split_bookmark_updates(doc, plan["assignments"], recorded)
    result = apply_table_bookmarks(doc, assignments, remove_names)
    if result["removed"]:
        log(f"[书签] 已移除旧的/锚错的工具书签 {result['removed']} 个")
    if result["inserted"]:
        log(f"[书签] 已把 {result['inserted']} 个书签整表锚定（点击可选中整张表格）")
    if not result["removed"] and not result["inserted"]:
        log("[书签] 书签已全部正确锚定，无需重建")

    markers = cleanup_excel_markers(doc)
    if markers:
        log(f"[清理] 已删除 Word 正文中遗留的 {markers} 处 {EXCEL_LINK_TEXT} 文字/断链")

    if result["removed"] or result["inserted"] or markers:
        backup = backup_file(docx_path)
        doc.save(str(docx_path))
        log(f"[书签] 已保存 {docx_path.name}" + (f"（原文件已备份：{backup.name}）" if backup else ""))
    else:
        log("[书签] Word 已是目标状态，无需保存")

    if ws is not None and blocks:
        name_by_index = {index: name for index, name in plan["assignments"]}
        infos_by_code: dict[str, dict] = {}
        for code, index in plan["mapping"].items():
            info = dict(infos[index]) if index < len(infos) else None
            if info is None:
                continue
            info["bookmark"] = name_by_index.get(index, "")
            infos_by_code[code] = info
        refresh_workbook_metadata(
            xlsx_path,
            docx_path.name,
            infos_by_code,
            with_links=not getattr(args, "no_excel_links", False),
        )
    return 0
