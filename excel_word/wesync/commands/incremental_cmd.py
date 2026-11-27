"""增量导出：只登记 Word 中尚未锚定的表格、标注与占位符。"""

from collections import Counter

import docx

from ..config import BOOKMARK_PREFIX, SHEET_NAME, SHEET_NEW_TABLES
from ..console import log, warn
from ..docx_bookmarks import bookmark_names, insert_table_bookmark, new_cpah_name, next_bookmark_id, table_index_map_by_bookmark
from ..docx_grid import body_items
from ..docx_marks import _duplicate_mark_bookmarks, ensure_mark_bookmarks, find_marks
from ..docx_meta import collect_table_info, dedupe_table_names
from ..docx_placeholders import append_placeholder_sheet, scan_placeholders
from ..errors import SyncError
from ..excel_blocks import parse_blocks
from ..excel_export import export_workbook
from ..excel_marks import append_marks_sheet, read_previous_marks
from ..fsops import check_file_free, load_workbook_for_save, resolve_docx, resolve_xlsx, save_docx_xlsx_transaction


def _load_blocks(wb):
    order: list[str] = []
    blocks: dict[str, dict] = {}
    for sheet_name in (SHEET_NAME, SHEET_NEW_TABLES):
        if sheet_name not in wb.sheetnames:
            continue
        sheet_order, sheet_blocks = parse_blocks(wb[sheet_name])
        duplicate = set(sheet_order) & set(blocks)
        if duplicate:
            raise SyncError(f"工作簿中表编号重复：{'、'.join(sorted(duplicate))}。请修正后再增量导出。")
        for code, block in sheet_blocks.items():
            block["sheet_name"] = sheet_name
            block["_ws"] = wb[sheet_name]
        order.extend(sheet_order)
        blocks.update(sheet_blocks)
    return order, blocks


def _ignored_table_content(doc) -> tuple[int, int]:
    from ..docx_marks import _paragraph_mark_highlighted, run_highlighted
    from ..docx_placeholders import PLACEHOLDER_RE

    marks = placeholders = 0
    roots = [doc.element.body]
    for section in doc.sections:
        roots.extend(part._element for part in (
            section.header, section.footer, section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer,
        ) if part is not None)
    for root in roots:
        for tc in root.iter(docx.oxml.ns.qn("w:tc")):
            for p_el in tc.iter(docx.oxml.ns.qn("w:p")):
                if _paragraph_mark_highlighted(p_el) or any(run_highlighted(r, doc) for r in p_el.iter(docx.oxml.ns.qn("w:r"))):
                    marks += 1
                placeholders += len(PLACEHOLDER_RE.findall("".join(t.text or "" for t in p_el.iter(docx.oxml.ns.qn("w:t")))))
    return marks, placeholders


def cmd_incremental_export(args) -> int:
    docx_path = resolve_docx(args.docx)
    xlsx_path = resolve_xlsx(docx_path, args.excel)
    if not xlsx_path.exists():
        raise SyncError(f"找不到工作簿：{xlsx_path}\n请先执行一次全量「导出表格到 Excel」。")
    if not args.dry_run:
        check_file_free(docx_path, args.force)
        check_file_free(xlsx_path, args.force)

    doc = docx.Document(str(docx_path))
    wb = load_workbook_for_save(xlsx_path)
    if SHEET_NAME not in wb.sheetnames:
        raise SyncError(f"工作簿中找不到工作表“{SHEET_NAME}”，请确认该文件由本工具导出")
    order, blocks = _load_blocks(wb)
    previous_marks = read_previous_marks(xlsx_path)
    known_mark_bookmarks = {row.get("bookmark") for row in previous_marks if row.get("bookmark")}

    recorded = {code: (block.get("bookmark") or "").strip() for code, block in blocks.items()}
    candidates = {name for name in recorded.values() if name}
    table_index = table_index_map_by_bookmark(doc, candidates, include_anchors=True)
    counts = Counter(bookmark_names(doc.element.body))
    duplicates = [name for name in counts if name.startswith(BOOKMARK_PREFIX) and counts[name] > 1]
    if duplicates:
        raise SyncError(f"Word 中存在重复的表格书签：{'、'.join(sorted(duplicates))}。请先清理重复书签。")
    _duplicate_mark_bookmarks(doc)

    items = body_items(doc)
    new_indices = [index for index in range(len(doc.tables)) if index not in set(table_index.values())]
    all_infos = [collect_table_info(doc, items, index, tab) for index, tab in enumerate(doc.tables)]
    dedupe_table_names(all_infos)
    existing_names = bookmark_names(doc.element.body)
    bookmark_id = next_bookmark_id(doc)
    new_infos = []
    next_number = max([int(code[2:]) for code in order if code.startswith("XT") and code[2:].isdigit()] + [0]) + 1
    for offset, index in enumerate(new_indices):
        info = dict(all_infos[index])
        name = new_cpah_name(existing_names)
        info["bookmark"] = name
        info["code"] = f"XT{next_number + offset:02d}"
        new_infos.append(info)

    marks = [mark for mark in find_marks(doc) if not mark.bookmark or mark.bookmark not in known_mark_bookmarks]
    fields = scan_placeholders(doc)
    existing_fields = set()
    from ..docx_placeholders import VALUES_SHEET
    if VALUES_SHEET in wb.sheetnames:
        existing_fields = {str(wb[VALUES_SHEET].cell(r, 1).value or "").strip() for r in range(2, (wb[VALUES_SHEET].max_row or 1) + 1)}
    new_fields = [field for field in fields if field not in existing_fields]
    ignored_marks, ignored_placeholders = _ignored_table_content(doc)

    log(f"读取 Word：{docx_path.name}")
    log(f"增量预演：新增表格 {len(new_infos)} 张、新增标注 {len(marks)} 处、新增占位符 {len(new_fields)} 个。")
    for info in new_infos:
        log(f"  {info['code']} {info['name']}（{info['nrows']}行×{info['ncols']}列，书签 {info['bookmark']}）")
    for mark in marks:
        log(f"  {mark.location}：{mark.text[:40]}")
    if new_fields:
        log(f"  新字段：{'、'.join(new_fields)}")
    if ignored_marks or ignored_placeholders:
        log(f"[忽略] 表格内高亮 {ignored_marks} 处、占位符 {ignored_placeholders} 个不导出，避免与表格同步冲突。")
    if args.dry_run:
        log("预演模式：未写入任何文件。")
        wb.close()
        if getattr(wb, "vba_archive", None):
            wb.vba_archive.close()
        return 0

    if not new_infos and not marks and not new_fields:
        log("增量预演：没有新增表格、标注或占位符，未修改任何文件。")
        wb.close()
        if getattr(wb, "vba_archive", None):
            wb.vba_archive.close()
        return 0

    for info in new_infos:
        insert_table_bookmark(doc.tables[info["index"]]._tbl, info["bookmark"], bookmark_id)
        bookmark_id += 1
    created_marks = ensure_mark_bookmarks(doc, marks)
    if new_infos:
        export_workbook(
            new_infos, docx_path.name, with_links=not getattr(args, "no_excel_links", False),
            totals_formula=not getattr(args, "no_total_formulas", False), wb=wb, sheet_name=SHEET_NEW_TABLES,
        )
    mark_stats = append_marks_sheet(wb, docx_path.name, marks, previous_marks) if marks else {"marks": 0}
    added_fields = append_placeholder_sheet(wb, new_fields, source="增量") if new_fields else []
    backups = save_docx_xlsx_transaction(
        docx_path=docx_path if new_infos or created_marks else None,
        doc=doc if new_infos or created_marks else None,
        xlsx_path=xlsx_path, wb=wb, backup=getattr(args, "backup", True),
    )
    log(
        f"[增量导出] 已登记表格 {len(new_infos)} 张、标注 {mark_stats['marks']} 处、占位符 {len(added_fields)} 个；"
        "既有工作簿行与公式位置未改动。"
    )
    if backups.get(docx_path):
        log(f"[增量导出] Word 已备份：{backups[docx_path].name}")
    if backups.get(xlsx_path):
        log(f"[增量导出] Excel 已备份：{backups[xlsx_path].name}")
    return 0
