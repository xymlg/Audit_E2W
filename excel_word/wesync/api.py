"""对外稳定接口：把参数打包后调用现有命令实现（供界面等程序调用）。

命令行入口仍走 ``wesync.cli``，行为不变；界面只依赖本模块，
这样以后调整命令行参数不会影响界面代码。
"""

from __future__ import annotations

import argparse
from pathlib import Path

from .commands.blank_cmd import cmd_blank_tables
from .commands.bookmark_cmd import cmd_bookmark
from .commands.clear_cmd import cmd_clear_bookmarks
from .commands.export_cmd import cmd_export
from .commands.marks_cmd import cmd_marks, cmd_marks_sync
from .commands.writeback_cmd import cmd_writeback
from .console import log
from .docx_blank import find_blank_tables
from .errors import SyncError
from .fsops import backup_file, check_file_free, resolve_docx, resolve_xlsx, tmp_root


def _options(
    docx: str | Path,
    excel: str | Path | None,
    dry_run: bool,
    force: bool,
    with_excel_links: bool,
    backup: bool = True,
) -> argparse.Namespace:
    return argparse.Namespace(
        docx=str(docx),
        excel=None if excel is None else str(excel),
        dry_run=bool(dry_run),
        force=bool(force),
        no_excel_links=not with_excel_links,
        backup=bool(backup),
    )


def workbook_path(docx: str | Path, excel: str | Path | None = None) -> Path:
    """按工具规则计算工作簿路径（界面上显示用）。"""
    return resolve_xlsx(Path(docx), None if excel is None else str(excel))


def run_export(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    dry_run: bool = False,
    force: bool = False,
    with_excel_links: bool = True,
    totals_formula: bool = True,
    backup: bool = True,
) -> int:
    """导出：Word → Excel（按当前 Word 重建工作簿）。

    totals_formula=True 时，合计行（首列以「合计」开头）的数值列写 =SUM 求和公式，
    在 Excel 改明细后合计自动更新；置 False 则按原值导出。
    backup=False 时覆盖前不备份原工作簿/Word 到 _backup（默认备份）。
    """
    options = _options(docx, excel, dry_run, force, with_excel_links, backup)
    options.no_total_formulas = not totals_formula
    return cmd_export(options)


def run_rejudge(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    dry_run: bool = False,
    force: bool = False,
    backup: bool = True,
) -> int:
    """按工作簿当前内容重新判定 B 列表头/✓/✕（只改标记，不改数据、不改 Word）。

    在 Excel 里增删改内容后点一下，标记自动跟上；「已删」行保持原样；
    旧布局工作簿没有标记列，会提示先重新导出。
    """
    from .commands.export_cmd import cmd_rejudge

    return cmd_rejudge(_options(docx, excel, dry_run, force, True, backup))


def run_bookmark(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    dry_run: bool = False,
    force: bool = False,
    with_excel_links: bool = True,
) -> int:
    """书签工具：迁移/补建书签并清理正文遗留文字，不改数据。"""
    return cmd_bookmark(_options(docx, excel, dry_run, force, with_excel_links))


def run_writeback(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    dry_run: bool = False,
    force: bool = False,
    backup: bool = True,
    changelog: bool = False,
    delete_empty_rows: bool = False,
) -> int:
    """回写：Excel → Word（含增删行、增删列）。

    backup=False 时，写盘前不创建 _backup 备份（默认 True，保持原安全行为）。
    changelog=True 时，回写后把逐格改动导出成 Excel 复核清单（保存在工作簿旁）。
    delete_empty_rows=True 时，B 列标记为✕的行（导出时判定为空）连 Word 行一起删除，
    Excel 该行保留为「已删」（改回✓可恢复）；旧布局工作簿无标记列时不生效。
    """
    options = _options(docx, excel, dry_run, force, True, backup)
    options.changelog = bool(changelog)
    options.delete_empty_rows = bool(delete_empty_rows)
    return cmd_writeback(options)


def scan_writeback_diff(docx: str | Path, excel: str | Path | None = None) -> list[dict]:
    """只读差异比对：以预演方式跑回写逻辑，返回逐格改动记录（旧值→新值），不写任何文件。"""
    options = _options(docx, excel, True, False, True)
    options.recorder = []
    cmd_writeback(options)
    return options.recorder


def export_diff_report(docx: str | Path, excel: str | Path | None, records: list[dict]):
    """把差异比对记录导出成 Excel 清单（保存在工作簿旁），返回路径；无记录返回 None。"""
    from .changelog import export_changelog

    docx_path = resolve_docx(str(docx))
    xlsx_path = resolve_xlsx(docx_path, None if excel is None else str(excel))
    return export_changelog(records, xlsx_path, docx_path.name)


def scan_backups(docx: str | Path, excel: str | Path | None = None) -> dict:
    """只读列出报告 Word 与工作簿 Excel 各自的历史备份（按时间倒序），供回滚页分列展示。"""
    from .backup_history import list_backups

    docx_path = resolve_docx(str(docx))
    xlsx_path = resolve_xlsx(docx_path, None if excel is None else str(excel))
    return {
        "word": list_backups(docx_path),
        "excel": list_backups(xlsx_path),
        "word_path": str(docx_path),
        "excel_path": str(xlsx_path),
    }


def run_restore_backup(backup: str | Path, target: str | Path) -> None:
    """用选中的历史备份恢复目标文件（恢复前自动备份现状，可再次回滚）。"""
    from .backup_history import restore_backup

    restore_backup(backup, target)


def run_marks(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    dry_run: bool = False,
    force: bool = False,
) -> int:
    """刷新工作簿里的「链接标注」工作表（高亮标记导出，不改 Word 文字）。"""
    return cmd_marks(_options(docx, excel, dry_run, force, True))


def run_marks_sync(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    dry_run: bool = False,
    force: bool = False,
    keep_highlight: bool = False,
    _rows: list[dict] | None = None,
) -> int:
    """把「链接标注」工作表里填写的新内容/链接同步回 Word。"""
    options = _options(docx, excel, dry_run, force, True)
    options.keep_highlight = bool(keep_highlight)
    options.rows = _rows
    return cmd_marks_sync(options)


def run_placeholder_export(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    dry_run: bool = False,
    force: bool = False,
) -> int:
    """按当前 Word 的 {字段名} 给工作簿「正文信息」表补齐骨架（保留已填值）。"""
    from .commands.placeholders_cmd import cmd_placeholders

    return cmd_placeholders(_options(docx, excel, dry_run, force, True))


def run_placeholder_sync(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    dry_run: bool = False,
    force: bool = False,
    _values: dict[str, str] | None = None,
) -> int:
    """把「正文信息」表填好的值写回当前 Word 占位符；未填字段保留原样并警告。"""
    from .commands.placeholders_cmd import cmd_placeholders_sync

    options = _options(docx, excel, dry_run, force, True)
    options.values = _values
    return cmd_placeholders_sync(options)


def run_sync_content(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    tables: bool = True,
    marks: bool = False,
    placeholders: bool = False,
    keep_highlight: bool = True,
    dry_run: bool = False,
    force: bool = False,
    backup: bool = False,
    changelog: bool = False,
    delete_empty_rows: bool = False,
) -> dict:
    """按勾选内容编排同步：表格 → 标注 → 占位符，返回各部分执行摘要。

    勾选了标注/占位符但工作簿缺对应工作表（或标注无数据行、占位符没填任何值）时
    警告并跳过该部分内容（不静默）；表格部分的备份/改动清单/删除✕行选项只作用于表格。
    标注/占位符在表格保存前读取，避免公式缓存值因工作簿重存而丢失。
    """
    from .fsops import load_workbook_for_save as load_workbook

    from .console import warn
    from .docx_placeholders import VALUES_SHEET, read_placeholder_values
    from .excel_marks import SHEET_MARKS, parse_marks_rows

    if not (tables or marks or placeholders):
        raise SyncError("请至少勾选一项同步内容（表格/标注/占位符）。")
    xlsx = workbook_path(docx, excel)
    marks_rows = None
    placeholder_values = None
    if (marks or placeholders) and xlsx.exists():
        wb = load_workbook(str(xlsx))
        if marks:
            if SHEET_MARKS not in wb.sheetnames:
                marks_rows = []
            else:
                cached_wb = load_workbook(str(xlsx), data_only=True)
                marks_rows = parse_marks_rows(wb[SHEET_MARKS], cached_wb[SHEET_MARKS])
            if not marks_rows:
                warn(f"工作簿没有「{SHEET_MARKS}」数据行，跳过标注内容；请先在「标注与占位」页导出/刷新标注。")
                marks = False
        if placeholders:
            placeholder_values = read_placeholder_values(xlsx)
            if not placeholder_values:
                warn(f"工作簿没有「{VALUES_SHEET}」表或表里没填任何值，跳过占位符内容；请先在「标注与占位」页导出/刷新占位符并填值。")
                placeholders = False
    summary: dict = {}
    if tables:
        summary["tables"] = run_writeback(
            docx, excel, dry_run=dry_run, force=force, backup=backup,
            changelog=changelog, delete_empty_rows=delete_empty_rows,
        )
    if marks:
        summary["marks"] = run_marks_sync(
            docx, excel, dry_run=dry_run, force=force, keep_highlight=keep_highlight, _rows=marks_rows
        )
    if placeholders:
        summary["placeholders"] = run_placeholder_sync(
            docx, excel, dry_run=dry_run, force=force, _values=placeholder_values
        )
    return summary


def run_clear_bookmarks(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    dry_run: bool = False,
    force: bool = False,
) -> int:
    """删除 Word 里所有本工具书签（CPAH_/TBL_/MRK_），保留 Word 自带书签。"""
    return cmd_clear_bookmarks(_options(docx, None, dry_run, force, True))


def scan_blank_tables(docx_path: str | Path) -> list[dict]:
    """只读扫描：返回空白表候选清单（整表全空 / 仅表头有字），供界面展示勾选。"""
    import docx as _docx

    path = resolve_docx(str(docx_path))
    return find_blank_tables(_docx.Document(str(path)))


def run_blank_tables(
    docx: str | Path,
    *,
    dry_run: bool = False,
    force: bool = False,
    pick: str | None = None,
    keep_caption: bool = False,
    keep_caption_for: str | None = None,
    reversible: bool = False,
    excel: str | Path | None = None,
) -> int:
    """删除空白表。pick 为候选编号（如 "1,3"）或 "all"；不给 pick 时只列候选不删除。

    keep_caption=True 全部只删表格；keep_caption_for 指定保留标题的候选编号（如 "2"）。
    reversible=True 时可逆删除：Excel 表块保留并标记「已删除」，之后可恢复。
    """
    options = _options(docx, excel, dry_run, force, True)
    options.pick = None if pick in (None, "all") else str(pick)
    options.all = pick == "all"
    options.keep_caption = bool(keep_caption)
    options.keep_caption_for = keep_caption_for
    options.reversible = bool(reversible)
    return cmd_blank_tables(options)


def list_deleted_blocks(docx: str | Path, excel: str | Path | None = None) -> list[dict]:
    """只读列出工作簿里「已可逆删除」的表块，供界面展示恢复清单。"""
    from .config import DELETED_BOOKMARK

    return [item for item in list_table_blocks(docx, excel) if item["bookmark"] == DELETED_BOOKMARK]


def run_restore_tables(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    pick: str = "all",
    dry_run: bool = False,
    force: bool = False,
) -> int:
    """恢复已可逆删除的表。pick 为恢复清单编号（如 "1,3"）或 "all"。"""
    options = _options(docx, excel, dry_run, force, True)
    options.restore = str(pick)
    return cmd_blank_tables(options)


def list_table_blocks(docx: str | Path, excel: str | Path | None = None) -> list[dict]:
    """只读列出工作簿里的表块（编号、表名、书签标记），供界面选择插入位置/恢复清单。"""
    from .fsops import load_workbook_for_save as load_workbook

    from .config import SHEET_NAME
    from .excel_blocks import parse_blocks

    docx_path = resolve_docx(str(docx))
    xlsx_path = resolve_xlsx(docx_path, None if excel is None else str(excel))
    if not xlsx_path.exists():
        return []
    wb = load_workbook(str(xlsx_path))
    if SHEET_NAME not in wb.sheetnames:
        return []
    order, blocks = parse_blocks(wb[SHEET_NAME])
    result = []
    for code in order:
        block = blocks[code]
        name = (block.get("meta", {}).get("name") or "").strip()
        result.append(
            {
                "code": code,
                "name": name,
                "bookmark": block.get("bookmark", ""),
                "nrows": len(block["grid_rows"]),
                "ncols": block["ncols"],
                "restore_after": block.get("meta", {}).get("restore_after", ""),
            }
        )
    return result


def run_append_table_block(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    name: str,
    nrows: int,
    ncols: int,
    after: str | None = None,
) -> str:
    """在指定表块下面插入带「书签：新增」标记的新表块，返回表编号。

    after 指定插到哪张表下面（如 "T02"），不指定时插到最后一张表下面；
    后续表块整体下移，全部表块按上下顺序重新编号。Word 建表位置与 Excel 一致。
    之后在 Excel 里填好内容，执行「同步到 Word」即会在 Word 里创建这张表。
    """
    from .fsops import load_workbook_for_save as load_workbook

    from .excel_blocks import append_new_block

    docx_path = resolve_docx(str(docx))
    xlsx_path = resolve_xlsx(docx_path, None if excel is None else str(excel))
    if not xlsx_path.exists():
        raise SyncError(f"找不到工作簿：{xlsx_path}\n请先执行「导出表格到 Excel」生成工作簿。")
    check_file_free(xlsx_path, False)
    wb = load_workbook(str(xlsx_path))
    code = append_new_block(wb, name, nrows, ncols, after=after)
    backup = backup_file(xlsx_path)
    wb.save(str(xlsx_path))
    position = f"，插在 {after} 下面" if after else "，插在最后一张表下面"
    log(f"[新表] 已插入 {code}「{name.strip()}」（{nrows}行×{ncols}列，书签：新增{position}；全部表块已按顺序重新编号）"
        + (f"（原工作簿已备份：{backup.name}）" if backup else ""))
    log("[新表] 请在 Excel 里填写该表块内容（首行为表头），再执行「同步到 Word」，新表会建在 Excel 里同一位置对应的 Word 表格之后。")
    return code


def _load_blocks(docx, excel):
    from .fsops import load_workbook_for_save as load_workbook

    from .config import SHEET_NAME
    from .excel_blocks import parse_blocks

    docx_path = resolve_docx(str(docx))
    xlsx_path = resolve_xlsx(docx_path, None if excel is None else str(excel))
    if not xlsx_path.exists():
        raise SyncError(f"找不到工作簿：{xlsx_path}\n请先执行「导出表格到 Excel」生成工作簿。")
    wb = load_workbook(str(xlsx_path))
    if SHEET_NAME not in wb.sheetnames:
        raise SyncError(f"工作簿中找不到工作表“{SHEET_NAME}”，请确认该文件由本工具导出")
    order, blocks = parse_blocks(wb[SHEET_NAME])
    return xlsx_path, wb, wb[SHEET_NAME], order, blocks


def list_table_sorts(docx: str | Path, excel: str | Path | None = None) -> list[dict]:
    """只读列出各表块的排序设置与可选金额列（表头文字），供界面展示。"""
    from .config import DELETED_BOOKMARK, NEW_BOOKMARK
    from .sort_rules import block_headers, parse_sort_setting

    _, _, ws, order, blocks = _load_blocks(docx, excel)
    result = []
    for code in order:
        block = blocks[code]
        bookmark = block.get("bookmark", "")
        if bookmark in (NEW_BOOKMARK, DELETED_BOOKMARK):
            continue
        setting = parse_sort_setting(block.get("meta", {}).get("sort", ""))
        result.append(
            {
                "code": code,
                "name": (block.get("meta", {}).get("name") or "").strip(),
                "headers": block_headers(ws, block),
                "sort_column": setting[0] if setting else "",
                "descending": setting[1] if setting else True,
                "has_sort": setting is not None,
            }
        )
    return result


def set_table_sort(
    docx: str | Path,
    excel: str | Path | None = None,
    *,
    code: str,
    column: str,
    descending: bool = True,
) -> None:
    """保存某表的排序设置（写进工作簿锚点行，随文件保存，同步时自动应用）。"""
    from .sort_rules import write_block_sort

    xlsx_path, wb, ws, order, blocks = _load_blocks(docx, excel)
    if code not in blocks:
        raise SyncError(f"工作簿中找不到表块 {code}")
    check_file_free(xlsx_path, False)
    write_block_sort(ws, blocks[code], column, descending)
    backup = backup_file(xlsx_path)
    wb.save(str(xlsx_path))
    log(f"[排序] {code}：已保存排序设置「{column} {'降序' if descending else '升序'}」"
        + (f"（原工作簿已备份：{backup.name}）" if backup else ""))


def clear_table_sort(docx: str | Path, excel: str | Path | None = None, *, code: str) -> None:
    """清除某表的排序设置。"""
    from .sort_rules import clear_block_sort

    xlsx_path, wb, ws, order, blocks = _load_blocks(docx, excel)
    if code not in blocks:
        raise SyncError(f"工作簿中找不到表块 {code}")
    check_file_free(xlsx_path, False)
    if clear_block_sort(ws, blocks[code]):
        backup = backup_file(xlsx_path)
        wb.save(str(xlsx_path))
        log(f"[排序] {code}：已清除排序设置" + (f"（原工作簿已备份：{backup.name}）" if backup else ""))
    else:
        log(f"[排序] {code}：本来就没有排序设置，无需清除")


def run_batch_writeback(
    docx_list: list,
    *,
    dry_run: bool = False,
    backup: bool = True,
    force: bool = False,
    contents: tuple = ("tables",),
    keep_highlight: bool = True,
) -> list[dict]:
    """多报告批量同步：每份 Word 配同名工作簿逐个回写，单份失败不影响其他。

    contents 为同步内容组合（"tables"/"marks"/"placeholders" 任意子集）。
    """
    results: list[dict] = []
    total = len(docx_list)
    for i, docx in enumerate(docx_list, 1):
        path = Path(docx)
        log(f"===== [{i}/{total}] {path.name} =====")
        item = {"docx": str(path), "excel": "", "status": "failed", "reason": ""}
        results.append(item)
        try:
            xlsx = workbook_path(path)
            item["excel"] = str(xlsx)
            if not path.exists():
                item.update(status="skipped", reason="Word 文件不存在")
                log(f"[跳过] {path.name}：Word 文件不存在")
                continue
            if not xlsx.exists():
                item.update(status="skipped", reason="找不到工作簿，请先导出")
                log(f"[跳过] {path.name}：找不到工作簿 {xlsx.name}，请先执行「导出表格到 Excel」")
                continue
            run_sync_content(
                path,
                tables="tables" in contents,
                marks="marks" in contents,
                placeholders="placeholders" in contents,
                keep_highlight=keep_highlight,
                dry_run=dry_run,
                backup=backup,
                force=force,
            )
            item.update(status="ok", reason="预演完成" if dry_run else "同步完成")
        except SyncError as exc:
            item["reason"] = str(exc)
            log(f"[失败] {path.name}：{exc}")
        except Exception as exc:
            item["reason"] = f"{type(exc).__name__}: {exc}"
            log(f"[失败] {path.name}：{item['reason']}")
    ok = sum(1 for r in results if r["status"] == "ok")
    skipped = sum(1 for r in results if r["status"] == "skipped")
    failed = sum(1 for r in results if r["status"] == "failed")
    log(f"批量{'预演' if dry_run else '同步'}结束：成功 {ok} 份，跳过 {skipped} 份，失败 {failed} 份（共 {total} 份）。")
    return results


def run_template_batch(
    templates: list,
    items: list[dict],
    out_dir: str | Path,
    *,
    dry_run: bool = False,
    contents: tuple = ("tables", "placeholders"),
    keep_highlight: bool = True,
) -> list[dict]:
    """模板生成：每个 Excel 同步到指定模板的一份副本，产出若干新 Word。

    templates 为模板路径清单（编号从 1 起）；items 每项 {"excel": 路径, "template": 模板编号}。
    模板原件与源工作簿都不改动：同步在模板副本与工作簿临时副本上进行。
    contents 为同步内容组合（"tables"/"marks"/"placeholders" 任意子集）。
    """
    import shutil
    import tempfile

    import docx as _docx

    from .docx_bookmarks import bookmark_names
    from .docx_placeholders import apply_placeholders, read_placeholder_values, scan_placeholders
    from .excel_blocks import read_recorded_bookmarks

    if not contents:
        raise SyncError("请至少勾选一项同步内容（表格/标注/占位符）。")
    tpl_paths = [Path(t) for t in templates]
    out_root = Path(out_dir)
    if not dry_run:
        out_root.mkdir(parents=True, exist_ok=True)
    tpl_marks: dict[int, set] = {}
    tpl_fields: dict[int, list[str]] = {}
    results: list[dict] = []
    total = len(items)
    for i, item in enumerate(items, 1):
        xlsx = Path(item["excel"])
        tpl_no = int(item.get("template") or 0)
        log(f"===== [{i}/{total}] {xlsx.name} =====")
        record = {"excel": str(xlsx), "template": tpl_no, "output": "", "status": "failed", "reason": ""}
        results.append(record)
        try:
            if not 1 <= tpl_no <= len(tpl_paths):
                record.update(status="skipped", reason="未指定模板")
                log(f"[跳过] {xlsx.name}：未指定所属模板，请在清单第三列选择模板")
                continue
            tpl = tpl_paths[tpl_no - 1]
            if not tpl.exists():
                record.update(status="skipped", reason=f"模板{tpl_no}不存在")
                log(f"[跳过] {xlsx.name}：模板{tpl_no}（{tpl.name}）不存在")
                continue
            if not xlsx.exists():
                record.update(status="skipped", reason="工作簿不存在")
                log(f"[跳过] 工作簿不存在：{xlsx}")
                continue
            if tpl_no not in tpl_marks:
                tpl_doc = _docx.Document(str(tpl))
                tpl_marks[tpl_no] = set(bookmark_names(tpl_doc.element.body))
                tpl_fields[tpl_no] = scan_placeholders(tpl_doc)
            recorded = {b for b in read_recorded_bookmarks(xlsx).values() if b and b not in ("新增", "已删除")}
            if tpl_marks[tpl_no] and not recorded & tpl_marks[tpl_no]:
                record.update(status="skipped", reason=f"工作簿不是从模板{tpl_no}导出的（书签对不上）")
                log(f"[跳过] {xlsx.name}：书签与模板{tpl_no}（{tpl.name}）对不上，请先用该模板执行「导出表格到 Excel」")
                continue
            fields = tpl_fields[tpl_no] if "placeholders" in contents else []
            values = read_placeholder_values(xlsx) if fields else {}
            missing = [f for f in fields if not str(values.get(f, "")).strip()]
            if missing:
                record.update(status="skipped", reason="正文占位符未填：" + "、".join(missing))
                log(f"[跳过] {xlsx.name}：正文占位符未填 {'、'.join('{'+f+'}' for f in missing)}，"
                    f"请在该工作簿「正文信息」表填写后重试")
                continue
            output = out_root / f"{tpl.stem}_{xlsx.stem}.docx"
            seq = 2
            while output.exists():
                output = out_root / f"{tpl.stem}_{xlsx.stem}_{seq}.docx"
                seq += 1
            record["output"] = str(output)
            if dry_run:
                log(f"[模板·预演] 将生成 {output.name}（模板{tpl_no}：{tpl.name}）")
                for f in fields:
                    log(f"[占位符·预演] {{{f}}} → {values[f]}")
                if "tables" in contents or "marks" in contents:
                    run_sync_content(
                        tpl, xlsx,
                        tables="tables" in contents,
                        marks="marks" in contents,
                        keep_highlight=keep_highlight,
                        dry_run=True,
                        backup=False,
                    )
                record.update(status="ok", reason="预演完成")
                continue
            shutil.copy2(tpl, output)
            if fields:
                out_doc = _docx.Document(str(output))
                result = apply_placeholders(out_doc, values)
                out_doc.save(str(output))
                for f, v in result["replaced"]:
                    shown = v if len(v) <= 60 else v[:60] + "…"
                    log(f"[占位符] {{{f}}} → {shown}")
            if "tables" in contents or "marks" in contents:
                with tempfile.TemporaryDirectory(prefix="wesync_tpl_", dir=tmp_root()) as tmp:
                    tmp_xlsx = Path(tmp) / xlsx.name
                    shutil.copy2(xlsx, tmp_xlsx)
                    run_sync_content(
                        output, tmp_xlsx,
                        tables="tables" in contents,
                        marks="marks" in contents,
                        keep_highlight=keep_highlight,
                        backup=False,
                    )
            record.update(status="ok", reason="已生成")
            log(f"[模板] 已生成 {output}")
        except SyncError as exc:
            record["reason"] = str(exc)
            log(f"[失败] {xlsx.name}：{exc}")
        except Exception as exc:
            record["reason"] = f"{type(exc).__name__}: {exc}"
            log(f"[失败] {xlsx.name}：{record['reason']}")
    ok = sum(1 for r in results if r["status"] == "ok")
    skipped = sum(1 for r in results if r["status"] == "skipped")
    failed = sum(1 for r in results if r["status"] == "failed")
    log(f"模板生成{'预演' if dry_run else ''}结束：成功 {ok} 份，跳过 {skipped} 份，失败 {failed} 份（共 {total} 份）。输出目录：{out_root}")
    return results


def scan_blank_batch(docx_list: list) -> list[dict]:
    """只读批量扫描空白表：每份报告返回候选清单或错误原因，供界面按报告分组展示。"""
    results = []
    for docx in docx_list:
        path = Path(docx)
        item = {"docx": str(path), "candidates": [], "error": ""}
        try:
            item["candidates"] = scan_blank_tables(path)
        except SyncError as exc:
            item["error"] = str(exc)
        except Exception as exc:
            item["error"] = f"{type(exc).__name__}: {exc}"
        results.append(item)
    return results


def run_blank_batch(items: list[dict], *, dry_run: bool = False, force: bool = False) -> list[dict]:
    """批量删除空白表：items 每项 {"docx", "pick", "keep_caption_for", "reversible", "excel"}。

    逐份执行，单份失败不影响其他；口径与单文件删除完全一致。
    """
    results = []
    total = len(items)
    for i, item in enumerate(items, 1):
        path = Path(item["docx"])
        log(f"===== [{i}/{total}] {path.name} =====")
        record = {"docx": str(path), "status": "failed", "reason": ""}
        results.append(record)
        try:
            run_blank_tables(
                path,
                dry_run=dry_run,
                force=force,
                pick=item.get("pick"),
                keep_caption_for=item.get("keep_caption_for"),
                reversible=bool(item.get("reversible", True)),
                excel=item.get("excel"),
            )
            record.update(status="ok", reason="预演完成" if dry_run else "删除完成")
        except SyncError as exc:
            record["reason"] = str(exc)
            log(f"[失败] {path.name}：{exc}")
        except Exception as exc:
            record["reason"] = f"{type(exc).__name__}: {exc}"
            log(f"[失败] {path.name}：{record['reason']}")
    ok = sum(1 for r in results if r["status"] == "ok")
    failed = sum(1 for r in results if r["status"] == "failed")
    log(f"批量删除空白表{'预演' if dry_run else ''}结束：成功 {ok} 份，失败 {failed} 份（共 {total} 份）。")
    return results


def scan_report_format(
    docx: str | Path,
    template: str | Path | None = None,
    *,
    profile_name: str | None = None,
) -> dict:
    """只读扫描格式调整方案：按样板或已存纯格式方案预演，不写任何文件。"""
    from .report_format import scan_report_format as scan

    return scan(docx, template, profile_name=profile_name)


def run_report_format(
    docx: str | Path,
    template: str | Path | None = None,
    *,
    output: str | Path | None = None,
    dry_run: bool = False,
    profile_name: str | None = None,
) -> dict:
    """按样板或已存纯格式方案生成副本；默认不覆盖原报告，也不改写原文件。"""
    from .report_format import run_report_format as run

    return run(docx, template, output=output, dry_run=dry_run, profile_name=profile_name)


def list_format_profiles() -> list[dict]:
    """列出项目内保存的纯格式方案。"""
    from .report_format import list_format_profiles as list_profiles

    return list_profiles()


def load_format_profile(name: str):
    """读取并校验指定的纯格式方案。"""
    from .report_format import load_format_profile as load

    return load(name)


def save_format_profile(template: str | Path, name: str) -> dict:
    """从样板提取并保存纯格式参数，不保存报告文字、表格内容或样板路径。"""
    from .report_format import save_format_profile as save

    return save(template, name)


def delete_format_profile(name: str) -> dict:
    """删除指定的纯格式方案。"""
    from .report_format import delete_format_profile as delete

    return delete(name)


def scan_template_placeholders(docx: str | Path) -> list[str]:
    """只读列出模板正文/页眉页脚里的占位符字段名（{字段名} 写法）。"""
    import docx as _docx

    from .docx_placeholders import scan_placeholders

    path = resolve_docx(str(docx))
    return scan_placeholders(_docx.Document(str(path)))


def fill_placeholder_sheets(template: str | Path, workbooks: list) -> list[dict]:
    """按模板占位符给各工作簿生成/补齐「正文信息」填写骨架（保留已填值）。"""
    from .docx_placeholders import write_placeholder_sheet

    fields = scan_template_placeholders(template)
    results = []
    for xlsx in workbooks:
        path = Path(xlsx)
        record = {"excel": str(path), "status": "failed", "reason": "", "added": []}
        results.append(record)
        try:
            if not path.exists():
                record.update(status="skipped", reason="工作簿不存在")
                continue
            added = write_placeholder_sheet(path, fields)
            record.update(status="ok", added=added,
                          reason=f"新增 {len(added)} 个字段" if added else "字段已齐全")
            log(f"[正文信息] {path.name}：" + (f"已加入字段 {'、'.join(added)}，请在「正文信息」表填值" if added else "字段已齐全，无需改动"))
        except SyncError as exc:
            record["reason"] = str(exc)
            log(f"[失败] {path.name}：{exc}")
    if not fields:
        log("模板正文里没有占位符（{字段名} 写法），无需填写清单。")
    return results
