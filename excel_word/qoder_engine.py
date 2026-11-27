"""Qoder 插件无窗口引擎：从 stdin 读 JSON 请求，向 stdout 输出 JSONL 日志与结果。"""

from __future__ import annotations

import contextlib
import datetime as dt
import io
import json
import sys
from pathlib import Path

from wesync import __version__, api
from wesync.console import reset_log_state, set_log_sink
from wesync.errors import SyncError

QUERY_OPERATIONS = {
    "blank-scan", "deleted-list", "blank-batch-scan", "diff", "table-list",
    "sort-list", "template-scan", "backup-list", "format-scan",
    "format-heading-candidates",
}
RUN_OPERATIONS = {
    "export", "writeback", "bookmark", "rejudge", "clear-bookmarks", "marks",
    "marks-sync", "placeholders", "placeholders-sync", "blank-delete",
    "restore-tables", "blank-batch-delete", "diff-export", "incremental-export",
    "sort-set", "sort-clear", "batch-writeback", "template-batch",
    "fill-placeholders", "backup-restore", "format-apply",
}
DRY_RUN_OPERATIONS = {
    "export", "writeback", "bookmark", "rejudge", "clear-bookmarks", "marks",
    "marks-sync", "placeholders", "placeholders-sync", "blank-delete",
    "restore-tables", "blank-batch-delete", "batch-writeback", "template-batch", "incremental-export",
    "format-apply",
}


def _emit(request_id: str, kind: str, **payload) -> None:
    message = {"id": request_id, "type": kind, **payload}
    data = (json.dumps(message, ensure_ascii=False) + "\n").encode("utf-8")
    sys.stdout.buffer.write(data)
    sys.stdout.buffer.flush()


def _required_path(request: dict, key: str, suffixes: tuple[str, ...]) -> str:
    value = request.get(key)
    if not isinstance(value, str) or not value.strip():
        raise SyncError(f"缺少{key}路径")
    value = value.strip()
    if any(ord(ch) < 32 for ch in value):
        raise SyncError(f"{key}路径含有控制字符")
    path = Path(value).expanduser()
    if not path.is_absolute():
        raise SyncError(f"{key}必须使用绝对路径")
    if suffixes and path.suffix.lower() not in suffixes:
        raise SyncError(f"{key}只支持 {'/'.join(suffixes)} 文件")
    return str(path)


def _optional_path(request: dict, key: str, suffixes: tuple[str, ...]) -> str | None:
    value = request.get(key)
    if value in (None, ""):
        return None
    return _required_path(request, key, suffixes)


def _path_list(request: dict, key: str, suffixes: tuple[str, ...]) -> list[str]:
    values = request.get(key)
    if not isinstance(values, list) or not values:
        raise SyncError(f"请至少选择一个{key}文件")
    if len(values) > 200:
        raise SyncError(f"{key}一次最多处理 200 个文件")
    result = []
    for value in values:
        if not isinstance(value, str):
            raise SyncError(f"{key}必须是路径数组")
        result.append(_required_path({key: value}, key, suffixes))
    return result


def _text(request: dict, key: str, label: str) -> str:
    value = request.get(key)
    if not isinstance(value, str) or not value.strip():
        raise SyncError(f"请填写{label}")
    return value.strip()


def _flag(request: dict, key: str, default: bool = False) -> bool:
    value = request.get(key, default)
    if not isinstance(value, bool):
        raise SyncError(f"{key}必须是布尔值")
    return value


def _integer(request: dict, key: str, label: str, minimum: int, maximum: int) -> int:
    value = request.get(key)
    if isinstance(value, bool) or not isinstance(value, int):
        raise SyncError(f"{label}必须是整数")
    if not minimum <= value <= maximum:
        raise SyncError(f"{label}必须在 {minimum}～{maximum} 之间")
    return value


def _contents(request: dict, default: tuple[str, ...]) -> tuple[str, ...]:
    raw = request.get("contents") or {}
    if not isinstance(raw, dict):
        raise SyncError("同步内容格式不正确")
    selected = tuple(name for name in ("tables", "marks", "placeholders") if raw.get(name, name in default))
    if not selected:
        raise SyncError("同步内容至少勾选一项")
    return selected


def _jsonable(value):
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(key): _jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_jsonable(item) for item in value]
    if isinstance(value, set):
        return [_jsonable(item) for item in sorted(value, key=str)]
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _blank_candidates(items: list[dict]) -> list[dict]:
    keys = ("no", "index", "code", "kind", "nrows", "ncols", "caption_text")
    return [{key: item.get(key) for key in keys} for item in items]


def _template_items(request: dict, template_count: int) -> list[dict]:
    raw = request.get("items")
    if not isinstance(raw, list) or not raw:
        raise SyncError("请至少添加一个 Excel 工作簿")
    if len(raw) > 200:
        raise SyncError("一次最多生成 200 份报告")
    result = []
    for item in raw:
        if not isinstance(item, dict):
            raise SyncError("模板清单格式不正确")
        excel = _required_path(item, "excel", (".xlsx", ".xlsm"))
        template = item.get("template")
        if isinstance(template, bool) or not isinstance(template, int) or not 1 <= template <= template_count:
            raise SyncError("请为每个工作簿选择有效的所属模板")
        result.append({"excel": excel, "template": template})
    return result


def _blank_batch_items(request: dict) -> list[dict]:
    raw = request.get("items")
    if not isinstance(raw, list) or not raw:
        raise SyncError("请先批量扫描并勾选要删除的空白表格")
    if len(raw) > 200:
        raise SyncError("一次最多处理 200 份报告")
    result = []
    for item in raw:
        if not isinstance(item, dict):
            raise SyncError("批量删除清单格式不正确")
        docx = _required_path(item, "docx", (".docx", ".docm"))
        pick = _text(item, "pick", "候选编号")
        keep_for = item.get("keepCaptionFor")
        if keep_for is not None and not isinstance(keep_for, str):
            raise SyncError("保留标题编号必须是文本")
        excel = _optional_path(item, "excel", (".xlsx", ".xlsm"))
        reversible = item.get("reversible", True)
        if not isinstance(reversible, bool):
            raise SyncError("reversible 必须是布尔值")
        result.append({
            "docx": docx,
            "pick": pick,
            "keep_caption_for": keep_for,
            "reversible": reversible,
            "excel": excel,
        })
    return result


def _format_args(request: dict) -> dict:
    profile_name = request.get("profileName")
    if profile_name is not None and (not isinstance(profile_name, str) or not profile_name.strip()):
        raise SyncError("profileName 必须是非空文本")
    levels = request.get("headingLevels")
    if levels is not None:
        if not isinstance(levels, dict) or len(levels) > 300:
            raise SyncError("headingLevels 必须是候选编号到 1～4 级标题的映射，且最多 300 项")
        if any(not isinstance(key, str) or type(value) is not int or not 1 <= value <= 4 for key, value in levels.items()):
            raise SyncError("headingLevels 的级别必须是 1～4 的整数")
    batch_id = request.get("headingBatchId")
    if batch_id is not None and (not isinstance(batch_id, str) or not batch_id):
        raise SyncError("headingBatchId 必须是非空文本")
    role_batch_id = request.get("roleBatchId")
    if role_batch_id is not None and (not isinstance(role_batch_id, str) or not role_batch_id):
        raise SyncError("roleBatchId 必须是非空文本")
    role_overrides = request.get("roleOverrides")
    if role_overrides is not None:
        if not isinstance(role_overrides, dict) or len(role_overrides) > 300:
            raise SyncError("roleOverrides 必须是最多 300 项的疑难角色映射")
        if any(not isinstance(key, str) or value not in ("body", "table_title") for key, value in role_overrides.items()):
            raise SyncError("roleOverrides 的角色只能是 body 或 table_title")
    region_start_id = request.get("regionStartId")
    if region_start_id is not None and (not isinstance(region_start_id, str) or not region_start_id):
        raise SyncError("regionStartId 必须是非空文本")
    return {
        "template": _optional_path(request, "template", (".docx", ".docm")),
        "profile_name": profile_name.strip() if profile_name is not None else None,
        "scope": request.get("scope"),
        "heading_batch_id": batch_id,
        "heading_levels": levels,
        "role_batch_id": role_batch_id,
        "role_overrides": role_overrides,
        "region_start_id": region_start_id,
    }


def _dispatch(request: dict):
    operation = request["operation"]
    docx = _optional_path(request, "docx", (".docx", ".docm"))
    excel = _optional_path(request, "excel", (".xlsx", ".xlsm"))
    dry_run = _flag(request, "dryRun")
    force = _flag(request, "force")
    if dry_run and operation not in DRY_RUN_OPERATIONS:
        raise SyncError("该操作不支持预演")
    if operation in RUN_OPERATIONS and not dry_run and not _flag(request, "confirmed"):
        raise SyncError("正式写盘前必须取得用户确认，并传 confirmed=true")

    if operation in QUERY_OPERATIONS:
        if operation == "blank-scan":
            return _blank_candidates(api.scan_blank_tables(_required_path(request, "docx", (".docx", ".docm"))))
        if operation == "deleted-list":
            return api.list_deleted_blocks(_required_path(request, "docx", (".docx", ".docm")), excel)
        if operation == "blank-batch-scan":
            results = api.scan_blank_batch(_path_list(request, "paths", (".docx", ".docm")))
            for item in results:
                item["candidates"] = _blank_candidates(item.get("candidates", []))
            return results
        if operation == "diff":
            return api.scan_writeback_diff(_required_path(request, "docx", (".docx", ".docm")), excel)
        if operation == "table-list":
            return api.list_table_blocks(_required_path(request, "docx", (".docx", ".docm")), excel)
        if operation == "sort-list":
            return api.list_table_sorts(_required_path(request, "docx", (".docx", ".docm")), excel)
        if operation == "template-scan":
            return api.scan_template_placeholders(_required_path(request, "docx", (".docx", ".docm")))
        if operation == "backup-list":
            return api.scan_backups(_required_path(request, "docx", (".docx", ".docm")), excel)
        if operation == "format-heading-candidates":
            return api.scan_format_heading_candidates(_required_path(request, "docx", (".docx", ".docm")))
        if operation == "format-scan":
            return api.scan_report_format(
                _required_path(request, "docx", (".docx", ".docm")),
                **_format_args(request),
            )

    if operation == "format-apply":
        return api.run_report_format(
            _required_path(request, "docx", (".docx", ".docm")),
            output=_optional_path(request, "output", (".docx", ".docm")),
            dry_run=dry_run,
            **_format_args(request),
        )

    if operation == "backup-restore":
        backup = _required_path(request, "backupPath", (".docx", ".docm", ".xlsx", ".xlsm"))
        target = _required_path(request, "target", (".docx", ".docm", ".xlsx", ".xlsm"))
        if Path(backup).suffix.lower() != Path(target).suffix.lower():
            raise SyncError("备份文件与目标文件类型必须一致")
        return api.run_restore_backup(backup, target)

    if operation == "template-batch":
        templates = _path_list(request, "templates", (".docx", ".docm"))
        items = _template_items(request, len(templates))
        out_dir = _required_path(request, "outDir", ())
        return api.run_template_batch(
            templates,
            items,
            out_dir,
            dry_run=dry_run,
            contents=_contents(request, ("tables", "placeholders")),
            keep_highlight=_flag(request, "keepHighlight", True),
        )

    if operation == "fill-placeholders":
        template = _required_path(request, "docx", (".docx", ".docm"))
        return api.fill_placeholder_sheets(template, _path_list(request, "paths", (".xlsx", ".xlsm")))

    if operation == "blank-batch-delete":
        return api.run_blank_batch(_blank_batch_items(request), dry_run=dry_run, force=force)

    if operation == "batch-writeback":
        return api.run_batch_writeback(
            _path_list(request, "paths", (".docx", ".docm")),
            dry_run=dry_run,
            backup=_flag(request, "backup", True),
            force=force,
            contents=_contents(request, ("tables",)),
            keep_highlight=_flag(request, "keepHighlight", True),
        )

    if docx is None:
        raise SyncError("缺少 docx 路径")
    backup = _flag(request, "backup", not _flag(request, "noBackup"))
    if operation == "export":
        return api.run_export(
            docx, excel, dry_run=dry_run, force=force,
            with_excel_links=_flag(request, "withExcelLinks", True),
            totals_formula=_flag(request, "totalsFormula", True), backup=backup,
        )
    if operation == "writeback":
        contents = request.get("contents") or {"tables": True, "marks": False, "placeholders": False}
        return api.run_sync_content(
            docx, excel,
            tables=_flag(contents, "tables", True),
            marks=_flag(contents, "marks"),
            placeholders=_flag(contents, "placeholders"),
            keep_highlight=_flag(contents, "keepHighlight", True),
            dry_run=dry_run, force=force, backup=backup,
            changelog=_flag(request, "changelog"),
            delete_empty_rows=_flag(request, "deleteEmptyRows"),
        )
    if operation == "bookmark":
        return api.run_bookmark(
            docx, excel, dry_run=dry_run, force=force,
            with_excel_links=_flag(request, "withExcelLinks", True),
        )
    if operation == "rejudge":
        return api.run_rejudge(docx, excel, dry_run=dry_run, force=force, backup=backup)
    if operation == "clear-bookmarks":
        return api.run_clear_bookmarks(docx, dry_run=dry_run, force=force)
    if operation == "marks":
        return api.run_marks(docx, excel, dry_run=dry_run, force=force)
    if operation == "marks-sync":
        return api.run_marks_sync(
            docx, excel, dry_run=dry_run, force=force,
            keep_highlight=_flag(request, "keepHighlight"),
        )
    if operation == "placeholders":
        return api.run_placeholder_export(docx, excel, dry_run=dry_run, force=force)
    if operation == "placeholders-sync":
        return api.run_placeholder_sync(docx, excel, dry_run=dry_run, force=force)
    if operation == "blank-delete":
        return api.run_blank_tables(
            docx, dry_run=dry_run, force=force,
            pick=_text(request, "pick", "候选编号"),
            keep_caption=_flag(request, "keepCaption"),
            keep_caption_for=request.get("keepCaptionFor"),
            reversible=_flag(request, "reversible", True), excel=excel,
        )
    if operation == "restore-tables":
        return api.run_restore_tables(
            docx, excel, pick=_text(request, "pick", "恢复编号"),
            dry_run=dry_run, force=force,
        )
    if operation == "diff-export":
        records = api.scan_writeback_diff(docx, excel)
        return api.export_diff_report(docx, excel, records)
    if operation == "incremental-export":
        return api.run_incremental_export(
            docx, excel, dry_run=dry_run, force=force,
            with_excel_links=_flag(request, "withExcelLinks", True),
            totals_formula=_flag(request, "totalsFormula", True),
            backup=_flag(request, "backup", True),
        )
    if operation == "sort-set":
        return api.set_table_sort(
            docx, excel,
            code=_text(request, "code", "表块编号"),
            column=_text(request, "column", "排序列"),
            descending=_flag(request, "descending", True),
        )
    if operation == "sort-clear":
        return api.clear_table_sort(docx, excel, code=_text(request, "code", "表块编号"))
    raise SyncError(f"不支持的操作：{operation}")


def _run(request: dict, request_id: str) -> int:
    operation = request.get("operation")
    if operation == "version":
        _emit(request_id, "result", code=0, version=__version__)
        return 0
    if operation not in QUERY_OPERATIONS | RUN_OPERATIONS:
        _emit(request_id, "log", level="error", message=f"不支持的操作：{operation}")
        _emit(request_id, "result", code=2, operation=str(operation or "unknown"), version=__version__)
        return 2

    stderr = io.StringIO()

    def sink(message: str) -> None:
        level = "warning" if message.startswith("警告：") else "info"
        _emit(request_id, "log", level=level, message=message)

    set_log_sink(sink)
    reset_log_state()
    try:
        with contextlib.redirect_stderr(stderr):
            data = _dispatch(request)
        code = 0
    except (SyncError, OSError, ValueError) as exc:
        _emit(request_id, "log", level="error", message=str(exc))
        data = None
        code = 2
    except Exception as exc:
        _emit(request_id, "log", level="error", message=f"引擎异常：{type(exc).__name__}: {exc}")
        data = None
        code = 2
    finally:
        set_log_sink(None)
        reset_log_state()

    for line in stderr.getvalue().splitlines():
        if line.strip():
            _emit(request_id, "log", level="error", message=line.strip())
    _emit(request_id, "result", code=int(code), operation=operation, version=__version__, data=_jsonable(data))
    return int(code)


def main() -> int:
    for raw_line in sys.stdin.buffer:
        line = raw_line.decode("utf-8")
        if not line.strip():
            continue
        try:
            request = json.loads(line)
            if not isinstance(request, dict):
                raise ValueError("请求必须是 JSON 对象")
            request_id = str(request.get("id") or "engine")
        except (ValueError, TypeError) as exc:
            _emit("engine", "log", level="error", message=f"请求格式错误：{exc}")
            _emit("engine", "result", code=2, operation="unknown", version=__version__)
            continue
        _run(request, request_id)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
