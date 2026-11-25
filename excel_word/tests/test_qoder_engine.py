import json
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from helpers import ws_dir
import qoder_engine
from wesync.docx_placeholders import write_placeholder_sheet
from wesync.errors import SyncError
from wesync.fsops import load_workbook_for_save
from wesync.valuefmt import formula_result


def test_formula_result_passes_cache_to_formula_chain():
    wb = Workbook()
    ws = wb.active
    ws["A1"] = 10
    ws["A2"] = "=A1*2"
    cached = Workbook().active
    cached["A1"] = 7
    assert formula_result(ws, 2, 1, cached) == 14
    wb.close()
    cached.parent.close()


def _xlsm_with_macro(path: Path) -> None:
    wb = Workbook()
    wb.active["A1"] = "保留宏"
    wb.save(str(path))
    wb.close()
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("xl/vbaProject.bin", b"fake-vba-project")


def test_load_workbook_for_save_preserves_vba():
    path = ws_dir("qoder_xlsm") / "宏工作簿.xlsm"
    _xlsm_with_macro(path)
    wb = load_workbook_for_save(path)
    wb.active["A2"] = "工具写入"
    wb.save(str(path))
    wb.close()
    wb.vba_archive.close()
    with zipfile.ZipFile(path) as archive:
        assert archive.read("xl/vbaProject.bin") == b"fake-vba-project"


def test_placeholder_writer_preserves_vba():
    path = ws_dir("qoder_xlsm_placeholder") / "宏工作簿.xlsm"
    _xlsm_with_macro(path)
    assert write_placeholder_sheet(path, ["公司名称"]) == ["公司名称"]
    with zipfile.ZipFile(path) as archive:
        assert archive.read("xl/vbaProject.bin") == b"fake-vba-project"
    wb = load_workbook(str(path), keep_vba=True)
    assert wb["正文信息"]["A2"].value == "公司名称"
    wb.close()
    wb.vba_archive.close()


def test_blank_query_sanitizes_xml_and_remains_read_only(monkeypatch):
    path = ws_dir("qoder_query") / "报告.docx"
    path.write_bytes(b"original")
    candidate = {"no": 1, "index": 0, "code": "T01", "kind": "all-empty", "nrows": 2, "ncols": 2, "caption_text": "", "_element": object()}
    monkeypatch.setattr(qoder_engine.api, "scan_blank_tables", lambda value: [candidate])
    result = qoder_engine._dispatch({"operation": "blank-scan", "docx": str(path)})
    assert result == [{key: candidate[key] for key in ("no", "index", "code", "kind", "nrows", "ncols", "caption_text")}]
    assert path.read_bytes() == b"original"


def test_batch_query_sanitizes_each_candidate(monkeypatch):
    path = str(ws_dir("qoder_batch_query") / "报告.docx")
    monkeypatch.setattr(qoder_engine.api, "scan_blank_batch", lambda paths: [{"docx": paths[0], "candidates": [{"no": 1, "_element": object()}], "error": ""}])
    result = qoder_engine._dispatch({"operation": "blank-batch-scan", "paths": [path]})
    assert result[0]["candidates"] == [{"no": 1, "index": None, "code": None, "kind": None, "nrows": None, "ncols": None, "caption_text": None}]


def test_template_items_validate_mapping(monkeypatch):
    d = ws_dir("qoder_template")
    template = d / "模板.docx"
    template.write_bytes(b"template")
    excel = d / "数据.xlsx"
    excel.write_bytes(b"excel")
    captured = {}

    def fake_batch(templates, items, out_dir, **kwargs):
        captured.update(templates=templates, items=items, out_dir=out_dir, kwargs=kwargs)
        return [{"excel": items[0]["excel"], "status": "ok", "output": str(d / "结果.docx")}]

    monkeypatch.setattr(qoder_engine.api, "run_template_batch", fake_batch)
    result = qoder_engine._dispatch({
        "operation": "template-batch", "templates": [str(template)],
        "items": [{"excel": str(excel), "template": 1}], "outDir": str(d),
        "dryRun": True, "contents": {"tables": True, "placeholders": True},
    })
    assert result[0]["status"] == "ok"
    assert captured["kwargs"]["dry_run"] is True
    with pytest.raises(SyncError, match="模板"):
        qoder_engine._template_items({"items": [{"excel": str(excel), "template": 2}]}, 1)


def test_blank_batch_items_use_engine_field_names():
    path = str(ws_dir("qoder_batch_items") / "报告.docx")
    items = qoder_engine._blank_batch_items({
        "items": [{"docx": path, "pick": "1", "keepCaptionFor": "1", "reversible": True}]
    })
    assert items == [{"docx": path, "pick": "1", "keep_caption_for": "1", "reversible": True, "excel": None}]


def test_blank_batch_delete_forwards_force(monkeypatch):
    path = str(ws_dir("qoder_batch_force") / "报告.docx")
    captured = {}

    def fake_batch(items, **kwargs):
        captured.update(items=items, kwargs=kwargs)
        return []

    monkeypatch.setattr(qoder_engine.api, "run_blank_batch", fake_batch)
    qoder_engine._dispatch({
        "operation": "blank-batch-delete",
        "items": [{"docx": path, "pick": "1"}],
        "dryRun": True,
        "force": True,
    })
    assert captured["kwargs"] == {"dry_run": True, "force": True}


def test_blank_batch_api_forwards_force(monkeypatch):
    from wesync import api

    captured = {}

    def fake_delete(path, **kwargs):
        captured.update(kwargs)
        return 0

    monkeypatch.setattr(api, "run_blank_tables", fake_delete)
    api.run_blank_batch([{"docx": r"D:\报告.docx", "pick": "1", "keep_caption_for": "2"}], dry_run=True, force=True)
    assert captured["force"] is True
    assert captured["dry_run"] is True
    assert captured["keep_caption_for"] == "2"


def test_reversible_blank_plan_forces_workbook_check(monkeypatch):
    from wesync.commands import blank_cmd
    from wesync.config import SHEET_NAME

    path = ws_dir("qoder_blank_force") / "报告_表格数据.xlsx"
    path.write_bytes(b"workbook")
    wb = Workbook()
    wb.active.title = SHEET_NAME
    checked = []
    monkeypatch.setattr(blank_cmd, "check_file_free", lambda value, force: checked.append((value, force)))
    monkeypatch.setattr(blank_cmd, "load_workbook_for_save", lambda value: wb)
    monkeypatch.setattr(blank_cmd, "parse_blocks", lambda ws: ([], {}))
    monkeypatch.setattr(blank_cmd, "table_index_map_by_bookmark", lambda doc, names: {})
    assert blank_cmd._plan_reversible(object(), path, [], force=True)[0] is wb
    assert checked == [(path, True)]
    wb.close()


def test_query_rejects_dry_run_and_backup_restore_requires_same_type():
    with pytest.raises(SyncError, match="不支持预演"):
        qoder_engine._dispatch({"operation": "diff", "docx": r"D:\a.docx", "dryRun": True})
    with pytest.raises(SyncError, match="类型必须一致"):
        qoder_engine._dispatch({"operation": "backup-restore", "backupPath": r"D:\a.docx", "target": r"D:\a.xlsx"})


def test_jsonable_serializes_paths_and_dates():
    path = Path(r"D:\报告.docx")
    assert qoder_engine._jsonable({"path": path, "time": datetime(2026, 11, 24, 9, 30)}) == {
        "path": str(path), "time": "2026-11-24T09:30:00",
    }


def test_engine_process_version_protocol():
    root = Path(__file__).resolve().parents[1]
    completed = subprocess.run(
        [sys.executable, "qoder_engine.py"],
        cwd=root,
        input=json.dumps({"id": "version-check", "operation": "version"}) + "\n",
        text=True,
        encoding="utf-8",
        capture_output=True,
        check=True,
        timeout=20,
    )
    event = json.loads(completed.stdout.strip())
    assert event["id"] == "version-check"
    assert event["type"] == "result"
    assert event["code"] == 0
    assert event["version"] == qoder_engine.__version__
