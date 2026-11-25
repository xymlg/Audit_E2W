import json
import os
from pathlib import Path
import subprocess
import sys
import textwrap

from helpers import ws_dir
from wesync import __version__


ROOT = Path(__file__).resolve().parents[2]
SKILL = ROOT / "agent_skills" / "audit-word-excel"
SCRIPT = SKILL / "scripts" / "audit_tool.py"
ENGINE = ROOT / "excel_word" / "qoder_engine.py"


def _env(*, engine: Path = sys.executable, args: list[str] | None = None) -> dict:
    env = os.environ.copy()
    env.pop("AUDIT_TOOLBOX_ENGINE", None)
    env.pop("AUDIT_TOOLBOX_ENGINE_ARGS", None)
    env["AUDIT_TOOLBOX_ENGINE"] = str(engine)
    env["AUDIT_TOOLBOX_ENGINE_ARGS"] = json.dumps(args if args is not None else [str(ENGINE)])
    return env


def _run(arguments: list[str], env: dict) -> tuple[subprocess.CompletedProcess, dict]:
    completed = subprocess.run(
        [sys.executable, str(SCRIPT), *arguments],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
        timeout=20,
        check=False,
    )
    payload = json.loads(completed.stdout)
    assert "Traceback" not in completed.stderr
    return completed, payload


def test_status_checks_skill_and_engine_versions():
    completed, payload = _run(["status"], _env())
    assert completed.returncode == 0
    assert payload["ok"] is True
    assert payload["available"] is True
    assert payload["versionMatched"] is True
    assert payload["skillVersion"] == payload["engineVersion"] == __version__


def test_request_accepts_inline_json_and_request_file():
    inline, inline_payload = _run(
        ["request", "--json", json.dumps({"operation": "version"})], _env()
    )
    assert inline.returncode == 0
    assert inline_payload["ok"] is True
    assert inline_payload["operation"] == "version"
    assert inline_payload["data"] is None

    request_file = ws_dir("skill_cli") / "request.json"
    request_file.write_text(json.dumps({"operation": "version"}), encoding="utf-8")
    from_file, file_payload = _run(
        ["request", "--file", str(request_file)], _env()
    )
    assert from_file.returncode == 0
    assert file_payload == inline_payload


def test_malformed_request_returns_json_error():
    completed, payload = _run(["request", "--json", "{"], _env())
    assert completed.returncode == 2
    assert payload["ok"] is False
    assert "JSON 格式错误" in payload["error"]


def test_missing_engine_and_bad_argument_json_are_reported():
    missing = ws_dir("skill_cli_missing") / "missing-engine.exe"
    completed, payload = _run(["status"], _env(engine=missing))
    assert completed.returncode == 2
    assert payload["ok"] is False
    assert "引擎文件不存在" in payload["error"]

    env = _env()
    env["AUDIT_TOOLBOX_ENGINE_ARGS"] = "not-json"
    completed, payload = _run(["status"], env)
    assert completed.returncode == 2
    assert "AUDIT_TOOLBOX_ENGINE_ARGS" in payload["error"]


def test_engine_failure_uses_error_log():
    missing_docx = ws_dir("skill_cli_failure") / "不存在.docx"
    request = json.dumps({"operation": "blank-scan", "docx": str(missing_docx)}, ensure_ascii=False)
    completed, payload = _run(["request", "--json", request], _env())
    assert completed.returncode == 2
    assert payload["ok"] is False
    assert payload["error"]
    assert payload["exitCode"] == 0


def test_protocol_without_result_is_not_success():
    fake = ws_dir("skill_cli_protocol") / "fake_engine.py"
    fake.write_text(textwrap.dedent("""
        import json
        import sys
        json.loads(sys.stdin.read())
        print(json.dumps({"type": "log", "level": "warning", "message": "only a log"}))
    """), encoding="utf-8")
    completed, payload = _run(
        ["request", "--json", '{"operation":"version"}'],
        _env(engine=sys.executable, args=[str(fake)]),
    )
    assert completed.returncode == 2
    assert payload["ok"] is False
    assert payload["logs"] == [{"level": "warning", "message": "only a log"}]
    assert "预期 1 条 result 事件" in payload["protocolErrors"][0]
