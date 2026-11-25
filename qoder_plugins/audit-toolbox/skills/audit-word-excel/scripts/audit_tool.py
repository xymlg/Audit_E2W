#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import re
import subprocess
import sys


class CliError(Exception):
    pass


def emit(payload: dict, code: int = 0) -> int:
    data = (json.dumps(payload, ensure_ascii=False) + "\n").encode("utf-8")
    sys.stdout.buffer.write(data)
    sys.stdout.buffer.flush()
    return code


def skill_version() -> str:
    text = (Path(__file__).resolve().parents[1] / "SKILL.md").read_text(encoding="utf-8")
    match = re.search(r"^version:\s*(\S+)\s*$", text, re.MULTILINE)
    return match.group(1) if match else "unknown"


def engine_command(explicit: str | None) -> list[str]:
    value = explicit or os.environ.get("AUDIT_TOOLBOX_ENGINE")
    if not value:
        bundled = Path(__file__).resolve().parents[3] / "engine" / "审计工具箱引擎.exe"
        if bundled.is_file():
            value = str(bundled)
    if not value:
        raise CliError("未配置审计工具箱引擎；请用 --engine 或 AUDIT_TOOLBOX_ENGINE 指定引擎绝对路径")
    engine = Path(value).expanduser()
    if not engine.is_absolute():
        raise CliError("引擎路径必须使用绝对路径")
    if not engine.is_file():
        raise CliError(f"引擎文件不存在：{engine}")
    raw_args = os.environ.get("AUDIT_TOOLBOX_ENGINE_ARGS", "[]")
    try:
        extra = json.loads(raw_args)
    except json.JSONDecodeError as exc:
        raise CliError(f"AUDIT_TOOLBOX_ENGINE_ARGS 不是合法 JSON：{exc.msg}") from exc
    if not isinstance(extra, list) or any(not isinstance(item, str) for item in extra):
        raise CliError("AUDIT_TOOLBOX_ENGINE_ARGS 必须是 JSON 字符串数组")
    return [str(engine), *extra]


def run_engine(command: list[str], request: dict) -> tuple[subprocess.CompletedProcess, list[dict], list[str], dict | None]:
    try:
        completed = subprocess.run(
            command,
            input=json.dumps(request, ensure_ascii=False) + "\n",
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
    except OSError as exc:
        raise CliError(f"引擎无法启动：{exc}") from exc
    events = []
    protocol_errors = []
    for line in completed.stdout.splitlines():
        if not line.strip():
            continue
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            protocol_errors.append(f"引擎返回了无法解析的数据：{line[:200]}")
            continue
        if isinstance(event, dict):
            events.append(event)
        else:
            protocol_errors.append("引擎返回了非对象 JSON 数据")
    results = [event for event in events if event.get("type") == "result"]
    if len(results) != 1:
        protocol_errors.append(f"预期 1 条 result 事件，实际收到 {len(results)} 条")
    return completed, events, protocol_errors, results[0] if len(results) == 1 else None


def request_from_args(args: argparse.Namespace) -> dict:
    if bool(args.json) == bool(args.file):
        raise CliError("request 必须且只能提供 --json 或 --file 之一")
    if args.json is not None:
        try:
            request = json.loads(args.json)
        except json.JSONDecodeError as exc:
            raise CliError(f"请求 JSON 格式错误：{exc.msg}") from exc
    else:
        try:
            request = json.loads(Path(args.file).read_text(encoding="utf-8-sig"))
        except OSError as exc:
            raise CliError(f"无法读取请求文件：{exc}") from exc
        except json.JSONDecodeError as exc:
            raise CliError(f"请求文件 JSON 格式错误：{exc.msg}") from exc
    if not isinstance(request, dict):
        raise CliError("请求必须是 JSON 对象")
    if not isinstance(request.get("operation"), str) or not request["operation"].strip():
        raise CliError("请求缺少 operation")
    return request


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="调用本地审计工具箱 JSONL 引擎")
    parser.add_argument("--engine", help="引擎可执行文件绝对路径；也可用 AUDIT_TOOLBOX_ENGINE")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("status", help="检查引擎版本")
    request = sub.add_parser("request", help="执行一个 JSON 请求")
    source = request.add_mutually_exclusive_group(required=True)
    source.add_argument("--json", help="JSON 请求字符串")
    source.add_argument("--file", help="UTF-8 JSON 请求文件")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        command = engine_command(args.engine)
        if args.command == "status":
            request = {"id": "skill-status", "operation": "version"}
        else:
            request = request_from_args(args)
        completed, events, protocol_errors, result = run_engine(command, request)
        logs = [
            {"level": event.get("level", "info"), "message": str(event.get("message", ""))}
            for event in events if event.get("type") == "log"
        ]
        engine_version = result.get("version") if result else None
        ok = bool(result and result.get("code") == 0 and completed.returncode == 0 and not protocol_errors)
        if args.command == "status":
            current = skill_version()
            matched = engine_version == current
            payload = {
                "ok": ok and matched,
                "available": ok,
                "engine": command[0],
                "engineVersion": engine_version,
                "skillVersion": current,
                "versionMatched": matched,
                "logs": logs[-200:],
                "protocolErrors": protocol_errors,
                "stderr": completed.stderr.strip()[-4000:],
            }
            if not ok:
                payload["error"] = "引擎检查失败，请查看日志"
            elif not matched:
                payload["error"] = f"Skill {current} 与引擎 {engine_version} 版本不一致"
            return emit(payload, 0 if payload["ok"] else 2)
        payload = {
            "ok": ok,
            "operation": request["operation"],
            "dryRun": request.get("dryRun") is True,
            "version": engine_version,
            "data": result.get("data") if result else None,
            "logs": logs[-200:],
            "logsTruncated": len(logs) > 200,
            "protocolErrors": protocol_errors,
            "stderr": completed.stderr.strip()[-4000:],
            "exitCode": completed.returncode,
        }
        if not ok:
            errors = [log["message"] for log in logs if log["level"] == "error"]
            payload["error"] = errors[-1] if errors else (protocol_errors[-1] if protocol_errors else "引擎执行失败")
        return emit(payload, 0 if ok else 2)
    except CliError as exc:
        return emit({"ok": False, "error": str(exc)}, 2)


if __name__ == "__main__":
    raise SystemExit(main())
