from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SKILL_NAME = "audit-word-excel"
SKILL_ROOT = ROOT / "agent_skills" / SKILL_NAME
PLUGIN_ROOT = ROOT / "qoder_plugins" / "audit-toolbox"
PLUGIN_SKILL = PLUGIN_ROOT / "skills" / SKILL_NAME
DIST = ROOT / "dist" / "agent_skills"
TMP = ROOT / ".tmp" / "agent_skills_stage"
FORBIDDEN = {".env", ".wesync_ui.json", "node_modules", "__pycache__", ".git", ".DS_Store"}
REQUIRED_FRONTMATTER = (
    "name", "display_name", "display_name_en", "description", "description_zh",
    "description_en", "version", "author", "allowed-tools",
)
REQUIRED_FILES = (
    "SKILL.md",
    "scripts/audit_tool.py",
    "references/operations.md",
    "references/local-cli.md",
    "references/platforms.md",
)


def sync_skill_version(version: str) -> None:
    path = SKILL_ROOT / "SKILL.md"
    text = path.read_text(encoding="utf-8")
    updated, count = re.subn(r"^version:\s*\S+\s*$", f"version: {version}", text, count=1, flags=re.MULTILINE)
    if count != 1:
        raise RuntimeError("无法更新通用 Skill 版本")
    path.write_text(updated, encoding="utf-8")


def sync_qoder_skill() -> None:
    if not SKILL_ROOT.is_dir():
        raise RuntimeError(f"通用 Skill 不存在：{SKILL_ROOT}")
    PLUGIN_SKILL.mkdir(parents=True, exist_ok=True)
    for source in sorted(SKILL_ROOT.rglob("*")):
        if not source.is_file():
            continue
        relative = source.relative_to(SKILL_ROOT)
        target = PLUGIN_SKILL / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)


def validate_skill(root: Path = SKILL_ROOT) -> dict:
    skill_path = root / "SKILL.md"
    text = skill_path.read_text(encoding="utf-8")
    match = re.match(r"^---\r?\n([\s\S]*?)\r?\n---\r?\n", text)
    if not match:
        raise RuntimeError("SKILL.md 缺少 YAML frontmatter")
    frontmatter = match.group(1)
    for field in REQUIRED_FRONTMATTER:
        if not re.search(rf"^{re.escape(field)}:\s*\S.*$", frontmatter, re.MULTILINE):
            raise RuntimeError(f"SKILL.md 缺少 {field}")
    if not re.search(rf"^name:\s*{re.escape(SKILL_NAME)}\s*$", frontmatter, re.MULTILINE):
        raise RuntimeError("Skill name 与目录名不一致")
    version = re.search(r"^version:\s*(\S+)\s*$", frontmatter, re.MULTILINE).group(1)
    if not re.fullmatch(r"\d+\.\d+\.\d+", version):
        raise RuntimeError("Skill version 格式错误")
    for field in ("description", "description_zh", "description_en"):
        value = re.search(rf"^{field}:\s*(.*)$", frontmatter, re.MULTILINE).group(1)
        if len(value) > 500:
            raise RuntimeError(f"{field} 超过 500 字符")
    body = text[match.end():]
    if len(body.splitlines()) >= 500:
        raise RuntimeError("SKILL.md 正文达到或超过 500 行")
    if "TODO:" in text:
        raise RuntimeError("SKILL.md 含未完成的 TODO")
    for relative in REQUIRED_FILES:
        if not (root / relative).is_file():
            raise RuntimeError(f"Skill 缺少文件：{relative}")
    for reference in re.findall(r"`(references/[^`]+)`", body):
        if not (root / reference).is_file():
            raise RuntimeError(f"Skill 引用的文件不存在：{reference}")
    compile((root / "scripts" / "audit_tool.py").read_text(encoding="utf-8"), "audit_tool.py", "exec")
    for file in sorted(root.rglob("*")):
        if file.name in FORBIDDEN or file.name.startswith("~$"):
            raise RuntimeError(f"Skill 含禁止打包的文件：{file.relative_to(root)}")
        if file.is_file() and file.suffix.lower() in {".exe", ".dll", ".pyd", ".pyc"}:
            raise RuntimeError(f"通用 Skill 不应包含二进制运行时：{file.relative_to(root)}")
    return {"name": SKILL_NAME, "version": version, "files": len([p for p in root.rglob("*") if p.is_file()])}


def _validate_archive(archive: Path, version: str) -> list[str]:
    with zipfile.ZipFile(archive) as zf:
        names = zf.namelist()
        expected = [f"{SKILL_NAME}/{relative}" for relative in REQUIRED_FILES]
        for name in expected:
            if name not in names:
                raise RuntimeError(f"ZIP 缺少 {name}：{archive.name}")
        for name in names:
            path = Path(name)
            if name.startswith(("/", "\\")) or ".." in path.parts:
                raise RuntimeError(f"ZIP 含不安全路径：{name}")
            if path.name in FORBIDDEN or path.name.startswith("~$"):
                raise RuntimeError(f"ZIP 含禁止文件：{name}")
            if path.suffix.lower() in {".exe", ".dll", ".pyd", ".pyc"}:
                raise RuntimeError(f"ZIP 含二进制运行时：{name}")
        text = zf.read(f"{SKILL_NAME}/SKILL.md").decode("utf-8")
        if f"version: {version}" not in text:
            raise RuntimeError(f"ZIP Skill 版本不是 {version}：{archive.name}")
        return names


def build_archives(version: str) -> list[Path]:
    if TMP.exists():
        shutil.rmtree(TMP)
    stage = TMP / SKILL_NAME
    shutil.copytree(SKILL_ROOT, stage, ignore=shutil.ignore_patterns("__pycache__", "node_modules"))
    validate_skill(stage)
    DIST.mkdir(parents=True, exist_ok=True)
    archives = []
    try:
        for platform in ("workbuddy", "dsh"):
            archive = DIST / f"审计工具箱_skill_{platform}_v{version}.zip"
            with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
                for file in sorted(stage.rglob("*")):
                    if file.is_file():
                        zf.write(file, (Path(SKILL_NAME) / file.relative_to(stage)).as_posix())
            _validate_archive(archive, version)
            archives.append(archive)
            print(f"已生成并校验：{archive}")
    finally:
        shutil.rmtree(TMP)
    return archives


def run_qoder_check() -> None:
    subprocess.run(
        ["node", str(PLUGIN_ROOT / "scripts" / "check.mjs"), str(PLUGIN_ROOT)],
        cwd=ROOT,
        check=True,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="校验并构建 WorkBuddy/DSH 通用审计 Skill 包")
    parser.add_argument("--sync-qoder", action="store_true", help="把通用 Skill 同步到 Qoder 插件")
    parser.add_argument("--build-zips", action="store_true", help="生成 WorkBuddy 与 DSH ZIP")
    parser.add_argument("--check-only", action="store_true", help="只校验通用 Skill")
    args = parser.parse_args()
    info = validate_skill()
    print(f"通用 Skill 校验通过：{info['name']} v{info['version']}，{info['files']} 个文件")
    if args.sync_qoder:
        sync_qoder_skill()
        run_qoder_check()
        print("Qoder Skill 已同步并通过插件检查")
    if args.check_only:
        return 0
    if args.build_zips:
        build_archives(info["version"])
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (RuntimeError, subprocess.CalledProcessError) as exc:
        print(f"Skill 构建失败：{exc}", file=sys.stderr)
        raise SystemExit(1)
