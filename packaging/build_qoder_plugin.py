from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import zipfile

from build_agent_skills import sync_qoder_skill, sync_skill_version

ROOT = Path(__file__).resolve().parents[1]
EXCEL_ROOT = ROOT / "excel_word"
PLUGIN_ROOT = ROOT / "qoder_plugins" / "audit-toolbox"
PYTHON = ROOT / ".venv" / "Scripts" / "python.exe"
DIST = ROOT / "dist"
TMP = ROOT / ".tmp"


def replace_once(path: Path, pattern: str, replacement: str) -> None:
    text = path.read_text(encoding="utf-8")
    updated, count = re.subn(pattern, replacement, text, count=1, flags=re.MULTILINE)
    if count != 1:
        raise RuntimeError(f"无法在 {path} 中更新版本号")
    path.write_text(updated, encoding="utf-8")


def sync_versions() -> str:
    init_text = (EXCEL_ROOT / "wesync" / "__init__.py").read_text(encoding="utf-8")
    match = re.search(r'^__version__\s*=\s*"([^"]+)"', init_text, re.MULTILINE)
    if not match:
        raise RuntimeError("未找到 wesync.__version__")
    version = match.group(1)
    replace_once(ROOT / "packaging" / "wxsync.spec", r'^VERSION\s*=\s*"[^"]+"', f'VERSION = "{version}"')
    replace_once(ROOT / "packaging" / "installer.iss", r'^#define MyAppVersion "[^"]+"', f'#define MyAppVersion "{version}"')
    replace_once(ROOT / "packaging" / "使用说明.txt", r"^审计工具箱 v[^（]+", f"审计工具箱 v{version}")
    for relative in (".qoder-plugin/plugin.json", "package.json"):
        path = PLUGIN_ROOT / relative
        data = json.loads(path.read_text(encoding="utf-8"))
        data["version"] = version
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    sync_skill_version(version)
    sync_qoder_skill()
    print(f"版本已同步：{version}")
    return version


def run(command: list[str], cwd: Path = ROOT) -> None:
    env = os.environ.copy()
    env["TMP"] = str(TMP)
    env["TEMP"] = str(TMP)
    print("运行：" + " ".join(str(item) for item in command), flush=True)
    subprocess.run(command, cwd=cwd, env=env, check=True)


def validate_source() -> None:
    run(["node", str(PLUGIN_ROOT / "scripts" / "check.mjs"), str(PLUGIN_ROOT)])


def stage_release_files(version: str) -> Path:
    built = DIST / f"审计工具箱_v{version}"
    if not built.is_dir():
        raise RuntimeError(f"便携版目录不存在：{built}")
    shutil.copy2(ROOT / "packaging" / "使用说明.txt", built / "使用说明.txt")
    return built


def build_executables() -> None:
    run([
        str(PYTHON), "-m", "PyInstaller", "--noconfirm",
        "--distpath", str(DIST), "--workpath", str(ROOT / "build"),
        str(ROOT / "packaging" / "wxsync.spec"),
    ])


def build_qoder_engine(version: str) -> Path:
    dist = TMP / "qoder_engine_dist"
    run([
        str(PYTHON), "-m", "PyInstaller", "--noconfirm",
        "--distpath", str(dist), "--workpath", str(TMP / "qoder_engine_build"),
        str(ROOT / "packaging" / "qoder_engine.spec"),
    ])
    built = dist / f"审计工具箱_qoder_engine_v{version}"
    if not built.is_dir():
        raise RuntimeError(f"未找到 Qoder 引擎产物：{built}")
    return built


def package_plugin(version: str, engine_built: Path) -> Path:
    engine = engine_built / "审计工具箱引擎.exe"
    internal = engine_built / "_internal"
    if not engine.is_file() or not internal.is_dir():
        raise RuntimeError(f"打包产物缺少引擎或运行时：{engine_built}")

    stage = TMP / "qoder_plugin_stage" / version
    if stage.exists():
        shutil.rmtree(stage)
    stage.mkdir(parents=True)
    for item in PLUGIN_ROOT.iterdir():
        if item.name == "engine":
            continue
        target = stage / item.name
        if item.is_dir():
            shutil.copytree(item, target, ignore=shutil.ignore_patterns("__pycache__", "node_modules"))
        else:
            shutil.copy2(item, target)
    engine_dir = stage / "engine"
    engine_dir.mkdir()
    shutil.copy2(engine, engine_dir / engine.name)
    shutil.copytree(internal, engine_dir / "_internal")

    run(["node", str(stage / "scripts" / "check.mjs"), str(stage)])
    output_dir = DIST / "qoder_plugins"
    output_dir.mkdir(parents=True, exist_ok=True)
    archive = output_dir / f"Audit_E2W_qoder_v{version}.zip"
    with zipfile.ZipFile(archive, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for file in sorted(stage.rglob("*")):
            if file.is_file():
                zf.write(file, file.relative_to(stage).as_posix())
    size = archive.stat().st_size
    if size > 20_000_000:
        raise RuntimeError(f"Qoder 插件 ZIP 超过发布页 20 MB 限制：{size} 字节")
    with zipfile.ZipFile(archive) as zf:
        names = zf.namelist()
        if ".qoder-plugin/plugin.json" not in names:
            raise RuntimeError("插件 ZIP 根目录缺少 manifest")
        if any(name.startswith(("/", "\\")) or ".." in Path(name).parts for name in names):
            raise RuntimeError("插件 ZIP 含有不安全路径")
    shutil.rmtree(stage)
    print(f"Qoder 插件已生成：{archive}")
    return archive


def build_installer() -> Path:
    compiler = ROOT / ".tools" / "Inno Setup 6" / "ISCC.exe"
    if not compiler.is_file():
        raise RuntimeError(f"未找到 Inno Setup 编译器：{compiler}")
    run([str(compiler), str(ROOT / "packaging" / "installer.iss")])
    version = re.search(r'^#define MyAppVersion "([^"]+)"', (ROOT / "packaging" / "installer.iss").read_text(encoding="utf-8"), re.MULTILINE)
    if not version:
        raise RuntimeError("无法读取安装程序版本")
    installer = DIST / f"审计工具箱_v{version.group(1)}_安装程序.exe"
    if not installer.is_file():
        raise RuntimeError(f"未找到安装程序产物：{installer}")
    print(f"安装程序已生成：{installer}")
    return installer


def main() -> int:
    parser = argparse.ArgumentParser(description="同步版本并构建审计工具箱便携版、Qoder 插件和安装程序")
    parser.add_argument("--skip-exe", action="store_true", help="复用已有 PyInstaller 产物")
    parser.add_argument("--skip-installer", action="store_true", help="不编译 Inno Setup 安装程序")
    parser.add_argument("--check-only", action="store_true", help="只同步版本并做静态检查")
    args = parser.parse_args()
    if not PYTHON.is_file():
        raise RuntimeError(f"未找到项目虚拟环境：{PYTHON}")
    TMP.mkdir(exist_ok=True)
    version = sync_versions()
    validate_source()
    if args.check_only:
        return 0
    if not args.skip_exe:
        build_executables()
    stage_release_files(version)
    engine_built = build_qoder_engine(version)
    archive = package_plugin(version, engine_built)
    if not args.skip_installer:
        build_installer()
    print(f"完成：{archive}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (RuntimeError, subprocess.CalledProcessError) as exc:
        print(f"构建失败：{exc}", file=sys.stderr)
        raise SystemExit(1)
