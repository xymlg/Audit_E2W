from __future__ import annotations

import argparse
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import stat
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

PLUGIN_SOURCE_FILES = (
    ".mcp.json",
    ".qoder-plugin/plugin.json",
    "README.md",
    "app.html",
    "app.js",
    "assets/icon.png",
    "catalog.mjs",
    "commands/batch.md",
    "commands/check.md",
    "commands/export.md",
    "commands/format.md",
    "commands/sync.md",
    "package.json",
    "scripts/check.mjs",
    "scripts/smoke.mjs",
    "server.mjs",
    "skills/audit-word-excel/SKILL.md",
    "skills/audit-word-excel/references/local-cli.md",
    "skills/audit-word-excel/references/operations.md",
    "skills/audit-word-excel/references/platforms.md",
    "skills/audit-word-excel/scripts/audit_tool.py",
)
PORTABLE_TOP_LEVEL = {"审计工具箱.exe", "审计工具箱引擎.exe", "使用说明.txt", "_internal"}
RUNTIME_OFFICE_TEMPLATE = "_internal/docx/templates/default.docx"
OFFICE_SUFFIXES = {
    ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".docx~", ".rtf",
    ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm", ".xlam", ".xlsx~", ".csv",
    ".ppt", ".pptx", ".pptm", ".pot", ".potx", ".potm", ".pps", ".ppsx", ".ppsm", ".ppam",
    ".odt", ".ods", ".odp", ".one", ".onetoc", ".pub", ".accdb", ".mdb", ".vsd", ".vsdx", ".vssx",
}
FORBIDDEN_PARTS = {
    "_backup", ".git", "__pycache__", "node_modules", ".pytest_cache", ".mypy_cache",
    ".ruff_cache", ".cache", ".tmp", ".venv", "build", "dist",
}
LOCK_SUFFIXES = {".lock", ".lck", ".tmp", ".temp", ".pyc", ".pyo"}
CACHE_FILES = {".ds_store", "thumbs.db", "desktop.ini"}
INSTALLER_SOURCES = (
    r"{#MySrcDir}\{#MyExeName}",
    r"{#MySrcDir}\审计工具箱引擎.exe",
    r"{#MySrcDir}\使用说明.txt",
    r"{#MySrcDir}\_internal\*",
    r"{#MySrcDir}\_internal\docx\templates\default.docx",
)
INSTALLER_EXCLUDES = (
    "*.doc*", "*.xls*", "*.ppt*", "*.pot*", "*.pps*", "*.ppam", "*.one*", "*.acc*",
    "*.mdb", "*.vs*", "*.od?", "*.pub", "*.rtf", "*.csv",
    "~$*", ".~lock*", "*.lock", "*.lck", "*.tmp*", "*.temp", "*.pyc", "*.pyo",
    "*_backup*", "*.git*", "*cache*", "*node_modules*", "*.venv*", r"build\*", r"dist\*",
    ".DS_Store", "Thumbs.db", "desktop.ini",
)


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
    replace_once(ROOT / "package.json", r'^  "version": "[^"]+"', f'  "version": "{version}"')
    for path in (ROOT / "pyproject.toml", EXCEL_ROOT / "pyproject.toml"):
        if path.is_file():
            replace_once(path, r'^version\s*=\s*"[^"]+"', f'version = "{version}"')
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


def validate_relative_path(relative: PurePosixPath) -> None:
    parts = relative.parts
    if not parts or relative.is_absolute() or any(part in {"", ".", ".."} for part in parts):
        raise RuntimeError(f"不安全的打包路径：{relative}")
    if any(part.casefold() in FORBIDDEN_PARTS for part in parts):
        raise RuntimeError(f"打包内容含备份、缓存或版本库目录：{relative}")
    name = relative.name
    if name.startswith(("~$", ".~lock.")) or relative.suffix.lower() in LOCK_SUFFIXES:
        raise RuntimeError(f"打包内容含锁文件或缓存：{relative}")
    if name.casefold() in CACHE_FILES:
        raise RuntimeError(f"打包内容含系统缓存：{relative}")
    relative_name = relative.as_posix()
    is_runtime_template = relative_name in {RUNTIME_OFFICE_TEMPLATE, f"engine/{RUNTIME_OFFICE_TEMPLATE}"}
    if not is_runtime_template and relative.suffix.lower() in OFFICE_SUFFIXES:
        raise RuntimeError(f"打包内容含 Office 文档：{relative}")


def validate_tree(root: Path, allowed_files: set[str] | None = None) -> list[Path]:
    if not root.is_dir() or root.is_symlink() or root.is_junction():
        raise RuntimeError(f"打包源目录无效或为链接：{root}")
    allowed_dirs = {
        parent.as_posix()
        for name in allowed_files or ()
        for parent in PurePosixPath(name).parents
        if parent.as_posix() != "."
    }
    files: list[Path] = []
    for current, directories, names in os.walk(root, followlinks=False):
        directories.sort()
        names.sort()
        for name in directories + names:
            path = Path(current) / name
            relative = PurePosixPath(path.relative_to(root).as_posix())
            if path.is_symlink() or path.is_junction():
                raise RuntimeError(f"打包内容含符号链接或 junction：{relative}")
            validate_relative_path(relative)
            if allowed_files is not None and relative.as_posix() not in allowed_files | allowed_dirs:
                raise RuntimeError(f"打包内容不在白名单：{relative}")
            if path.is_file():
                files.append(path)
            elif not path.is_dir():
                raise RuntimeError(f"打包内容不是普通文件或目录：{relative}")
    if allowed_files is not None:
        actual = {path.relative_to(root).as_posix() for path in files}
        missing = allowed_files - actual
        if missing:
            raise RuntimeError(f"打包源缺少白名单文件：{', '.join(sorted(missing))}")
    return files


def validate_installer_contract(path: Path = ROOT / "packaging" / "installer.iss") -> None:
    text = path.read_text(encoding="utf-8")
    section = text.split("[Files]", 1)[1].split("\n[", 1)[0]
    sources = re.findall(r'^Source:\s*"([^"]+)"', section, re.MULTILINE)
    if len(sources) != len(INSTALLER_SOURCES) or set(sources) != set(INSTALLER_SOURCES):
        raise RuntimeError("installer.iss 的便携版文件白名单不符合约定")
    runtime = next(line for line in section.splitlines() if 'Source: "{#MySrcDir}\\_internal\\*"' in line)
    if "recursesubdirs" not in runtime or "createallsubdirs" not in runtime:
        raise RuntimeError("installer.iss 未递归保留 _internal 运行时")
    match = re.search(r"Excludes:\s*\"([^\"]+)\"", runtime)
    excludes = set(match.group(1).split(",")) if match else set()
    missing = set(INSTALLER_EXCLUDES) - excludes
    if missing:
        raise RuntimeError(f"installer.iss 缺少排除规则：{', '.join(sorted(missing))}")


def validate_portable_release(version: str) -> Path:
    built = DIST / f"审计工具箱_v{version}"
    validate_tree(built)
    unexpected = {path.name for path in built.iterdir()} - PORTABLE_TOP_LEVEL
    if unexpected:
        raise RuntimeError(f"便携版根目录含白名单外内容：{', '.join(sorted(unexpected))}")
    for name in PORTABLE_TOP_LEVEL - {"_internal"}:
        if not (built / name).is_file():
            raise RuntimeError(f"便携版缺少文件：{name}")
    if not (built / "_internal").is_dir():
        raise RuntimeError("便携版缺少 _internal 运行时")
    return built


def validate_archive(archive: Path, expected_files: set[str] | None = None) -> set[str]:
    with zipfile.ZipFile(archive) as zf:
        names: set[str] = set()
        for info in zf.infolist():
            name = info.filename
            if info.is_dir() or "\\" in name:
                raise RuntimeError(f"插件 ZIP 含目录或非标准路径：{name}")
            relative = PurePosixPath(name)
            validate_relative_path(relative)
            if stat.S_ISLNK(info.external_attr >> 16):
                raise RuntimeError(f"插件 ZIP 含符号链接：{name}")
            if not (name in PLUGIN_SOURCE_FILES or (len(relative.parts) > 1 and relative.parts[0] == "engine")):
                raise RuntimeError(f"插件 ZIP 含白名单外内容：{name}")
            if name in names:
                raise RuntimeError(f"插件 ZIP 含重复条目：{name}")
            names.add(name)
        missing_sources = set(PLUGIN_SOURCE_FILES) - names
        if missing_sources:
            raise RuntimeError(f"插件 ZIP 缺少源文件：{', '.join(sorted(missing_sources))}")
        if expected_files is not None and names != expected_files:
            missing = expected_files - names
            extra = names - expected_files
            raise RuntimeError(f"插件 ZIP 内容不一致：缺少 {sorted(missing)}，多出 {sorted(extra)}")
        corrupted = zf.testzip()
        if corrupted:
            raise RuntimeError(f"插件 ZIP 校验失败：{corrupted}")
    return names


def validate_source() -> None:
    validate_tree(PLUGIN_ROOT, set(PLUGIN_SOURCE_FILES))
    run(["node", str(PLUGIN_ROOT / "scripts" / "check.mjs"), str(PLUGIN_ROOT)])


def stage_release_files(version: str) -> Path:
    built = DIST / f"审计工具箱_v{version}"
    if not built.is_dir():
        raise RuntimeError(f"便携版目录不存在：{built}")
    shutil.copy2(ROOT / "packaging" / "使用说明.txt", built / "使用说明.txt")
    return validate_portable_release(version)


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
    source_files = validate_tree(PLUGIN_ROOT, set(PLUGIN_SOURCE_FILES))
    engine_files = validate_tree(engine_built)

    stage = TMP / "qoder_plugin_stage" / version
    if stage.exists():
        shutil.rmtree(stage)
    stage.mkdir(parents=True)
    expected_files = set(PLUGIN_SOURCE_FILES)
    try:
        for source in source_files:
            target = stage / source.relative_to(PLUGIN_ROOT)
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, target)
        for source in engine_files:
            relative = source.relative_to(engine_built)
            expected_files.add(f"engine/{relative.as_posix()}")
            target = stage / "engine" / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, target)

        validate_tree(stage, expected_files)
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
        validate_archive(archive, expected_files)
    finally:
        shutil.rmtree(stage, ignore_errors=True)
    print(f"Qoder 插件已生成：{archive}")
    return archive


def build_installer() -> Path:
    compiler = ROOT / ".tools" / "Inno Setup 6" / "ISCC.exe"
    if not compiler.is_file():
        raise RuntimeError(f"未找到 Inno Setup 编译器：{compiler}")
    version = re.search(r'^#define MyAppVersion "([^"]+)"', (ROOT / "packaging" / "installer.iss").read_text(encoding="utf-8"), re.MULTILINE)
    if not version:
        raise RuntimeError("无法读取安装程序版本")
    validate_portable_release(version.group(1))
    run([str(compiler), str(ROOT / "packaging" / "installer.iss")])
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
    validate_installer_contract()
    if args.check_only:
        portable = DIST / f"审计工具箱_v{version}"
        if portable.exists():
            validate_portable_release(version)
            print(f"便携版目录检查通过：{portable}")
        print("构建静态检查通过")
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
