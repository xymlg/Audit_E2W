# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：免安装 onedir 程序（1 个图形界面 + 多个命令行）。

重建命令（在工具箱根目录执行）：
    .venv/Scripts/python.exe -m PyInstaller --noconfirm packaging/wxsync.spec
出安装版单文件：再用 Inno Setup 6 编译 packaging/installer.iss（见该脚本头部注释）。

产物：dist/审计工具箱_v<版本>/  （含 GUI exe 与命令行 exe 及 _internal 运行时）
说明：报告 docx / 工作簿 xlsx / _backup / .venv / .git / __pycache__ 都不会进包。
"""

from pathlib import Path

TOOLS_ROOT = Path(SPECPATH).resolve().parent  # SPECPATH=packaging → 工具箱根（packaging 在根目录）
ROOT = TOOLS_ROOT / "excel_word"                     # 各功能命令行与 wesync 包所在
VERSION = "1.11.3"
DIST_NAME = f"审计工具箱_v{VERSION}"

# 界面入口（无控制台窗口），位于工具箱根目录
GUI_ENTRY = TOOLS_ROOT / "审计工具箱.pyw"
# Qoder 插件无窗口引擎（JSONL 协议，保留控制台），位于 excel_word。
ENGINE_ENTRY = ROOT / "qoder_engine.py"
APP_ICON = ROOT / "wesync" / "resources" / "audit_toolbox.ico"
APP_ICON_PNG = ROOT / "wesync" / "resources" / "audit_toolbox.png"
MANUAL_TEXT = TOOLS_ROOT / "packaging" / "使用说明.txt"
# 命令行入口（保留控制台以便看日志），位于 excel_word。
# 默认只打图形界面以减小体积（GUI 已含全部功能）；如需命令行 exe，把脚本名加回本列表：
#   "1-导出表格到Excel.py", "2-Excel同步到Word.py", "3-重建书签.py",
#   "4-删除空白表格.py", "5-导出标注.py", "6-同步标注.py"
CLI_ENTRIES: list[str] = []
# GUI 入口文件名（用于判断无控制台）
GUI_SCRIPT_NAME = GUI_ENTRY.name

# 明确排除用不到的重型库，缩小产物体积（openpyxl 对 numpy/PIL 的引用都是可选的）
EXCLUDED_MODULES = [
    "numpy", "pandas", "PIL", "matplotlib", "pytest", "IPython", "jupyter",
    "setuptools", "pip", "wheel", "PyInstaller", "nuitka", "cx_Freeze",
    "unittest", "pydoc_data", "lib2to3", "tkinter.test", "test",
]


def analyze(script) -> Analysis:  # noqa: F821 - Analysis 由 PyInstaller 注入
    return Analysis(
        [str(script)],
        pathex=[str(ROOT), str(TOOLS_ROOT)],
        binaries=[],
        datas=[
            (str(APP_ICON_PNG), "wesync/resources"),
            (str(MANUAL_TEXT), "wesync/resources"),
        ],
        hiddenimports=["wesync.gui", "wesync.cli"],
        hookspath=[],
        excludes=EXCLUDED_MODULES,
        noarchive=False,
    )


def build_exe(script) -> tuple:  # noqa: F821 - EXE/PYZ 由 PyInstaller 注入
    analysis = analyze(script)
    pyz = PYZ(analysis.pure)
    is_gui = Path(script).name == GUI_SCRIPT_NAME
    if is_gui:
        name = "审计工具箱"
    elif Path(script) == ENGINE_ENTRY:
        name = "审计工具箱引擎"
    else:
        name = Path(script).stem
    exe = EXE(
        pyz,
        analysis.scripts,
        [],
        exclude_binaries=True,
        name=name,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=not is_gui,
        icon=str(APP_ICON),
    )
    return exe, analysis


_targets = [build_exe(GUI_ENTRY), build_exe(ENGINE_ENTRY)] + [build_exe(ROOT / script) for script in CLI_ENTRIES]

COLLECT(  # noqa: F821 - COLLECT 由 PyInstaller 注入
    *[item for target in _targets for item in (target[0], target[1].binaries, target[1].datas)],
    strip=False,
    upx=False,
    name=DIST_NAME,
)
