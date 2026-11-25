# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import re

TOOLS_ROOT = Path(SPECPATH).resolve().parent
ROOT = TOOLS_ROOT / "excel_word"
VERSION = re.search(
    r'^__version__\s*=\s*"([^"]+)"',
    (ROOT / "wesync" / "__init__.py").read_text(encoding="utf-8"),
    re.MULTILINE,
).group(1)
ENGINE_ENTRY = ROOT / "qoder_engine.py"
APP_ICON = ROOT / "wesync" / "resources" / "audit_toolbox.ico"
APP_ICON_PNG = ROOT / "wesync" / "resources" / "audit_toolbox.png"
MANUAL_TEXT = TOOLS_ROOT / "packaging" / "使用说明.txt"

EXCLUDED_MODULES = [
    "numpy", "pandas", "PIL", "matplotlib", "pytest", "IPython", "jupyter",
    "setuptools", "pip", "wheel", "PyInstaller", "nuitka", "cx_Freeze",
    "unittest", "pydoc_data", "lib2to3", "tkinter.test", "test",
    "PySide6", "shiboken6",
]

analysis = Analysis(
    [str(ENGINE_ENTRY)],
    pathex=[str(ROOT), str(TOOLS_ROOT)],
    binaries=[],
    datas=[
        (str(APP_ICON_PNG), "wesync/resources"),
        (str(MANUAL_TEXT), "wesync/resources"),
    ],
    hiddenimports=["wesync.cli"],
    hookspath=[],
    excludes=EXCLUDED_MODULES,
    noarchive=False,
)
pyz = PYZ(analysis.pure)
exe = EXE(
    pyz,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name="审计工具箱引擎",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    icon=str(APP_ICON),
)
COLLECT(
    exe,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=False,
    name=f"审计工具箱_qoder_engine_v{VERSION}",
)
