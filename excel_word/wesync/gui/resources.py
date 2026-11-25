"""界面资源路径：同时兼容源码运行与 PyInstaller 打包。"""

from __future__ import annotations

import sys
from pathlib import Path


def _resource_path(name: str) -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / "wesync" / "resources" / name
    return Path(__file__).resolve().parents[1] / "resources" / name


def icon_path() -> Path:
    return _resource_path("audit_toolbox.png")


def manual_path() -> Path:
    if getattr(sys, "frozen", False):
        return _resource_path("使用说明.txt")
    return Path(__file__).resolve().parents[3] / "packaging" / "使用说明.txt"
