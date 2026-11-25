"""界面偏好持久化：窗口尺寸与上次使用的文件（写在项目内，不进入系统目录）。"""

from __future__ import annotations

import json
from pathlib import Path

from ..fsops import project_root


def _config_path() -> Path:
    return project_root() / ".wesync_ui.json"


def load() -> dict:
    path = _config_path()
    try:
        with path.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save(values: dict) -> None:
    path = _config_path()
    data = load()
    data.update(values)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
    except Exception:
        pass
