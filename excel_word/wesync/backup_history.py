"""备份历史与回滚：列出 _backup 里某文件的历史版本，按时间点恢复（恢复前先备份现状）。

备份命名沿用 fsops.backup_file：<主名>.<时间戳>.bak<扩展名>（重名追加 _1、_2）。
恢复是可逆的：先把当前文件再备份一次，再用选中的历史版本覆盖。
"""

from __future__ import annotations

import datetime as _dt
import os
import shutil
import tempfile
from pathlib import Path

from .console import log
from .errors import SyncError
from .fsops import backup_file, check_file_free


def _parse_stamp(mid: str) -> _dt.datetime | None:
    try:
        return _dt.datetime.strptime(mid[:15], "%Y%m%d_%H%M%S")
    except ValueError:
        return None


def list_backups(path: str | Path) -> list[dict]:
    """列出某文件（docx/xlsx）在 _backup 里的历史版本，按时间倒序。"""
    path = Path(path)
    backup_dir = path.parent / "_backup"
    if not backup_dir.is_dir():
        return []
    prefix = f"{path.stem}."
    suffix = f".bak{path.suffix}"
    items = []
    for f in backup_dir.iterdir():
        if not f.is_file() or not f.name.startswith(prefix) or not f.name.endswith(suffix):
            continue
        mid = f.name[len(prefix):-len(suffix)]
        stamp = _parse_stamp(mid)
        if stamp is None:
            continue
        sequence = int(mid[16:]) if len(mid) > 15 and mid[15] == "_" and mid[16:].isdigit() else 0
        items.append({
            "file": f.name,
            "path": str(f),
            "time": stamp.strftime("%Y-%m-%d %H:%M:%S"),
            "stamp": stamp,
            "sequence": sequence,
            "size": f.stat().st_size,
        })
    items.sort(key=lambda x: (x["stamp"], x["sequence"]), reverse=True)
    return items


def restore_backup(backup: str | Path, target: str | Path) -> Path | None:
    """用历史版本覆盖目标文件；恢复前先把现状备份一次，返回该安全备份路径。"""
    backup = Path(backup)
    target = Path(target)
    if not backup.is_file():
        raise SyncError(f"备份文件不存在：{backup}")
    if not target.exists():
        raise SyncError(f"要恢复的文件不存在：{target}\n（可能已被移动或改名）")
    if backup.suffix.lower() != target.suffix.lower():
        raise SyncError("备份文件与目标文件类型必须一致")
    if backup.resolve() not in {Path(item["path"]).resolve() for item in list_backups(target)}:
        raise SyncError(f"该文件不是 {target.name} 的历史备份；请从备份列表中选择。")
    check_file_free(target, False)
    safety = backup_file(target)
    fd, staged_name = tempfile.mkstemp(prefix=f".{target.stem}.restore.", suffix=target.suffix, dir=target.parent)
    os.close(fd)
    staged = Path(staged_name)
    try:
        shutil.copy2(backup, staged)
        check_file_free(target, False)
        os.replace(staged, target)
    finally:
        staged.unlink(missing_ok=True)
    log(f"[回滚] 已用 {backup.name} 恢复 {target.name}"
        + (f"（恢复前的现状已备份：{safety.name}，可再次回滚）" if safety else ""))
    return safety