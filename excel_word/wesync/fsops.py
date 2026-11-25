"""文件操作：路径解析、写前备份、占用检测、未指定文件时的自动识别。"""

import datetime as _dt
import os
import shutil
import sys
import tempfile
from pathlib import Path
from .console import log, warn
from .errors import SyncError


def project_root() -> Path:
    """项目根目录；打包 exe 时为 exe 所在目录（约定：本项目文件不落项目外）。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def tmp_root() -> Path:
    """项目内临时/工作目录（.tmp，已 gitignore），替代系统 Temp。"""
    root = project_root() / ".tmp"
    root.mkdir(parents=True, exist_ok=True)
    return root


def resolve_docx(path_text: str) -> Path:
    path = Path(path_text).expanduser().resolve()
    if not path.exists():
        raise SyncError(f"找不到 Word 文件：{path}")
    if path.suffix.lower() not in (".docx", ".docm"):
        raise SyncError(f"只支持 .docx/.docm 文件：{path}")
    return path


def resolve_xlsx(docx_path: Path, excel_text: str | None) -> Path:
    if excel_text:
        return Path(excel_text).expanduser().resolve()
    return docx_path.with_name(f"{docx_path.stem}_表格数据.xlsx")


def find_default_docx(base_dir: Path) -> Path:
    """在指定目录里找唯一的 .docx（忽略 Word 临时文件）。"""
    candidates = sorted(
        path for path in base_dir.glob("*.docx") if path.is_file() and not path.name.startswith("~$")
    )
    if not candidates:
        raise SyncError(f"未在当前目录找到 .docx 文件，请把报告放到：{base_dir}，或显式指定文件路径")
    if len(candidates) > 1:
        names = "\n  ".join(path.name for path in candidates)
        raise SyncError(f"当前目录有多个 .docx，请显式指定要处理哪一个：\n  {names}")
    return candidates[0]


def default_base_dir(script_dir: Path) -> Path:
    """入口的“工作目录”：当前目录里有 .docx 就用当前目录，否则用入口所在目录。

    这样既能“双击界面（入口目录=工作目录）”，也能“在报告目录里直接运行命令行 exe”。
    """
    try:
        cwd = Path.cwd()
    except OSError:
        return script_dir
    if cwd != script_dir:
        try:
            if any(p.is_file() and not p.name.startswith("~$") for p in cwd.glob("*.docx")):
                return cwd
        except OSError:
            return script_dir
    return script_dir


def fill_docx_argument(argv: list[str], base_dir: Path) -> list[str]:
    """未显式给出 Word 文件时，自动补上目录里唯一的 .docx。"""
    if argv and not argv[0].startswith("-"):
        return argv
    docx_path = find_default_docx(base_dir)
    log(f"未指定文件，已自动使用：{docx_path.name}")
    return [str(docx_path), *argv]


def check_file_free(path: Path, force: bool) -> None:
    """独占写测试；文件被 Word/Excel 打开时给出中文提示。"""
    if not path.exists():
        return
    lock_names = {f"~${path.name}", f"~${path.name[2:]}"}
    stale_lock = any((path.parent / name).exists() for name in lock_names)
    try:
        with open(path, "r+b"):
            pass
    except OSError:
        if force:
            warn(f"{path.name} 似乎正被其他程序占用，已按 --force 继续（可能失败）")
        else:
            raise SyncError(
                f"{path.name} 正被其他程序占用，请先关闭 Word/Excel 后重试（或加 --force 强制继续）"
            )
    else:
        if stale_lock:
            warn(f"{path.name} 存在残留的临时锁文件（文件本身未被占用），已忽略并继续")


def backup_file(path: Path) -> Path | None:
    """落盘前备份到同目录 _backup，文件名带时间戳。"""
    if not path.exists():
        return None
    backup_dir = path.parent / "_backup"
    backup_dir.mkdir(exist_ok=True)
    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    target = backup_dir / f"{path.stem}.{stamp}.bak{path.suffix}"
    counter = 1
    while target.exists():
        target = backup_dir / f"{path.stem}.{stamp}_{counter}.bak{path.suffix}"
        counter += 1
    shutil.copy2(path, target)
    return target


def load_workbook_for_save(path: Path, **kwargs):
    """读取待修改工作簿；.xlsm 保留 VBA，避免工具写回时丢宏。"""
    from openpyxl import load_workbook

    if str(path).lower().endswith(".xlsm"):
        kwargs["keep_vba"] = True
    return load_workbook(str(path), **kwargs)


def _stage_office_file(path: Path, save, validate) -> dict:
    fd, temp_name = tempfile.mkstemp(prefix=f".{path.stem}.", suffix=path.suffix, dir=path.parent)
    os.close(fd)
    staged = Path(temp_name)
    rollback = None
    try:
        save(staged)
        validate(staged)
        fd, rollback_name = tempfile.mkstemp(prefix=f".{path.stem}.rollback.", suffix=path.suffix, dir=path.parent)
        os.close(fd)
        rollback = Path(rollback_name)
        shutil.copy2(path, rollback)
        return {"path": path, "new": staged, "old": rollback, "backup": None, "keep_old": False}
    except Exception:
        staged.unlink(missing_ok=True)
        if rollback is not None:
            rollback.unlink(missing_ok=True)
        raise


def save_docx_xlsx_transaction(
    *,
    docx_path: Path | None = None,
    doc=None,
    xlsx_path: Path | None = None,
    wb=None,
    backup: bool = True,
) -> dict[Path, Path | None]:
    """先序列化并校验 Word/Excel，再成对替换；第二步失败时恢复第一步。"""
    import docx
    from openpyxl import load_workbook

    staged: list[dict] = []

    def validate_xlsx(path: Path) -> None:
        check = load_workbook(str(path), read_only=True, data_only=False)
        try:
            if not check.sheetnames:
                raise ValueError("工作簿没有工作表")
        finally:
            check.close()

    def validate_docx(path: Path) -> None:
        docx.Document(str(path))

    try:
        if xlsx_path is not None:
            staged.append(_stage_office_file(xlsx_path, lambda path: wb.save(str(path)), validate_xlsx))
        if docx_path is not None:
            staged.append(_stage_office_file(docx_path, lambda path: doc.save(str(path)), validate_docx))
        if backup:
            for item in staged:
                item["backup"] = backup_file(item["path"])

        replaced: list[dict] = []
        for item in staged:
            try:
                os.replace(item["new"], item["path"])
            except OSError as exc:
                rollback_errors = []
                for previous in reversed(replaced):
                    try:
                        os.replace(previous["old"], previous["path"])
                        previous["old"] = None
                    except OSError as rollback_exc:
                        previous["keep_old"] = True
                        rollback_errors.append(f"{previous['path'].name}（{rollback_exc}）")
                if rollback_errors:
                    kept = "、".join(str(previous["old"]) for previous in replaced if previous["keep_old"])
                    raise SyncError(
                        f"保存 {item['path'].name} 失败，且自动回滚失败：{'、'.join(rollback_errors)}。"
                        f"原文件副本保留在：{kept}"
                    ) from exc
                raise SyncError(f"保存 {item['path'].name} 失败，Word/Excel 已恢复为同步前状态：{exc}") from exc
            item["new"] = None
            replaced.append(item)
        return {item["path"]: item["backup"] for item in staged}
    except SyncError:
        raise
    except Exception as exc:
        raise SyncError(f"保存前校验或备份失败，Word/Excel 均未替换：{exc}") from exc
    finally:
        for item in staged:
            if item["new"] is not None:
                item["new"].unlink(missing_ok=True)
            if item["old"] is not None and not item["keep_old"]:
                item["old"].unlink(missing_ok=True)
