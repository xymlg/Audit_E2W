"""中文日志输出：普通日志、逐格改动日志（超限折叠）、警告与用法提示。

默认把日志打印到控制台；设置日志接收器后（界面用）改为投递给接收器，
因此命令行行为不受影响。
"""

from typing import Callable

from .config import MAX_LOG_LINES


_log_state = {"shown": 0, "suppressed": 0}
_log_sink: Callable[[str], None] | None = None


def set_log_sink(sink: Callable[[str], None] | None) -> None:
    """设置日志接收器；传 None 恢复为打印到控制台。"""
    global _log_sink
    _log_sink = sink


def reset_log_state() -> None:
    """重置逐格日志的折叠计数（每个任务开始前调用）。"""
    _log_state["shown"] = 0
    _log_state["suppressed"] = 0


def _emit(message: str) -> None:
    if _log_sink is not None:
        _log_sink(message)
    else:
        print(message, flush=True)


def log(message: str = "") -> None:
    _emit(message)


def log_change(message: str) -> None:
    """逐格改动日志，超过上限后折叠，避免刷屏。"""
    if _log_state["shown"] < MAX_LOG_LINES:
        _emit(message)
        _log_state["shown"] += 1
    else:
        _log_state["suppressed"] += 1


def flush_log_summary() -> None:
    if _log_state["suppressed"]:
        log(f"...（另有 {_log_state['suppressed']} 条改动日志已省略）")


def warn(message: str) -> None:
    _emit(f"警告：{message}")


USAGE_HINT = """用法：word_excel_sync.py <子命令> "<Word文件.docx>" [选项]

子命令：
  export     把 Word 全部表格导出为 Excel（含书签、跳转链接），会重建工作簿
  rejudge    按工作簿当前内容重新判定 B 列表头/✓/✕（只改标记，不改数据）
  bookmark   只重建/迁移 Word 书签并清理正文遗留的〔→Excel〕文字，不动 Excel 数据
  writeback  把 Excel 的改动（含增行删行、增列删列）回写到 Word；
             --with-marks / --with-placeholders 可叠加同步标注与占位符
  marks      把 Word 高亮标记导出/刷新到工作簿「链接标注」表
  marks-sync 把「链接标注」表填写的内容同步回 Word
  placeholders      按 Word 的 {字段名} 补齐工作簿「正文信息」填写骨架
  placeholders-sync 把「正文信息」表填好的值写回 Word 占位符
  blank      删除空白表格（整表全空/仅表头有字），--pick 编号 或 --all 选择后才删
  clear-bookmarks   删除本工具写入的书签，保留 Word 自带书签

常用示例：
  .venv\\Scripts\\python.exe word_excel_sync.py bookmark "0.0.0审计报告.docx"
  .venv\\Scripts\\python.exe word_excel_sync.py writeback "0.0.0审计报告.docx" --dry-run
  .venv\\Scripts\\python.exe word_excel_sync.py export "0.0.0审计报告.docx"

各子命令的完整参数请执行：word_excel_sync.py <子命令> -h
"""
