"""主窗口：左侧导航 + 顶部文件条 + 中间页面 + 右侧可隐藏日志 + 状态栏。"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QIcon, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from .. import api
from ..errors import SyncError
from ..fsops import find_default_docx
from . import store
from .pages import iter_pages
from .resources import icon_path
from .runner import TaskRunner
from .widgets.filebar import FileBar
from .widgets.logpane import LogPane
from .widgets.nav import NavBar

DEFAULT_SIZE = (1100, 640)
MIN_SIZE = (1080, 560)
DEFAULT_LOG_WIDTH = 320


class UiContext:
    """页面与主窗口之间的桥梁：文件状态、日志、弹窗、任务执行。"""

    def __init__(self, app: "WesyncApp") -> None:
        self._app = app

    # ---------- 文件 ----------
    @property
    def docx(self) -> Path | None:
        return self._app.docx_path

    @property
    def excel(self) -> Path | None:
        return self._app.excel_path

    def workbook_path(self) -> Path:
        if self._app.docx_path is None:
            return Path("")
        return api.workbook_path(self._app.docx_path, self._app.excel_path)

    def require_docx(self) -> bool:
        if self._app.docx_path is None:
            self.error("未选择文件", "请先在顶部选择要处理的 Word 文件。")
            return False
        if not Path(self._app.docx_path).exists():
            self.error("文件不存在", f"找不到 Word 文件：\n{self._app.docx_path}")
            return False
        return True

    def set_docx(self, path: Path) -> None:
        self._app.set_docx(Path(path))

    def set_excel(self, path: Path | None) -> None:
        self._app.set_excel(None if path is None else Path(path))

    # ---------- 日志与状态 ----------
    def log(self, text: str, level: str = "info") -> None:
        self._app.append_log(text, level)

    def set_status(self, text: str) -> None:
        self._app.set_status(text)

    # ---------- 弹窗 ----------
    def confirm(self, title: str, message: str) -> bool:
        box = QMessageBox(self._app)
        box.setIcon(QMessageBox.Question)
        box.setWindowTitle(title)
        box.setText(message)
        box.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        box.button(QMessageBox.Ok).setText("确定")
        box.button(QMessageBox.Cancel).setText("取消")
        return box.exec() == QMessageBox.Ok

    def info(self, title: str, message: str) -> None:
        QMessageBox.information(self._app, title, message)

    def error(self, title: str, message: str, log: bool = True) -> None:
        if log:
            self._app.append_log(f"错误：{message}", "error")
        QMessageBox.critical(self._app, title, message)

    def open_folder(self) -> None:
        self._app.open_folder()

    # ---------- 任务 ----------
    @property
    def busy(self) -> bool:
        return self._app.runner.busy

    def run_steps(self, steps) -> bool:
        if self._app.runner.busy:
            self.log("已有任务在执行，请等待完成后再操作。", "warn")
            return False
        started = self._app.runner.submit(steps)
        if started:
            self._app.set_busy(True)
        return started


class WesyncApp(QMainWindow):
    """界面主窗口。"""

    def __init__(self, base_dir: Path | None = None) -> None:
        super().__init__()
        self.base_dir = Path(base_dir) if base_dir else Path(__file__).resolve().parents[2]
        self.settings = store.load()
        self.docx_path: Path | None = None
        self.excel_path: Path | None = None
        self.excel_is_custom = False
        self.runner = TaskRunner()
        self.ctx = UiContext(self)
        self._pages: dict[str, QWidget] = {}

        self.setWindowTitle("审计工具箱")
        self.setWindowIcon(QIcon(str(icon_path())))
        size = self.settings.get("qt_size") or list(DEFAULT_SIZE)
        self.resize(int(size[0]), int(size[1]))
        pos = self.settings.get("qt_pos")
        if pos:
            self.move(int(pos[0]), int(pos[1]))
        self.setMinimumSize(*MIN_SIZE)

        self._build_layout()
        self._load_initial_paths()
        self._select_page(next(iter(self._pages)))
        self._help_shortcut = QShortcut(QKeySequence("F1"), self)
        self._help_shortcut.activated.connect(lambda: self._select_page("HelpPage"))
        QTimer(self, timeout=self._poll_queue, interval=120).start()

    # ------------------------------------------------------------------ 布局

    def _build_layout(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.filebar = FileBar(
            central,
            on_pick_docx=self.pick_docx,
            on_pick_excel=self.pick_excel,
            on_reset_excel=self.reset_excel,
            on_open_folder=self.open_folder,
        )
        layout.addWidget(self.filebar)

        self.splitter = QSplitter(Qt.Horizontal)
        self.nav = NavBar(self.splitter, self._select_page)
        self.stack = QStackedWidget()
        self.logpane = LogPane(self.splitter, on_hide=lambda: self.set_log_visible(False))
        self.splitter.addWidget(self.nav)
        self.splitter.addWidget(self.stack)
        self.splitter.addWidget(self.logpane)
        self.splitter.setStretchFactor(0, 0)
        self.splitter.setStretchFactor(1, 1)
        self.splitter.setStretchFactor(2, 0)
        self.splitter.setCollapsible(0, False)
        self.splitter.setCollapsible(2, False)
        log_width = int(self.settings.get("log_width") or DEFAULT_LOG_WIDTH)
        # 首次布局会按 sizeHint 重分宽度，等布局生效后再按设定值分配
        QTimer.singleShot(0, lambda: self._apply_splitter_sizes(log_width))
        layout.addWidget(self.splitter, 1)

        for page_class in iter_pages():
            page = page_class(self.ctx)
            key = page_class.__name__
            self._pages[key] = page
            self.stack.addWidget(page)
            self.nav.add_item(key, getattr(page, "title", key))

        self._status_label = QLabel("就绪")
        self._status_label.setObjectName("Status")
        self.statusBar().addWidget(self._status_label, 1)
        self._toggle_log_btn = QPushButton("隐藏日志")
        self._toggle_log_btn.clicked.connect(lambda _checked=False: self.set_log_visible(not self.logpane.isVisible()))
        self.statusBar().addPermanentWidget(self._toggle_log_btn)
        self._progress = QProgressBar()
        self._progress.setFixedWidth(150)
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        self._progress.setTextVisible(False)
        self.statusBar().addPermanentWidget(self._progress)
        self.set_log_visible(bool(self.settings.get("log_visible", True)))

    def _apply_splitter_sizes(self, log_width: int) -> None:
        total = sum(self.splitter.sizes()) or self.width()
        self.splitter.setSizes([168, max(total - 168 - log_width, 300), log_width])

    def set_log_visible(self, visible: bool) -> None:
        self.logpane.setVisible(visible)
        self._toggle_log_btn.setText("隐藏日志" if visible else "显示日志")
        if visible:
            self._apply_splitter_sizes(int(self.settings.get("log_width") or DEFAULT_LOG_WIDTH))

    # ------------------------------------------------------------ 文件与路径

    def _load_initial_paths(self) -> None:
        docx: Path | None = None
        remembered = self.settings.get("last_docx")
        if remembered and Path(remembered).exists():
            docx = Path(remembered)
        else:
            try:
                docx = find_default_docx(self.base_dir)
            except SyncError:
                docx = None
        remembered_excel = self.settings.get("last_excel")
        if remembered_excel and Path(remembered_excel).exists():
            self.excel_path = Path(remembered_excel)
            self.excel_is_custom = True
        self.set_docx(docx)

    def set_docx(self, path: Path | None) -> None:
        self.docx_path = path
        if path is not None and not self.excel_is_custom:
            self.excel_path = api.workbook_path(path, None)
        self.filebar.set_paths(self.docx_path, self.excel_path)
        if path is None:
            self.set_status("尚未选择 Word 文件")
        else:
            exists = self.excel_path.exists() if self.excel_path else False
            suffix = "" if exists else "（工作簿尚未生成）"
            self.set_status(f"当前文件：{path.name}{suffix}")

    def set_excel(self, path: Path | None) -> None:
        self.excel_path = path
        self.excel_is_custom = path is not None
        self.filebar.set_paths(self.docx_path, self.excel_path)
        self.set_status("已指定工作簿：" + (path.name if path else "（未指定）"))

    def pick_docx(self) -> None:
        initial = str(self.docx_path.parent if self.docx_path else self.base_dir)
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 Word 文件", initial, "Word 文档 (*.docx);;所有文件 (*.*)"
        )
        if path:
            self.set_docx(Path(path))
            self.append_log(f"已选择 Word 文件：{path}", "muted")

    def pick_excel(self) -> None:
        initial = self.excel_path or self.base_dir
        path, _ = QFileDialog.getSaveFileName(
            self,
            "选择或指定工作簿",
            str(initial.parent / (initial.name if self.excel_path else "表格数据.xlsx")),
            "Excel 工作簿 (*.xlsx);;所有文件 (*.*)",
        )
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.set_excel(Path(path))
            self.append_log(f"已指定工作簿：{path}", "muted")

    def reset_excel(self) -> None:
        self.excel_is_custom = False
        self.set_excel_default()

    def set_excel_default(self) -> None:
        if self.docx_path is None:
            self.excel_path = None
        else:
            self.excel_path = api.workbook_path(self.docx_path, None)
        self.filebar.set_paths(self.docx_path, self.excel_path)
        self.set_status("工作簿已恢复为默认命名：" + (self.excel_path.name if self.excel_path else "（无）"))

    def open_folder(self) -> None:
        target = self.docx_path.parent if self.docx_path else self.base_dir
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(target))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(target)])
            else:
                subprocess.Popen(["xdg-open", str(target)])
        except Exception as exc:
            self.ctx.error("无法打开目录", f"{target}\n{exc}")

    # -------------------------------------------------------------- 页面切换

    def _select_page(self, key: str) -> None:
        if key not in self._pages:
            return
        page = self._pages[key]
        self.stack.setCurrentWidget(page)
        self.nav.select(key)
        if hasattr(page, "on_show"):
            page.on_show()
        self.set_status(f"当前页面：{getattr(page, 'title', key)}")

    # -------------------------------------------------------- 日志与状态栏

    @staticmethod
    def _level_for(text: str) -> str:
        stripped = text.strip()
        if stripped.startswith("警告："):
            return "warn"
        if stripped.startswith("错误：") or stripped.startswith("Traceback"):
            return "error"
        if any(mark in stripped for mark in ("已保存", "已写入", "已更新", "已创建", "完成")):
            return "ok"
        return "info"

    def append_log(self, text: str, level: str | None = None) -> None:
        self.logpane.append(text, level if level else self._level_for(text))

    def set_status(self, text: str) -> None:
        self._status_label.setText(text)

    # ---------------------------------------------------------- 后台任务处理

    def set_busy(self, busy: bool) -> None:
        buttons = self.stack.findChildren(QPushButton) + self.filebar.findChildren(QPushButton)
        if busy:
            # 记住执行前就禁用的按钮，结束后不能一律启用
            self._disabled_before_busy = {btn for btn in buttons if not btn.isEnabled()}
            for widget in buttons:
                widget.setEnabled(False)
            self.set_status("执行中…请稍候")
            self._progress.setRange(0, 0)
        else:
            disabled = getattr(self, "_disabled_before_busy", set())
            for widget in buttons:
                widget.setEnabled(widget not in disabled)
            self._disabled_before_busy = set()
            # 任务期间页面可能清空了清单，按当前数据重算按钮可用性
            for page in self._pages.values():
                refresh = getattr(page, "refresh_button_state", None)
                if refresh is not None:
                    refresh()
            self._progress.setRange(0, 100)
            self._progress.setValue(0)

    def _poll_queue(self) -> None:
        for event, payload in self.runner.drain():
            if event == "log":
                self.append_log(str(payload))
            elif event == "step":
                self.append_log(f"▶ {payload} …", "step")
                self.set_status(f"执行中：{payload}")
            elif event == "done":
                data = payload if isinstance(payload, dict) else {}
                titles = "、".join(title for title, _ in data.get("steps", []))
                elapsed = float(data.get("elapsed", 0.0))
                self.append_log(f"✔ 完成：{titles}（耗时 {elapsed:.1f} 秒）", "ok")
                self.set_status(f"最近一次：{titles} · 耗时 {elapsed:.1f} 秒")
                self.set_busy(False)
                QTimer.singleShot(
                    200, lambda: self.ctx.info("执行完成", f"{titles}\n耗时：{elapsed:.1f} 秒\n\n详细过程见日志区。")
                )
            elif event == "error":
                data = payload if isinstance(payload, dict) else {}
                title = str(data.get("title", "任务"))
                message = str(data.get("message", "未知错误"))
                elapsed = float(data.get("elapsed", 0.0))
                self.append_log(f"✘ 「{title}」执行失败（耗时 {elapsed:.1f} 秒）：{message}", "error")
                self.set_status(f"失败：{title}")
                self.set_busy(False)
                QTimer.singleShot(200, lambda: self.ctx.error("执行失败", f"{title}\n\n{message}", log=False))

    # ------------------------------------------------------------------ 退出

    def closeEvent(self, event) -> None:
        sizes = self.splitter.sizes()
        store.save(
            {
                "qt_size": [self.width(), self.height()],
                "qt_pos": [self.x(), self.y()],
                "log_visible": self.logpane.isVisible(),
                "log_width": sizes[2] if sizes[2] >= 80 else int(self.settings.get("log_width") or DEFAULT_LOG_WIDTH),
                "last_docx": str(self.docx_path) if self.docx_path else "",
                "last_excel": str(self.excel_path) if (self.excel_path and self.excel_is_custom) else "",
            }
        )
        super().closeEvent(event)