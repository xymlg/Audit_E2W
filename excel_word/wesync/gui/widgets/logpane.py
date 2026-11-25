"""右侧日志区：分色显示、清空/复制/保存、自动滚动开关、隐藏按钮。"""

from __future__ import annotations

import html
from typing import Callable

from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

LEVEL_TAGS = {
    "info": ("#1F2937", False),
    "muted": ("#6B7280", False),
    "ok": ("#15803D", False),
    "warn": ("#B45309", False),
    "error": ("#B91C1C", False),
    "step": ("#1F4E79", True),
}


class LogPane(QWidget):
    def __init__(self, parent, on_hide: Callable[[], None] | None = None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        header = QHBoxLayout()
        header.setSpacing(6)
        title = QLabel("运行日志")
        title.setObjectName("SectionTitle")
        header.addWidget(title)
        header.addStretch(1)
        if on_hide is not None:
            btn = QPushButton("隐藏")
            btn.setObjectName("LogTool")
            btn.clicked.connect(lambda _checked=False: on_hide())
            header.addWidget(btn)
        layout.addLayout(header)

        tools = QHBoxLayout()
        tools.setSpacing(6)
        self.auto_scroll = QCheckBox("自动滚动")
        self.auto_scroll.setChecked(True)
        tools.addWidget(self.auto_scroll)
        tools.addStretch(1)
        for text, slot in (("复制全部", self.copy_all), ("保存…", self.save_to_file), ("清空", self.clear)):
            btn = QPushButton(text)
            btn.setObjectName("LogTool")
            btn.clicked.connect(lambda _checked=False, fn=slot: fn())
            tools.addWidget(btn)
        layout.addLayout(tools)

        self.text = QPlainTextEdit()
        self.text.setObjectName("LogText")
        self.text.setReadOnly(True)
        self.text.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        layout.addWidget(self.text, 1)

    def append(self, message: str, level: str = "info") -> None:
        color, bold = LEVEL_TAGS.get(level, LEVEL_TAGS["info"])
        body = html.escape(message).replace("\n", "<br>")
        if bold:
            body = f"<b>{body}</b>"
        self.text.appendHtml(f'<span style="color:{color};">{body}</span>')
        if self.auto_scroll.isChecked():
            bar = self.text.verticalScrollBar()
            bar.setValue(bar.maximum())

    def clear(self) -> None:
        self.text.clear()

    def content(self) -> str:
        return self.text.toPlainText()

    def copy_all(self) -> None:
        text = self.content()
        if not text.strip():
            return
        from PySide6.QtWidgets import QApplication

        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "复制成功", "日志已复制到剪贴板。")

    def save_to_file(self) -> None:
        text = self.content()
        if not text.strip():
            QMessageBox.information(self, "没有日志", "当前没有可保存的日志内容。")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "保存运行日志", "运行日志.txt", "文本文件 (*.txt);;所有文件 (*.*)"
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(text + "\n")
        except OSError as exc:
            QMessageBox.critical(self, "保存失败", f"写入文件失败：{exc}")
            return
        QMessageBox.information(self, "保存成功", f"日志已保存到：\n{path}")