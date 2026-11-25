"""内置使用说明书：章节导航与全文搜索。"""

from __future__ import annotations

from PySide6.QtCore import QTimer, Qt
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QSplitter,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from ... import __version__
from ..manual import load_manual_text, manual_html, parse_manual
from ._page import BasePage


class HelpPage(BasePage):
    title = "使用说明"
    subtitle = "完整操作指南已内置，无需联网；选择左侧章节，或搜索关键词快速定位。"

    def build(self, body: QWidget) -> None:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        search_row = QWidget()
        search_layout = QHBoxLayout(search_row)
        search_layout.setContentsMargins(0, 0, 0, 0)
        search_layout.setSpacing(8)
        self.search_edit = QLineEdit()
        self.search_edit.setObjectName("ManualSearch")
        self.search_edit.setPlaceholderText("搜索说明书内容，回车查找下一处")
        self.search_edit.returnPressed.connect(self._find_next)
        search_layout.addWidget(self.search_edit, 1)
        self.find_button = QPushButton("查找下一处")
        self.find_button.clicked.connect(self._find_next)
        search_layout.addWidget(self.find_button)
        self.search_status = QLabel(f"当前版本 v{__version__}")
        self.search_status.setObjectName("CardDesc")
        search_layout.addWidget(self.search_status)
        layout.addWidget(search_row)

        splitter = QSplitter(Qt.Horizontal)
        self.chapter_list = QListWidget()
        self.chapter_list.setObjectName("ManualChapters")
        self.chapter_list.setMaximumWidth(230)
        self.chapter_list.currentRowChanged.connect(self._show_chapter)
        self.browser = QTextBrowser()
        self.browser.setObjectName("ManualView")
        splitter.addWidget(self.chapter_list)
        splitter.addWidget(self.browser)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([230, 620])
        layout.addWidget(splitter, 1)

        body.layout().insertWidget(body.layout().count() - 1, container)

        try:
            self.sections = parse_manual(load_manual_text())
            self.browser.setHtml(manual_html(self.sections))
        except OSError as exc:
            self.sections = []
            self.browser.setPlainText(f"说明书资源加载失败：{exc}")
        for index, (title, _lines) in enumerate(self.sections):
            item = QListWidgetItem(title)
            item.setData(Qt.UserRole, f"chapter-{index}")
            self.chapter_list.addItem(item)
        if self.sections:
            self.chapter_list.setCurrentRow(0)

    def _show_chapter(self, row: int) -> None:
        if row < 0 or row >= len(self.sections):
            return
        anchor = f"chapter-{row}"
        QTimer.singleShot(0, lambda: self.browser.scrollToAnchor(anchor))

    def _find_next(self) -> None:
        query = self.search_edit.text().strip()
        if not query:
            self.search_status.setText("请输入要查找的关键词")
            return
        found = self.browser.find(query)
        self.search_status.setText("已定位到下一处" if found else "未找到，请换个关键词")
