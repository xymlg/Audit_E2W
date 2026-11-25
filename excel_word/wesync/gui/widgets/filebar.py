"""顶部文件条：选择 Word 文件、显示/选择工作簿、打开所在文件夹（长地址自动换行完整显示）。"""

from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtWidgets import QGridLayout, QLabel, QPushButton, QWidget


class FileBar(QWidget):
    def __init__(
        self,
        parent,
        on_pick_docx: Callable[[], None],
        on_pick_excel: Callable[[], None],
        on_reset_excel: Callable[[], None],
        on_open_folder: Callable[[], None],
    ):
        super().__init__(parent)
        self.setStyleSheet("FileBar { background: #FFFFFF; }")
        grid = QGridLayout(self)
        grid.setContentsMargins(16, 10, 16, 10)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        grid.setColumnStretch(1, 1)

        grid.addWidget(QLabel("Word 文件"), 0, 0)
        self.docx_box = self._make_path_box()
        grid.addWidget(self.docx_box, 0, 1)
        btn = QPushButton("选择…")
        btn.clicked.connect(lambda _checked=False: on_pick_docx())
        grid.addWidget(btn, 0, 2)
        btn = QPushButton("打开所在文件夹")
        btn.clicked.connect(lambda _checked=False: on_open_folder())
        grid.addWidget(btn, 0, 3)

        grid.addWidget(QLabel("Excel 工作簿"), 1, 0)
        self.excel_box = self._make_path_box()
        grid.addWidget(self.excel_box, 1, 1)
        btn = QPushButton("选择…")
        btn.clicked.connect(lambda _checked=False: on_pick_excel())
        grid.addWidget(btn, 1, 2)
        btn = QPushButton("用默认名")
        btn.clicked.connect(lambda _checked=False: on_reset_excel())
        grid.addWidget(btn, 1, 3)

    @staticmethod
    def _make_path_box() -> QLabel:
        box = QLabel("（未选择）")
        box.setObjectName("PathBox")
        box.setWordWrap(True)
        box.setMinimumWidth(0)
        return box

    def set_paths(self, docx: Path | None, excel: Path | None) -> None:
        self.docx_box.setText(str(docx) if docx else "（未选择）")
        self.excel_box.setText(str(excel) if excel else "（未选择）")