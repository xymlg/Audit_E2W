"""左侧导航：功能列表 + 当前项高亮。"""

from __future__ import annotations

from typing import Callable

from PySide6.QtWidgets import QLabel, QPushButton, QVBoxLayout, QWidget


class NavBar(QWidget):
    def __init__(self, parent, on_select: Callable[[str], None]):
        super().__init__(parent)
        self.setObjectName("Nav")
        self.setStyleSheet("QWidget#Nav { background: #FFFFFF; }")
        self._on_select = on_select
        self._buttons: dict[str, QPushButton] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 14, 10, 14)
        layout.setSpacing(6)
        title = QLabel("审计工具箱")
        title.setObjectName("NavTitle")
        layout.addWidget(title)
        self._body = QVBoxLayout()
        self._body.setSpacing(6)
        layout.addLayout(self._body)
        layout.addStretch(1)

    def add_item(self, key: str, text: str) -> None:
        button = QPushButton(text)
        button.setObjectName("Nav")
        button.setCheckable(True)
        button.clicked.connect(lambda _checked=False, k=key: self._on_select(k))
        self._body.addWidget(button)
        self._buttons[key] = button

    def select(self, key: str) -> None:
        for item_key, button in self._buttons.items():
            button.setChecked(item_key == key)