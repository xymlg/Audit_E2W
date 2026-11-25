"""卡片组件：白底细边框卡片与复选框。"""

from __future__ import annotations

from PySide6.QtWidgets import QCheckBox, QFrame, QVBoxLayout


def make_card(parent=None, padding: int = 16) -> QFrame:
    card = QFrame(parent)
    card.setObjectName("Card")
    layout = QVBoxLayout(card)
    layout.setContentsMargins(padding, padding, padding, padding)
    layout.setSpacing(8)
    return card


def make_checkbutton(parent, text: str, checked: bool = False) -> QCheckBox:
    box = QCheckBox(text, parent)
    box.setChecked(checked)
    return box