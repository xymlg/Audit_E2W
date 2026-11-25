"""页面基类与共用构件：标题、说明、卡片、按钮行。"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from ..widgets.cards import make_card


class BasePage(QWidget):
    """所有功能页的基类：上方标题与说明，下方是可滚动的页面内容。"""

    title = "页面"
    subtitle = ""

    def __init__(self, ctx, parent=None):
        super().__init__(parent)
        self.ctx = ctx
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 12, 18, 12)
        layout.setSpacing(4)
        title = QLabel(self.title)
        title.setObjectName("PageTitle")
        layout.addWidget(title)
        if self.subtitle:
            desc = QLabel(self.subtitle)
            desc.setObjectName("Desc")
            desc.setWordWrap(True)
            desc.setMinimumWidth(0)
            layout.addWidget(desc)

        scroller = QScrollArea()
        scroller.setWidgetResizable(True)
        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(0, 8, 0, 0)
        body_layout.setSpacing(12)
        body_layout.addStretch(1)
        scroller.setWidget(body)
        layout.addWidget(scroller, 1)
        self.body = body
        self.build(body)

    def build(self, body: QWidget) -> None:  # pragma: no cover - 子类实现
        raise NotImplementedError

    def on_show(self) -> None:
        """切换到本页时调用，默认什么都不做。"""

    # ---------- 共用构件 ----------

    @staticmethod
    def card(body: QWidget) -> QWidget:
        card = make_card()
        # 插到末尾 stretch 之前，卡片保持从上往下排
        body.layout().insertWidget(body.layout().count() - 1, card)
        return card

    @staticmethod
    def tip_label(card: QWidget, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("CardDesc")
        label.setWordWrap(True)
        label.setMinimumWidth(0)
        card.layout().addWidget(label)
        return label

    def button_row(self, card: QWidget, on_dry, on_run, dry_text="仅预演", run_text="开始执行") -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        dry = QPushButton(dry_text)
        dry.clicked.connect(lambda _checked=False: on_dry())
        layout.addWidget(dry)
        run = QPushButton(run_text)
        run.setObjectName("Primary")
        run.clicked.connect(lambda _checked=False: on_run())
        layout.addWidget(run)
        layout.addStretch(1)
        card.layout().addWidget(row)
        return row