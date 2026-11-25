"""备份与回滚页：Word 与 Excel 的历史备份分开两栏列出，选中一键恢复（恢复前自动备份现状）。"""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QWidget,
)

from ... import api
from ...console import set_log_sink
from ...errors import SyncError
from ._page import BasePage


class HistoryPage(BasePage):
    title = "备份与回滚"
    subtitle = "浏览每次写盘前自动留下的历史版本，按时间点恢复；Word 与 Excel 分开列，可各自独立回滚。"

    def build(self, body: QWidget) -> None:
        card = self.card(body)
        self.tip_label(
            card,
            "每次同步/回写写盘前，工具会把原文件备份到同目录 _backup（带时间戳）。这里列出这些历史版本，"
            "选中某一行点「恢复」即可回滚到该时间点。恢复前会先把当前现状再备份一次，因此回滚本身也可逆。"
            "恢复前请先关闭对应的 Word / Excel。",
        )
        row = QHBoxLayout()
        refresh = QPushButton("刷新备份清单")
        refresh.clicked.connect(lambda _checked=False: self._refresh())
        row.addWidget(refresh)
        row.addStretch(1)
        card.layout().addLayout(row)
        self.status = self.tip_label(card, "尚未加载。选择 Word 文件后点「刷新备份清单」。")

        self.word_label = self.tip_label(card, "Word 报告备份")
        self.word_table = self._make_table()
        card.layout().addWidget(self.word_table)
        self.word_restore = QPushButton("恢复选中的 Word 版本")
        self.word_restore.setObjectName("Primary")
        self.word_restore.clicked.connect(lambda _checked=False: self._restore("word"))
        card.layout().addWidget(self.word_restore)

        self.excel_label = self.tip_label(card, "Excel 工作簿备份")
        self.excel_table = self._make_table()
        card.layout().addWidget(self.excel_table)
        self.excel_restore = QPushButton("恢复选中的 Excel 版本")
        self.excel_restore.setObjectName("Primary")
        self.excel_restore.clicked.connect(lambda _checked=False: self._restore("excel"))
        card.layout().addWidget(self.excel_restore)

        self._info: dict = {}
        self._set_restore_enabled()

    @staticmethod
    def _make_table() -> QTableWidget:
        table = QTableWidget(0, 3)
        table.setHorizontalHeaderLabels(["备份时间", "文件名", "大小"])
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.SingleSelection)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setMinimumHeight(140)
        return table

    def on_show(self) -> None:
        if not self._info:
            self._refresh(quiet=True)

    def _set_restore_enabled(self) -> None:
        self.word_restore.setEnabled(bool(self._info.get("word")))
        self.excel_restore.setEnabled(bool(self._info.get("excel")))

    def _refresh(self, quiet: bool = False) -> None:
        self._info = {}
        self.word_table.setRowCount(0)
        self.excel_table.setRowCount(0)
        if self.ctx.docx is None:
            self.status.setText("尚未加载。请先在上方选择 Word 文件。")
            self._set_restore_enabled()
            return
        set_log_sink(lambda message: self.ctx.log(message))
        try:
            self._info = api.scan_backups(self.ctx.docx, self.ctx.excel)
        except SyncError as exc:
            self.status.setText(f"加载失败：{exc}")
            self._set_restore_enabled()
            return
        finally:
            set_log_sink(None)
        self._fill(self.word_table, self._info.get("word", []))
        self._fill(self.excel_table, self._info.get("excel", []))
        nw, ne = len(self._info.get("word", [])), len(self._info.get("excel", []))
        self.word_label.setText(f"Word 报告备份（{nw} 个）：{self._info.get('word_path', '')}")
        self.excel_label.setText(f"Excel 工作簿备份（{ne} 个）：{self._info.get('excel_path', '')}")
        self.status.setText(f"Word 历史版本 {nw} 个、Excel 历史版本 {ne} 个。选中一行后点「恢复」。")
        self._set_restore_enabled()
        if not quiet:
            self.ctx.log("备份清单已刷新。", "muted")

    @staticmethod
    def _fill(table: QTableWidget, items: list[dict]) -> None:
        table.setRowCount(0)
        for item in items:
            row = table.rowCount()
            table.insertRow(row)
            size_kb = f"{item['size'] / 1024:.1f} KB"
            for c, value in enumerate([item["time"], item["file"], size_kb]):
                cell = QTableWidgetItem(str(value))
                if c == 0:
                    cell.setData(Qt.UserRole, item["path"])
                table.setItem(row, c, cell)

    def _restore(self, which: str) -> None:
        table = self.word_table if which == "word" else self.excel_table
        target = self._info.get("word_path" if which == "word" else "excel_path", "")
        rows = table.selectionModel().selectedRows()
        if not rows:
            self.ctx.info("未选择版本", "请先在列表里选中要恢复的历史版本。")
            return
        row = rows[0].row()
        backup = table.item(row, 0).data(Qt.UserRole)
        time_text = table.item(row, 0).text()
        label = "Word 报告" if which == "word" else "Excel 工作簿"
        if not self.ctx.confirm(
            "确认回滚",
            f"将把{label}恢复到 {time_text} 的版本：\n\n目标：{target}\n备份：{backup}\n\n"
            "恢复前会先把当前现状再备份一次（可再次回滚）。请先关闭该文件。确认恢复？",
        ):
            return
        set_log_sink(lambda message: self.ctx.log(message))
        try:
            api.run_restore_backup(backup, target)
        except SyncError as exc:
            self.ctx.error("回滚失败", str(exc))
            return
        finally:
            set_log_sink(None)
        self._refresh(quiet=True)