"""标注与占位页：链接标注（高亮标记）与正文 {字段名} 占位符的导出/同步。"""

from __future__ import annotations

from PySide6.QtWidgets import QHBoxLayout, QPushButton, QTabWidget, QVBoxLayout, QWidget

from ... import api
from ..widgets.cards import make_checkbutton
from ._page import BasePage


class MarksPage(BasePage):
    title = "标注与占位"
    subtitle = "高亮标注与正文 {字段名} 占位符：导出到 Excel，填好后一键写回。"

    def build(self, body: QWidget) -> None:
        card = self.card(body)
        tabs = QTabWidget()
        card.layout().addWidget(tabs)
        tabs.addTab(self._build_marks_tab(), "链接标注")
        tabs.addTab(self._build_placeholder_tab(), "正文占位符")

    # ---------- 链接标注 tab ----------

    def _build_marks_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        # 默认勾选“同步后保留高亮”：便于同步后在 Word 里复核，需要自动去色时手动取消。
        self.keep_highlight = make_checkbutton(tab, "同步后保留高亮（不去色，便于复核）", True)
        layout.addWidget(self.keep_highlight)
        layout.addWidget(self._marks_buttons())
        layout.addStretch(1)
        return tab

    def _marks_buttons(self) -> QWidget:
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        btn = QPushButton("导出/刷新标注")
        btn.clicked.connect(lambda _checked=False: self._export_marks())
        layout.addWidget(btn)
        btn = QPushButton("仅预演同步")
        btn.clicked.connect(lambda _checked=False: self._sync_marks(True))
        layout.addWidget(btn)
        btn = QPushButton("开始同步")
        btn.setObjectName("Primary")
        btn.clicked.connect(lambda _checked=False: self._sync_marks(False))
        layout.addWidget(btn)
        layout.addStretch(1)
        return row

    def _export_marks(self) -> None:
        if not self.ctx.require_docx():
            return
        xlsx = self.ctx.workbook_path()
        if xlsx.exists():
            steps = [("导出/刷新标注", lambda: api.run_marks(self.ctx.docx, self.ctx.excel))]
        else:
            if not self.ctx.confirm(
                "尚未生成工作簿",
                f"未找到工作簿：\n{xlsx}\n\n是否现在先执行一次完整导出（表格 + 链接标注）？",
            ):
                self.ctx.log("已取消导出。", "muted")
                return
            steps = [("导出表格到 Excel（含链接标注）", lambda: api.run_export(self.ctx.docx, self.ctx.excel))]
        self.ctx.run_steps(steps)

    def _sync_marks(self, dry_run: bool) -> None:
        if not self.ctx.require_docx():
            return
        xlsx = self.ctx.workbook_path()
        keep_highlight = self.keep_highlight.isChecked()
        if not xlsx.exists():
            self.ctx.log(f"未找到工作簿：{xlsx}\n请先执行“导出/刷新标注”。", "error")
            return
        if not dry_run:
            message = (
                f"将把「链接标注」工作表里填写的内容写回 Word：\n\nWord：{self.ctx.docx}\nExcel：{xlsx}\n\n"
                + ("同步后保留高亮。\n" if keep_highlight else "同步后自动去掉该处高亮。\n")
                + "\n"
                "写盘前会自动备份，确认开始同步？"
            )
            if not self.ctx.confirm("确认同步标注", message):
                self.ctx.log("已取消标注同步。", "muted")
                return
        self.ctx.run_steps(
            [
                (
                    ("同步标注（预演）" if dry_run else "同步标注") + ("·保留高亮" if keep_highlight else ""),
                    lambda: api.run_marks_sync(
                        self.ctx.docx,
                        self.ctx.excel,
                        dry_run=dry_run,
                        keep_highlight=keep_highlight,
                    ),
                )
            ]
        )

    # ---------- 正文占位符 tab ----------

    def _build_placeholder_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        self.tip_label(
            tab,
            "扫描 Word 正文与表格里的 {字段名} 占位符，在工作簿「正文信息」表生成填写骨架"
            "（A 列字段名、B 列填值，已填的值不会被覆盖）。填好后同步回 Word；"
            "未填值的占位符保留原样并逐条警告。该表与「模板生成」共用。",
        )
        row = QWidget()
        hl = QHBoxLayout(row)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(12)
        btn = QPushButton("导出/刷新占位符")
        btn.clicked.connect(lambda _checked=False: self._export_placeholders())
        hl.addWidget(btn)
        btn = QPushButton("仅预演同步")
        btn.clicked.connect(lambda _checked=False: self._sync_placeholders(True))
        hl.addWidget(btn)
        btn = QPushButton("开始同步")
        btn.setObjectName("Primary")
        btn.clicked.connect(lambda _checked=False: self._sync_placeholders(False))
        hl.addWidget(btn)
        hl.addStretch(1)
        layout.addWidget(row)
        layout.addStretch(1)
        return tab

    def _export_placeholders(self) -> None:
        if not self.ctx.require_docx():
            return
        xlsx = self.ctx.workbook_path()
        if xlsx.exists():
            steps = [("导出/刷新占位符", lambda: api.run_placeholder_export(self.ctx.docx, self.ctx.excel))]
        else:
            if not self.ctx.confirm(
                "尚未生成工作簿",
                f"未找到工作簿：\n{xlsx}\n\n是否现在先执行一次完整导出（表格 + 链接标注），再补齐占位符骨架？",
            ):
                self.ctx.log("已取消导出。", "muted")
                return
            steps = [
                ("导出表格到 Excel（含链接标注）", lambda: api.run_export(self.ctx.docx, self.ctx.excel)),
                ("导出/刷新占位符", lambda: api.run_placeholder_export(self.ctx.docx, self.ctx.excel)),
            ]
        self.ctx.run_steps(steps)

    def _sync_placeholders(self, dry_run: bool) -> None:
        if not self.ctx.require_docx():
            return
        xlsx = self.ctx.workbook_path()
        if not xlsx.exists():
            self.ctx.log(f"未找到工作簿：{xlsx}\n请先执行“导出/刷新占位符”。", "error")
            return
        if not dry_run:
            message = (
                f"将按「正文信息」表填写的值替换 Word 里的 {{字段名}} 占位符：\n\n"
                f"Word：{self.ctx.docx}\nExcel：{xlsx}\n\n"
                "未填值的占位符保留原样并逐条警告。\n"
                "写盘前会自动备份，确认开始同步？"
            )
            if not self.ctx.confirm("确认同步占位符", message):
                self.ctx.log("已取消占位符同步。", "muted")
                return
        self.ctx.run_steps(
            [
                (
                    "同步占位符（预演）" if dry_run else "同步占位符",
                    lambda: api.run_placeholder_sync(self.ctx.docx, self.ctx.excel, dry_run=dry_run),
                )
            ]
        )