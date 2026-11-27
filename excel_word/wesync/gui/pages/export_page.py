"""导出页：Word → Excel。"""

from __future__ import annotations

from PySide6.QtWidgets import QPushButton, QWidget

from ... import api
from ..widgets.cards import make_checkbutton
from ._page import BasePage


class ExportPage(BasePage):
    title = "创建Excel"
    subtitle = "Word 全部表格 → 工作簿（正文含 {字段名} 时自动生成「正文信息」表，重导出保留已填值；默认覆盖前自动备份，可取消勾选）。"

    def build(self, body: QWidget) -> None:
        card = self.card(body)
        self.links = make_checkbutton(card, "在 Excel 写入「→定位Word表」跳转链接", True)
        card.layout().addWidget(self.links)
        self.totals = make_checkbutton(card, "合计行自动写求和公式（在 Excel 改明细后合计自动更新，同步时按公式结果回写）", True)
        card.layout().addWidget(self.totals)
        self.backup = make_checkbutton(card, "覆盖前备份原文件到 _backup（取消勾选则直接覆盖、不可回滚）", True)
        card.layout().addWidget(self.backup)

        row = self.button_row(card, lambda: self._start(True), lambda: self._start(False), "仅预演", "开始导出")
        incremental_btn = QPushButton("增量导出")
        incremental_btn.setToolTip("只登记 Word 中新增的表格、标注与占位符；既有工作簿行和公式位置不变")
        incremental_btn.clicked.connect(lambda _checked=False: self._incremental())
        row.layout().insertWidget(row.layout().count() - 1, incremental_btn)
        rejudge_btn = QPushButton("重新判定✓/✕")
        rejudge_btn.setToolTip("在 Excel 改过内容后，按当前内容重判 B 列表头/✓/✕（只改标记，不改数据；「已删」行保持原样）")
        rejudge_btn.clicked.connect(lambda _checked=False: self._rejudge())
        row.layout().insertWidget(row.layout().count() - 1, rejudge_btn)
        btn = QPushButton("清理 Word 书签")
        btn.clicked.connect(lambda _checked=False: self._clear_bookmarks())
        row.layout().insertWidget(row.layout().count() - 1, btn)

    def _incremental(self) -> None:
        if not self.ctx.require_docx():
            return
        excel = self.ctx.workbook_path()
        if not excel.exists():
            self.ctx.error("未找到工作簿", f"未找到工作簿：\n{excel}\n\n请先执行一次全量「开始导出」。")
            return
        if not self.ctx.confirm(
            "确认增量导出",
            f"只登记 Word 中尚未锚定的新增表格、标注与占位符：\n\nWord：{self.ctx.docx}\nExcel：{excel}\n\n"
            "既有工作簿行、编号与公式位置不会改动；新表格按发现顺序写入「新增表格」（XT01、XT02…）。\n"
            "写盘前会自动备份，确认开始？",
        ):
            self.ctx.log("已取消增量导出。", "muted")
            return
        self.ctx.run_steps(
            [("增量导出", lambda: api.run_incremental_export(
                self.ctx.docx, self.ctx.excel,
                with_excel_links=self.links.isChecked(),
                totals_formula=self.totals.isChecked(),
                backup=self.backup.isChecked(),
            ))]
        )

    def _rejudge(self) -> None:
        if not self.ctx.require_docx():
            return
        backup = self.backup.isChecked()
        self.ctx.run_steps(
            [("重新判定✓/✕", lambda: api.run_rejudge(self.ctx.docx, self.ctx.excel, backup=backup))]
        )

    def _clear_bookmarks(self) -> None:
        if not self.ctx.require_docx():
            return
        if not self.ctx.confirm(
            "清理 Word 书签",
            f"将删除 Word 里所有本工具写入的书签（CPAH_/TBL_/MRK_）：\n\nWord：{self.ctx.docx}\n\n"
            "保留 Word 自带书签（如 _Toc 目录书签）。写盘前会自动备份，确认清理？",
        ):
            self.ctx.log("已取消清理书签。", "muted")
            return
        self.ctx.run_steps([("清理 Word 书签", lambda: api.run_clear_bookmarks(self.ctx.docx))])

    def _start(self, dry_run: bool) -> None:
        if not self.ctx.require_docx():
            return
        excel = self.ctx.workbook_path()
        with_links = self.links.isChecked()
        totals_formula = self.totals.isChecked()
        backup = self.backup.isChecked()
        if not dry_run:
            backup_note = (
                "原工作簿（若存在）会先备份到 _backup，确认开始导出？"
                if backup
                else "已取消备份：原工作簿（若存在）将被直接覆盖、不可回滚，确认开始导出？"
            )
            message = (
                f"将按当前 Word 重建工作簿：\n\nWord：{self.ctx.docx}\nExcel：{excel}\n\n"
                + backup_note
            )
            if not self.ctx.confirm("确认导出", message):
                self.ctx.log("已取消导出。", "muted")
                return
        self.ctx.run_steps(
            [
                (
                    "导出到 Excel（预演）" if dry_run else "导出到 Excel",
                    lambda: api.run_export(
                        self.ctx.docx,
                        self.ctx.excel,
                        dry_run=dry_run,
                        with_excel_links=with_links,
                        totals_formula=totals_formula,
                        backup=backup,
                    ),
                )
            ]
        )