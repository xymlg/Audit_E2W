"""书签页：只处理书签与正文遗留文字，不动数据。"""

from __future__ import annotations

from PySide6.QtWidgets import QWidget

from ... import api
from ._page import BasePage


class BookmarkPage(BasePage):
    title = "书签工具"
    subtitle = "按内容重新对齐书签，不改动数据。"

    def build(self, body: QWidget) -> None:
        card = self.card(body)
        self.button_row(card, lambda: self._start(True), lambda: self._start(False), "仅预演", "开始处理")

    def _start(self, dry_run: bool) -> None:
        if not self.ctx.require_docx():
            return
        if not dry_run:
            message = (
                f"将处理 Word 书签与正文遗留文字：\n\nWord：{self.ctx.docx}\n\n"
                "若有改动，写盘前会自动备份。确认开始？"
            )
            if not self.ctx.confirm("确认处理书签", message):
                self.ctx.log("已取消书签处理。", "muted")
                return
        self.ctx.run_steps(
            [
                (
                    "处理书签（预演）" if dry_run else "处理书签",
                    lambda: api.run_bookmark(self.ctx.docx, self.ctx.excel, dry_run=dry_run),
                )
            ]
        )