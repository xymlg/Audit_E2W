"""空白表格页：删除 / 恢复 两个 tab。扫描候选 → 逐表勾选确认 → 删除（可逆，Excel 保留表块可恢复）。"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from ... import api
from ...errors import SyncError
from ..widgets.cards import make_checkbutton
from ._page import BasePage


def _clear_layout(layout) -> None:
    while layout.count():
        item = layout.takeAt(0)
        widget = item.widget()
        if widget is not None:
            widget.hide()
            widget.deleteLater()


class BlankPage(BasePage):
    title = "空白表格"
    subtitle = "找出整表全空或只有表头的表格，逐表勾选确认后删除；默认 Excel 保留表块，可随时恢复。"

    def build(self, body: QWidget) -> None:
        card = self.card(body)
        tabs = QTabWidget()
        card.layout().addWidget(tabs)
        tabs.addTab(self._build_delete_tab(), "删除空白表格")
        tabs.addTab(self._build_restore_tab(), "恢复已删除表格")
        tabs.addTab(self._build_batch_tab(), "批量删除空表")

    # ---------- 删除 tab ----------

    def _build_delete_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        tip = QLabel(
            "「扫描」只读不改文件。空白口径：首列（行次/项目名称）不参与判断，其余列全空，"
            "或只有表头有字、数据区全空；数据区只有 —、0 等占位内容的不算空白。"
            "删除前自动备份到 _backup；默认「可逆删除」：Excel 表块保留并标记「已删除」，"
            "之后可在「恢复已删除表格」tab 里恢复。"
        )
        tip.setObjectName("CardDesc")
        tip.setWordWrap(True)
        tip.setMinimumWidth(0)
        layout.addWidget(tip)

        scan = QPushButton("扫描空白表格")
        scan.setObjectName("Primary")
        scan.clicked.connect(lambda _checked=False: self._scan())
        row = QHBoxLayout()
        row.addWidget(scan)
        row.addStretch(1)
        layout.addLayout(row)

        self.list_frame = QWidget()
        self.list_layout = QVBoxLayout(self.list_frame)
        self.list_layout.setContentsMargins(0, 0, 0, 0)
        self.list_layout.setSpacing(6)
        layout.addWidget(self.list_frame)

        self.reversible = make_checkbutton(
            tab, "可逆删除：Excel 表块保留并标记「已删除」，之后可恢复（需已生成工作簿）", True
        )
        layout.addWidget(self.reversible)

        bottom = QHBoxLayout()
        bottom.setSpacing(12)
        self.btn_all = QPushButton("全选")
        self.btn_all.clicked.connect(lambda _checked=False: self._toggle_all(True))
        bottom.addWidget(self.btn_all)
        self.btn_none = QPushButton("全不选")
        self.btn_none.clicked.connect(lambda _checked=False: self._toggle_all(False))
        bottom.addWidget(self.btn_none)
        self.btn_delete = QPushButton("删除选中表格")
        self.btn_delete.setObjectName("Primary")
        self.btn_delete.clicked.connect(lambda _checked=False: self._delete())
        bottom.addWidget(self.btn_delete)
        bottom.addStretch(1)
        layout.addLayout(bottom)
        layout.addStretch(1)

        self._rows: list[tuple[dict, QCheckBox, QCheckBox | None]] = []
        self._show_placeholder("尚未扫描。选择 Word 文件后点「扫描空白表格」。")
        self._set_list_buttons(False)
        return tab

    # ---------- 恢复 tab ----------

    def _build_restore_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        tip = QLabel(
            "这里列出已可逆删除、还能恢复的表（Excel 里标记「书签：已删除」的表块）。"
            "勾选后点「恢复选中表格」，会按记录的原位置在 Word 里重建表格和标题。"
        )
        tip.setObjectName("CardDesc")
        tip.setWordWrap(True)
        tip.setMinimumWidth(0)
        layout.addWidget(tip)

        row = QHBoxLayout()
        row.setSpacing(12)
        refresh = QPushButton("刷新恢复清单")
        refresh.clicked.connect(lambda _checked=False: self._refresh_deleted())
        row.addWidget(refresh)
        self.btn_restore = QPushButton("恢复选中表格")
        self.btn_restore.setObjectName("Primary")
        self.btn_restore.clicked.connect(lambda _checked=False: self._restore())
        row.addWidget(self.btn_restore)
        row.addStretch(1)
        layout.addLayout(row)

        self.restore_frame = QWidget()
        self.restore_layout = QVBoxLayout(self.restore_frame)
        self.restore_layout.setContentsMargins(0, 0, 0, 0)
        self.restore_layout.setSpacing(6)
        layout.addWidget(self.restore_frame)
        layout.addStretch(1)

        self._restore_rows: list[tuple[dict, QCheckBox]] = []
        self._show_restore_placeholder("尚未加载。点「刷新恢复清单」查看可恢复的表。")
        self.btn_restore.setEnabled(False)
        return tab

    # ---------- 列表 ----------

    def _show_placeholder(self, text: str) -> None:
        _clear_layout(self.list_layout)
        label = QLabel(text)
        label.setObjectName("CardDesc")
        label.setWordWrap(True)
        label.setMinimumWidth(0)
        self.list_layout.addWidget(label)

    def _set_list_buttons(self, enabled: bool) -> None:
        for btn in (self.btn_all, self.btn_none, self.btn_delete):
            btn.setEnabled(enabled)

    def refresh_button_state(self) -> None:
        """任务结束后由主窗口调用：按当前清单重算按钮可用性。"""
        self._set_list_buttons(bool(self._rows))
        self.btn_restore.setEnabled(bool(self._restore_rows))
        self.batch_btn_delete.setEnabled(bool(self._batch_rows))

    def _toggle_all(self, value: bool) -> None:
        for _, var, _ in self._rows:
            var.setChecked(value)

    def _scan(self) -> None:
        if not self.ctx.require_docx():
            return
        try:
            candidates = api.scan_blank_tables(self.ctx.docx)
        except SyncError as exc:
            self.ctx.error("扫描失败", str(exc))
            return
        self._rows = []
        self._refresh_deleted()
        if not candidates:
            self._show_placeholder("未发现空白表格（整表全空或仅表头有字），无需处理。")
            self._set_list_buttons(False)
            self.ctx.log(f"扫描完成：{self.ctx.docx.name} 未发现空白表格。", "ok")
            return
        _clear_layout(self.list_layout)
        for cand in candidates:
            var = make_checkbutton(
                self.list_frame,
                f"候选{cand['no']} · {cand['code']}（{cand['kind']}，{cand['nrows']}行×{cand['ncols']}列）",
            )
            self.list_layout.addWidget(var)
            cap_var = None
            if cand["caption_text"]:
                cap_var = make_checkbutton(self.list_frame, f"　　连表标题一起删：{cand['caption_text']}", True)
                self.list_layout.addWidget(cap_var)
            else:
                label = QLabel("　　未识别到表标题，将只删表格")
                label.setObjectName("CardDesc")
                self.list_layout.addWidget(label)
            self._rows.append((cand, var, cap_var))
        self._set_list_buttons(True)
        self.ctx.log(f"扫描完成：发现 {len(candidates)} 张候选空白表，请逐表勾选后删除。")

    # ---------- 删除 ----------

    def _delete(self) -> None:
        if not self.ctx.require_docx():
            return
        selected = [
            (cand, bool(cap_var.isChecked()) if cap_var is not None else False)
            for cand, var, cap_var in self._rows
            if var.isChecked()
        ]
        reversible = self.reversible.isChecked()
        if not selected:
            self.ctx.info("未勾选", "请先勾选要删除的空白表格。")
            return
        with_caption = sum(1 for cand, with_cap in selected if with_cap and cand["caption_text"])
        mode = (
            "可逆删除：Excel 表块保留并标记「已删除」，之后可在本页恢复。"
            if reversible
            else "普通删除：Excel 表块不保留，删除后需重新导出工作簿。"
        )
        message = (
            f"将删除 {len(selected)} 张空白表格"
            + (f"（其中 {with_caption} 张连表标题一起删）" if with_caption else "")
            + f"：\n\nWord：{self.ctx.docx}\n\n{mode}\n写盘前会自动备份到 _backup。确认删除？"
        )
        if not self.ctx.confirm("确认删除空白表格", message):
            self.ctx.log("已取消删除空白表格。", "muted")
            return
        pick = ",".join(str(cand["no"]) for cand, _ in selected)
        keep_for = ",".join(
            str(cand["no"]) for cand, with_cap in selected if not with_cap and cand["caption_text"]
        )
        docx_path = self.ctx.docx
        excel_path = self.ctx.excel
        started = self.ctx.run_steps(
            [
                (
                    "删除空白表格",
                    lambda: api.run_blank_tables(
                        docx_path,
                        pick=pick,
                        keep_caption_for=keep_for or None,
                        reversible=reversible,
                        excel=excel_path,
                    ),
                )
            ]
        )
        if started:
            # 清单编号在删除后会变化，清掉旧清单防止按过期编号再次删除
            self._rows = []
            self._show_placeholder("删除任务已提交。完成后请重新点「扫描空白表格」刷新清单。")
            self._set_list_buttons(False)

    # ---------- 恢复 ----------

    def _show_restore_placeholder(self, text: str) -> None:
        _clear_layout(self.restore_layout)
        label = QLabel(text)
        label.setObjectName("CardDesc")
        label.setWordWrap(True)
        label.setMinimumWidth(0)
        self.restore_layout.addWidget(label)

    def _refresh_deleted(self) -> None:
        if not self.ctx.require_docx():
            return
        try:
            blocks = api.list_deleted_blocks(self.ctx.docx, self.ctx.excel)
        except SyncError as exc:
            self.ctx.error("加载恢复清单失败", str(exc))
            return
        self._restore_rows = []
        if not blocks:
            self._show_restore_placeholder("没有可恢复的表（工作簿里没有「已删除」标记的表块）。")
            self.btn_restore.setEnabled(False)
            return
        _clear_layout(self.restore_layout)
        for block in blocks:
            var = QCheckBox()
            pos = block.get("restore_after") or "末尾"
            name = block["name"] or block["code"]
            var.setText(f"{block['code']} · {name}（{block['nrows']}行×{block['ncols']}列，原位置：{pos}之后）")
            self.restore_layout.addWidget(var)
            self._restore_rows.append((block, var))
        self.btn_restore.setEnabled(True)
        self.ctx.log(f"恢复清单：{len(blocks)} 张表可恢复。")

    def _restore(self) -> None:
        if not self.ctx.require_docx():
            return
        nos = [str(i + 1) for i, (_, var) in enumerate(self._restore_rows) if var.isChecked()]
        if not nos:
            self.ctx.info("未勾选", "请先勾选要恢复的表格。")
            return
        message = (
            f"将恢复 {len(nos)} 张表：按工作簿保留的内容在原位置重建 Word 表格和标题。\n\n"
            f"Word：{self.ctx.docx}\n\n写盘前会自动备份到 _backup。确认恢复？"
        )
        if not self.ctx.confirm("确认恢复表格", message):
            self.ctx.log("已取消恢复表格。", "muted")
            return
        pick = ",".join(nos)
        docx_path = self.ctx.docx
        excel_path = self.ctx.excel
        started = self.ctx.run_steps(
            [("恢复已删除表格", lambda: api.run_restore_tables(docx_path, excel_path, pick=pick))]
        )
        if started:
            self._restore_rows = []
            self._show_restore_placeholder("恢复任务已提交。完成后请点「刷新恢复清单」查看剩余可恢复的表。")
            self.btn_restore.setEnabled(False)

    # ---------- 批量删除 tab ----------

    def _build_batch_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        tip = QLabel(
            "多份报告一次扫描、按报告分组勾选删除，口径与「删除空白表格」tab 完全一致。"
            "默认可逆删除（需各报告已生成工作簿）；恢复请到「恢复已删除表格」tab 按报告分别恢复。"
        )
        tip.setObjectName("CardDesc")
        tip.setWordWrap(True)
        tip.setMinimumWidth(0)
        layout.addWidget(tip)

        row = QHBoxLayout()
        add = QPushButton("添加报告")
        add.clicked.connect(lambda _checked=False: self._batch_add())
        row.addWidget(add)
        remove = QPushButton("移除选中")
        remove.clicked.connect(lambda _checked=False: self._batch_remove())
        row.addWidget(remove)
        scan = QPushButton("扫描全部")
        scan.setObjectName("Primary")
        scan.clicked.connect(lambda _checked=False: self._batch_scan())
        row.addWidget(scan)
        row.addStretch(1)
        layout.addLayout(row)

        self.batch_files = QListWidget()
        self.batch_files.setMaximumHeight(90)
        layout.addWidget(self.batch_files)

        self.batch_frame = QWidget()
        self.batch_layout = QVBoxLayout(self.batch_frame)
        self.batch_layout.setContentsMargins(0, 0, 0, 0)
        self.batch_layout.setSpacing(6)
        layout.addWidget(self.batch_frame)

        self.batch_reversible = make_checkbutton(
            tab, "可逆删除：Excel 表块保留并标记「已删除」，之后可按报告恢复（需已生成工作簿）", True
        )
        layout.addWidget(self.batch_reversible)

        bottom = QHBoxLayout()
        self.batch_btn_delete = QPushButton("删除选中表格")
        self.batch_btn_delete.setObjectName("Primary")
        self.batch_btn_delete.clicked.connect(lambda _checked=False: self._batch_delete())
        bottom.addWidget(self.batch_btn_delete)
        bottom.addStretch(1)
        layout.addLayout(bottom)
        layout.addStretch(1)

        self._batch_rows: list[tuple[str, dict, QCheckBox, QCheckBox | None]] = []
        self._batch_placeholder("尚未扫描。点「添加报告」加入多份 Word，再点「扫描全部」。")
        self.batch_btn_delete.setEnabled(False)
        return tab

    def _batch_placeholder(self, text: str) -> None:
        _clear_layout(self.batch_layout)
        label = QLabel(text)
        label.setObjectName("CardDesc")
        label.setWordWrap(True)
        label.setMinimumWidth(0)
        self.batch_layout.addWidget(label)

    def _batch_add(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "选择要批量处理的 Word 报告", "", "Word 文档 (*.docx)")
        existing = {self.batch_files.item(i).text() for i in range(self.batch_files.count())}
        for p in paths:
            if p not in existing:
                self.batch_files.addItem(p)

    def _batch_remove(self) -> None:
        for item in self.batch_files.selectedItems():
            self.batch_files.takeItem(self.batch_files.row(item))

    def _batch_scan(self) -> None:
        paths = [self.batch_files.item(i).text() for i in range(self.batch_files.count())]
        if not paths:
            self.ctx.info("未添加报告", "请先点「添加报告」加入要处理的 Word 文件。")
            return
        results = api.scan_blank_batch(paths)
        self._batch_rows = []
        _clear_layout(self.batch_layout)
        total = 0
        for result in results:
            name = Path(result["docx"]).name
            header = QLabel(f"■ {name}")
            header.setObjectName("SectionTitle")
            self.batch_layout.addWidget(header)
            if result["error"]:
                label = QLabel(f"　　扫描失败：{result['error']}")
                label.setObjectName("CardDesc")
                label.setWordWrap(True)
                label.setMinimumWidth(0)
                self.batch_layout.addWidget(label)
                continue
            if not result["candidates"]:
                label = QLabel("　　未发现空白表格")
                label.setObjectName("CardDesc")
                self.batch_layout.addWidget(label)
                continue
            for cand in result["candidates"]:
                var = make_checkbutton(
                    self.batch_frame,
                    f"　候选{cand['no']} · {cand['code']}（{cand['kind']}，{cand['nrows']}行×{cand['ncols']}列）",
                )
                self.batch_layout.addWidget(var)
                cap_var = None
                if cand["caption_text"]:
                    cap_var = make_checkbutton(self.batch_frame, f"　　　连表标题一起删：{cand['caption_text']}", True)
                    self.batch_layout.addWidget(cap_var)
                self._batch_rows.append((result["docx"], cand, var, cap_var))
                total += 1
        self.batch_btn_delete.setEnabled(bool(self._batch_rows))
        self.ctx.log(f"批量扫描完成：{len(paths)} 份报告，共 {total} 张候选空白表。")
        if not self._batch_rows:
            self._batch_placeholder("所有报告均未发现空白表格（或扫描失败，见上方提示）。")

    def _batch_delete(self) -> None:
        selected = [(docx, cand, bool(cap.isChecked()) if cap is not None else False)
                    for docx, cand, var, cap in self._batch_rows if var.isChecked()]
        if not selected:
            self.ctx.info("未勾选", "请先勾选要删除的空白表格。")
            return
        reversible = self.batch_reversible.isChecked()
        by_docx: dict[str, list] = {}
        for docx, cand, with_cap in selected:
            by_docx.setdefault(docx, []).append((cand, with_cap))
        mode = "可逆删除（Excel 表块保留，可按报告恢复）" if reversible else "普通删除（删除后需重新导出工作簿）"
        message = (
            f"将在 {len(by_docx)} 份报告中删除 {len(selected)} 张空白表格：\n\n"
            + "\n".join(f"· {Path(d).name}：{len(v)} 张" for d, v in by_docx.items())
            + f"\n\n{mode}，写盘前自动备份到 _backup。确认删除？"
        )
        if not self.ctx.confirm("确认批量删除空白表格", message):
            return
        items = []
        for docx, pairs in by_docx.items():
            items.append({
                "docx": docx,
                "pick": ",".join(str(c["no"]) for c, _ in pairs),
                "keep_caption_for": ",".join(str(c["no"]) for c, w in pairs if not w and c["caption_text"]) or None,
                "reversible": reversible,
            })
        if self.ctx.run_steps([("批量删除空白表格", lambda: (api.run_blank_batch(items), 0)[1])]):
            self._batch_rows = []
            self._batch_placeholder("删除任务已提交。完成后请重新点「扫描全部」刷新清单。")
            self.batch_btn_delete.setEnabled(False)