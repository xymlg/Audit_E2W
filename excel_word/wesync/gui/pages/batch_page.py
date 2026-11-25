"""批量同步页：多报告批量同步 / 模板生成（单模板或多模板分组）两个 tab。"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QRadioButton,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ... import api
from ..widgets.cards import make_checkbutton
from ._page import BasePage


class BatchPage(BasePage):
    title = "批量同步"
    subtitle = "多份报告一次同步；或用一个/多个模板配多个 Excel，批量生成若干份新 Word（模板与源 Excel 不改动）。"

    def build(self, body: QWidget) -> None:
        card = self.card(body)
        tabs = QTabWidget()
        card.layout().addWidget(tabs)
        tabs.addTab(self._build_sync_tab(), "多报告批量同步")
        tabs.addTab(self._build_template_tab(), "模板生成")
        self._waiting: str | None = None
        self._results: list[dict] = []
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._check_done)
        self._timer.start(300)

    def _check_done(self) -> None:
        if not self._waiting or self.ctx.busy:
            return
        mode, self._waiting = self._waiting, None
        if mode == "sync":
            by_path = {r["docx"]: r for r in self._results}
            for row in range(self.sync_table.rowCount()):
                path = self.sync_table.item(row, 0).data(Qt.UserRole)
                r = by_path.get(path)
                if r:
                    text = "✔ 成功" if r["status"] == "ok" else f"✘ {r['reason']}"
                    self.sync_table.setItem(row, 2, QTableWidgetItem(text))
        else:
            for row, r in enumerate(self._results):
                if row >= self.tpl_table.rowCount():
                    break
                text = "✔ 已生成" if r["status"] == "ok" and r["output"] else (
                    "✔ 预演完成" if r["status"] == "ok" else f"✘ {r['reason']}")
                self.tpl_table.setItem(row, 2, QTableWidgetItem(text))

    @staticmethod
    def _make_table(headers: list[str]) -> QTableWidget:
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(len(headers) - 1, QHeaderView.Stretch)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setMinimumHeight(160)
        return table

    # ---------- tab 1：多报告批量同步 ----------

    def _build_sync_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        self.tip_label(tab, "每份报告自动配同名工作簿，逐个同步；单份失败或文件被占用只跳过该份，不影响其他。")
        content_row = QHBoxLayout()
        content_row.setSpacing(8)
        content_row.addWidget(QLabel("同步内容："))
        self.batch_tables = make_checkbutton(tab, "表格", True)
        self.batch_marks = make_checkbutton(tab, "标注", False)
        self.batch_placeholders = make_checkbutton(tab, "占位符", False)
        for box in (self.batch_tables, self.batch_marks, self.batch_placeholders):
            content_row.addWidget(box)
        content_row.addStretch(1)
        layout.addLayout(content_row)
        self.batch_keep_highlight = make_checkbutton(tab, "同步标注后保留高亮（不去色，便于复核）", True)
        self.batch_keep_highlight.setVisible(False)
        self.batch_marks.toggled.connect(self.batch_keep_highlight.setVisible)
        layout.addWidget(self.batch_keep_highlight)
        row = QHBoxLayout()
        add = QPushButton("添加报告")
        add.clicked.connect(lambda _checked=False: self._add_reports())
        row.addWidget(add)
        remove = QPushButton("移除选中行")
        remove.clicked.connect(lambda _checked=False: self._remove_rows(self.sync_table))
        row.addWidget(remove)
        row.addStretch(1)
        layout.addLayout(row)
        self.sync_table = self._make_table(["报告（勾选参与）", "工作簿", "状态"])
        layout.addWidget(self.sync_table)
        self.sync_backup = make_checkbutton(tab, "同步前创建备份（取消勾选则直接覆盖）", False)
        layout.addWidget(self.sync_backup)
        self.button_row(tab, lambda: self._start_sync(True), lambda: self._start_sync(False), "全部仅预演", "开始批量同步")
        layout.addStretch(1)
        return tab

    def _add_reports(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "选择要批量同步的 Word 报告", "", "Word 文档 (*.docx)")
        existing = {self.sync_table.item(r, 0).data(Qt.UserRole) for r in range(self.sync_table.rowCount())}
        for p in paths:
            if p in existing:
                continue
            row = self.sync_table.rowCount()
            self.sync_table.insertRow(row)
            item = QTableWidgetItem(Path(p).name)
            item.setData(Qt.UserRole, p)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            self.sync_table.setItem(row, 0, item)
            xlsx = api.workbook_path(p)
            note = xlsx.name if xlsx.exists() else f"{xlsx.name}（未找到，将跳过）"
            self.sync_table.setItem(row, 1, QTableWidgetItem(note))
            self.sync_table.setItem(row, 2, QTableWidgetItem("待执行"))

    @staticmethod
    def _remove_rows(table: QTableWidget) -> None:
        for row in sorted({i.row() for i in table.selectedIndexes()}, reverse=True):
            table.removeRow(row)

    def _start_sync(self, dry_run: bool) -> None:
        contents = tuple(
            key for key, box in (
                ("tables", self.batch_tables), ("marks", self.batch_marks), ("placeholders", self.batch_placeholders)
            ) if box.isChecked()
        )
        if not contents:
            self.ctx.error("未选择同步内容", "请至少勾选一项同步内容（表格/标注/占位符）。")
            return
        content_text = "＋".join(
            {"tables": "表格", "marks": "标注", "placeholders": "占位符"}[key] for key in contents
        )
        paths = [
            self.sync_table.item(r, 0).data(Qt.UserRole)
            for r in range(self.sync_table.rowCount())
            if self.sync_table.item(r, 0).checkState() == Qt.Checked
        ]
        if not paths:
            self.ctx.info("未选择报告", "请先点「添加报告」并勾选要同步的报告。")
            return
        if not dry_run and not self.ctx.confirm(
            "确认批量同步",
            f"将把 {len(paths)} 份报告各自的工作簿改动回写到 Word。\n同步内容：{content_text}\n"
            + ("同步前会备份有改动的文件。" if self.sync_backup.isChecked() else "未勾选备份：直接覆盖原文件。")
            + "\n\n请先关闭这些 Word 与 Excel。确认开始？",
        ):
            return
        backup = self.sync_backup.isChecked()
        keep_highlight = self.batch_keep_highlight.isChecked()

        def run() -> int:
            self._results = api.run_batch_writeback(
                paths, dry_run=dry_run, backup=backup, contents=contents, keep_highlight=keep_highlight
            )
            return 0

        if self.ctx.run_steps([(("批量同步（预演）" if dry_run else "批量同步") + f"·{content_text}", run)]):
            self._waiting = "sync"
            for r in range(self.sync_table.rowCount()):
                self.sync_table.setItem(r, 2, QTableWidgetItem("执行中…"))

    # ---------- tab 2：模板生成 ----------

    def _build_template_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        self.tip_label(
            tab,
            "每个 Excel 同步到所选模板的一份副本，生成新 Word（命名：模板名_Excel名.docx）。"
            "模板原件与源 Excel 都不改动。多模板时在清单「所属模板」列给每个 Excel 指定模板编号（分组）。"
            "工作簿必须是从对应模板导出的，否则该份跳过。"
            "若模板正文含 {字段名} 占位符且勾选「占位符」，先点「生成正文信息填写表」，在工作簿「正文信息」表填值后再生成；未填齐的份会跳过。",
        )
        content_row = QHBoxLayout()
        content_row.setSpacing(8)
        content_row.addWidget(QLabel("同步内容："))
        self.tpl_tables = make_checkbutton(tab, "表格", True)
        self.tpl_marks = make_checkbutton(tab, "标注", False)
        self.tpl_placeholders = make_checkbutton(tab, "占位符", True)
        for box in (self.tpl_tables, self.tpl_marks, self.tpl_placeholders):
            content_row.addWidget(box)
        content_row.addStretch(1)
        layout.addLayout(content_row)
        self.tpl_keep_highlight = make_checkbutton(tab, "同步标注后保留高亮（不去色，便于复核）", True)
        self.tpl_keep_highlight.setVisible(False)
        self.tpl_marks.toggled.connect(self.tpl_keep_highlight.setVisible)
        layout.addWidget(self.tpl_keep_highlight)
        mode_row = QHBoxLayout()
        self.mode_single = QRadioButton("单模板")
        self.mode_single.setChecked(True)
        self.mode_multi = QRadioButton("多模板")
        self.mode_single.toggled.connect(lambda _c: self._mode_changed())
        mode_row.addWidget(self.mode_single)
        mode_row.addWidget(self.mode_multi)
        mode_row.addStretch(1)
        layout.addLayout(mode_row)

        tpl_row = QHBoxLayout()
        self.btn_pick_tpl = QPushButton("选择模板")
        self.btn_pick_tpl.clicked.connect(lambda _checked=False: self._pick_templates())
        tpl_row.addWidget(self.btn_pick_tpl)
        self.btn_del_tpl = QPushButton("移除选中模板")
        self.btn_del_tpl.clicked.connect(lambda _checked=False: self._remove_template())
        tpl_row.addWidget(self.btn_del_tpl)
        tpl_row.addStretch(1)
        layout.addLayout(tpl_row)
        self.tpl_label = QLabel("尚未选择模板")
        self.tpl_label.setObjectName("CardDesc")
        self.tpl_label.setWordWrap(True)
        self.tpl_label.setMinimumWidth(0)
        layout.addWidget(self.tpl_label)
        self.tpl_list = QListWidget()
        self.tpl_list.setMaximumHeight(90)
        layout.addWidget(self.tpl_list)

        xl_row = QHBoxLayout()
        add = QPushButton("添加 Excel 工作簿")
        add.clicked.connect(lambda _checked=False: self._add_workbooks())
        xl_row.addWidget(add)
        remove = QPushButton("移除选中行")
        remove.clicked.connect(lambda _checked=False: self._remove_rows(self.tpl_table))
        xl_row.addWidget(remove)
        fill = QPushButton("生成正文信息填写表")
        fill.setToolTip("按模板里的 {字段名} 占位符，给各工作簿加入/补齐「正文信息」表（保留已填值），供填写正文内容。")
        fill.clicked.connect(lambda _checked=False: self._fill_placeholder_sheets())
        xl_row.addWidget(fill)
        xl_row.addStretch(1)
        layout.addLayout(xl_row)
        self.tpl_table = self._make_table(["Excel 工作簿", "所属模板", "状态"])
        self.tpl_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        layout.addWidget(self.tpl_table)

        out_row = QHBoxLayout()
        out_row.addWidget(QLabel("输出目录"))
        self.out_edit = QLineEdit()
        out_row.addWidget(self.out_edit, 1)
        browse = QPushButton("浏览…")
        browse.clicked.connect(lambda _checked=False: self._pick_out_dir())
        out_row.addWidget(browse)
        layout.addLayout(out_row)
        self.button_row(tab, lambda: self._start_tpl(True), lambda: self._start_tpl(False), "全部仅预演", "开始生成")
        layout.addStretch(1)
        self._mode_changed()
        return tab

    def _templates(self) -> list[str]:
        if self.mode_single.isChecked():
            text = self.tpl_label.text()
            return [text] if text != "尚未选择模板" and Path(text).exists() else []
        return [self.tpl_list.item(i).data(Qt.UserRole) for i in range(self.tpl_list.count())]

    def _mode_changed(self) -> None:
        multi = self.mode_multi.isChecked()
        self.tpl_list.setVisible(multi)
        self.tpl_label.setVisible(not multi)
        self.btn_del_tpl.setVisible(multi)
        self.btn_pick_tpl.setText("添加模板" if multi else "选择模板")
        self.tpl_table.setColumnHidden(1, not multi)
        self._refresh_tpl_combos()

    def _pick_templates(self) -> None:
        if self.mode_single.isChecked():
            path, _ = QFileDialog.getOpenFileName(self, "选择 Word 模板", "", "Word 文档 (*.docx)")
            if path:
                self.tpl_label.setText(path)
                if not self.out_edit.text():
                    self.out_edit.setText(str(Path(path).parent))
        else:
            paths, _ = QFileDialog.getOpenFileNames(self, "添加 Word 模板（按添加顺序编号）", "", "Word 文档 (*.docx)")
            existing = self._templates()
            for p in paths:
                if p not in existing:
                    item = QListWidgetItem(f"模板{self.tpl_list.count() + 1} · {Path(p).name}")
                    item.setData(Qt.UserRole, p)
                    self.tpl_list.addItem(item)
            if self._templates() and not self.out_edit.text():
                self.out_edit.setText(str(Path(self._templates()[0]).parent))
        self._refresh_tpl_combos()

    def _remove_template(self) -> None:
        for item in self.tpl_list.selectedItems():
            self.tpl_list.takeItem(self.tpl_list.row(item))
        for i in range(self.tpl_list.count()):
            path = self.tpl_list.item(i).data(Qt.UserRole)
            self.tpl_list.item(i).setText(f"模板{i + 1} · {Path(path).name}")
        self._refresh_tpl_combos()

    def _refresh_tpl_combos(self) -> None:
        n = len(self._templates())
        for row in range(self.tpl_table.rowCount()):
            combo = self.tpl_table.cellWidget(row, 1)
            if combo is None:
                continue
            current = combo.currentIndex()
            combo.clear()
            combo.addItems([f"模板{i + 1}" for i in range(n)])
            if 0 <= current < n:
                combo.setCurrentIndex(current)

    def _add_workbooks(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "选择 Excel 工作簿", "", "Excel 工作簿 (*.xlsx)")
        existing = {self.tpl_table.item(r, 0).data(Qt.UserRole) for r in range(self.tpl_table.rowCount())}
        for p in paths:
            if p in existing:
                continue
            row = self.tpl_table.rowCount()
            self.tpl_table.insertRow(row)
            item = QTableWidgetItem(Path(p).name)
            item.setData(Qt.UserRole, p)
            self.tpl_table.setItem(row, 0, item)
            combo = QComboBox()
            combo.addItems([f"模板{i + 1}" for i in range(len(self._templates()))])
            self.tpl_table.setCellWidget(row, 1, combo)
            self.tpl_table.setItem(row, 2, QTableWidgetItem("待执行"))

    def _pick_out_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "选择输出目录", self.out_edit.text() or "")
        if path:
            self.out_edit.setText(path)

    def _fill_placeholder_sheets(self) -> None:
        templates = self._templates()
        if not templates:
            self.ctx.info("未选择模板", "请先选择 Word 模板，再据其占位符生成填写表。")
            return
        if self.tpl_table.rowCount() == 0:
            self.ctx.info("未添加工作簿", "请先点「添加 Excel 工作簿」。")
            return
        groups: dict[int, list[str]] = {}
        unassigned = 0
        for row in range(self.tpl_table.rowCount()):
            combo = self.tpl_table.cellWidget(row, 1)
            tpl_no = 1 if self.mode_single.isChecked() else (combo.currentIndex() + 1 if combo else 0)
            path = self.tpl_table.item(row, 0).data(Qt.UserRole)
            if not 1 <= tpl_no <= len(templates):
                unassigned += 1
                continue
            groups.setdefault(tpl_no, []).append(path)
        if not groups:
            self.ctx.info("未指定模板", "请先在「所属模板」列给工作簿选择模板。")
            return
        total = sum(len(v) for v in groups.values())
        msg = (f"将按模板里的 {{字段名}} 占位符，给 {total} 个工作簿加入/补齐「正文信息」表（保留已填值）。\n\n"
               "这会修改源工作簿（新增一张工作表），请先关闭这些 Excel。")
        if unassigned:
            msg += f"\n另有 {unassigned} 个未指定模板的工作簿将跳过。"
        if not self.ctx.confirm("生成正文信息填写表", msg + "\n\n确认开始？"):
            return

        def run() -> int:
            for tpl_no, paths in groups.items():
                api.fill_placeholder_sheets(templates[tpl_no - 1], paths)
            return 0

        self.ctx.run_steps([("生成正文信息填写表", run)])

    def _start_tpl(self, dry_run: bool) -> None:
        contents = tuple(
            key for key, box in (
                ("tables", self.tpl_tables), ("marks", self.tpl_marks), ("placeholders", self.tpl_placeholders)
            ) if box.isChecked()
        )
        if not contents:
            self.ctx.error("未选择同步内容", "请至少勾选一项同步内容（表格/标注/占位符）。")
            return
        content_text = "＋".join(
            {"tables": "表格", "marks": "标注", "placeholders": "占位符"}[key] for key in contents
        )
        templates = self._templates()
        if not templates:
            self.ctx.info("未选择模板", "请先选择 Word 模板。")
            return
        if self.tpl_table.rowCount() == 0:
            self.ctx.info("未添加工作簿", "请先点「添加 Excel 工作簿」。")
            return
        out_dir = self.out_edit.text().strip()
        if not out_dir:
            self.ctx.info("未指定输出目录", "请先指定生成 Word 的输出目录。")
            return
        items = []
        for row in range(self.tpl_table.rowCount()):
            combo = self.tpl_table.cellWidget(row, 1)
            tpl_no = 1 if self.mode_single.isChecked() else (combo.currentIndex() + 1 if combo else 0)
            items.append({"excel": self.tpl_table.item(row, 0).data(Qt.UserRole), "template": tpl_no})
        if not dry_run and not self.ctx.confirm(
            "确认模板生成",
            f"将用 {len(templates)} 个模板为 {len(items)} 个工作簿各生成一份新 Word：\n\n同步内容：{content_text}\n输出目录：{out_dir}\n\n"
            "模板原件与源 Excel 不会改动；重名输出自动加序号。确认开始？",
        ):
            return
        keep_highlight = self.tpl_keep_highlight.isChecked()

        def run() -> int:
            self._results = api.run_template_batch(
                templates, items, out_dir, dry_run=dry_run, contents=contents, keep_highlight=keep_highlight
            )
            return 0

        if self.ctx.run_steps([(("模板生成（预演）" if dry_run else "模板生成") + f"·{content_text}", run)]):
            self._waiting = "tpl"
            for r in range(self.tpl_table.rowCount()):
                self.tpl_table.setItem(r, 2, QTableWidgetItem("执行中…"))