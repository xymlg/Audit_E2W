"""同步页：Excel → Word（含增删行、增删列、按「书签：新增」标记创建新表）。"""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ... import api
from ...console import set_log_sink
from ...errors import SyncError
from ..widgets.cards import make_checkbutton
from ._page import BasePage


class WritebackPage(BasePage):
    title = "同步Word"
    subtitle = "工作簿改动回写 Word，保留格式、支持增删行列；带「新增」标记的表块会在 Word 里建出新表。"

    def build(self, body: QWidget) -> None:
        card = self.card(body)
        tabs = QTabWidget()
        card.layout().addWidget(tabs)
        tabs.addTab(self._build_sync_tab(), "同步回写")
        tabs.addTab(self._build_sort_tab(), "排序设置")
        tabs.addTab(self._build_diff_tab(), "差异比对")

    # ---------- 同步回写 tab ----------

    def _build_sync_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        content_row = QHBoxLayout()
        content_row.setSpacing(8)
        content_row.addWidget(QLabel("同步内容："))
        self.sync_tables = make_checkbutton(tab, "表格", True)
        self.sync_marks = make_checkbutton(tab, "标注", False)
        self.sync_placeholders = make_checkbutton(tab, "占位符", False)
        for box in (self.sync_tables, self.sync_marks, self.sync_placeholders):
            content_row.addWidget(box)
        content_row.addStretch(1)
        layout.addLayout(content_row)
        self.wb_keep_highlight = make_checkbutton(tab, "同步标注后保留高亮（不去色，便于复核）", True)
        self.wb_keep_highlight.setVisible(False)
        self.sync_marks.toggled.connect(self.wb_keep_highlight.setVisible)
        layout.addWidget(self.wb_keep_highlight)
        self.tip_label(
            tab,
            "标注/占位符分别读取「链接标注」「正文信息」工作表（在「标注与占位」页导出并填写）；"
            "未填的内容保留原样并逐条警告。备份、改动清单、删除✕行只作用于表格。",
        )
        self.force = make_checkbutton(tab, "强制跳过文件占用检查（文件仍被打开时用）", False)
        layout.addWidget(self.force)
        self.backup = make_checkbutton(tab, "同步前创建备份（取消勾选则不备份，文件直接覆盖）", False)
        layout.addWidget(self.backup)
        self.changelog = make_checkbutton(tab, "同时导出改动清单（逐格 旧值→新值，存到工作簿旁，作复核底稿）", False)
        layout.addWidget(self.changelog)
        self.delete_empty = make_checkbutton(
            tab, "删除✕行：同步时把 B 列✕（导出时判定为空）的行连 Word 行一起删除；Excel 该行保留为「已删」，改回✓可恢复", False
        )
        layout.addWidget(self.delete_empty)

        row = self.button_row(tab, lambda: self._start(True), lambda: self._start(False), "仅预演", "开始同步")
        btn = QPushButton("追加新表块")
        btn.clicked.connect(lambda _checked=False: self._append_block())
        row.layout().insertWidget(row.layout().count() - 1, btn)
        self.tip_label(
            tab,
            "要在报告里加一张新表：点「追加新表块」→ 选插在哪张表下面 → 新表块会直接插到 Excel 里那张表的下面"
            "（后续表块下移、编号自动顺延）→ 填好内容（首行为表头，可合并单元格、写公式）→ 点「开始同步」，"
            "新表会建到 Word 里同一位置。",
        )
        layout.addStretch(1)
        return tab

    # ---------- 排序设置 tab ----------

    def _build_sort_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        self.tip_label(
            tab,
            "排序设置：给某张表选定金额列后，每次同步自动按设置重排（Excel 与 Word 一起排），"
            "设置保存在工作簿里、随文件走，不用每次重设。项目名为「其他」的行不参与排序，放在合计行之前；"
            "表头与合计行不动；数据区有合并单元格的表会跳过排序并提示。",
        )
        row = QHBoxLayout()
        row.setSpacing(8)
        row.addWidget(QLabel("表"))
        self.sort_table = QComboBox()
        self.sort_table.setMinimumWidth(220)
        self.sort_table.currentIndexChanged.connect(lambda _i: self._sort_table_changed())
        row.addWidget(self.sort_table, 2)
        row.addWidget(QLabel("金额列"))
        self.sort_column = QComboBox()
        self.sort_column.setMinimumWidth(150)
        row.addWidget(self.sort_column, 1)
        row.addWidget(QLabel("方向"))
        self.sort_dir = QComboBox()
        self.sort_dir.addItems(["降序（从大到小）", "升序（从小到大）"])
        row.addWidget(self.sort_dir)
        layout.addLayout(row)

        btns = QHBoxLayout()
        btns.setSpacing(12)
        refresh = QPushButton("刷新清单")
        refresh.clicked.connect(lambda _checked=False: self._reload_sorts())
        btns.addWidget(refresh)
        save = QPushButton("保存排序设置")
        save.setObjectName("Primary")
        save.clicked.connect(lambda _checked=False: self._save_sort())
        btns.addWidget(save)
        clear = QPushButton("清除该表排序")
        clear.clicked.connect(lambda _checked=False: self._clear_sort())
        btns.addWidget(clear)
        btns.addStretch(1)
        layout.addLayout(btns)
        self.sort_status = self.tip_label(tab, "尚未加载。选择 Word 文件并生成工作簿后，点「刷新清单」。")
        self._sort_items: list[dict] = []
        layout.addStretch(1)
        return tab

    def _append_block(self) -> None:
        if not self.ctx.require_docx():
            return
        excel = self.ctx.workbook_path()
        if not excel.exists():
            self.ctx.error("未找到工作簿", f"未找到工作簿：\n{excel}\n\n请先执行「导出表格到 Excel」生成工作簿。")
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("追加新表块")
        dlg.setModal(True)
        grid = QGridLayout(dlg)
        grid.setContentsMargins(16, 16, 16, 16)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)
        grid.addWidget(QLabel("表名（将成为 Word 里的表标题）："), 0, 0, 1, 2)
        name_edit = QLineEdit()
        grid.addWidget(name_edit, 1, 0, 1, 2)
        grid.addWidget(QLabel("行数（含表头）"), 2, 0)
        rows_spin = QSpinBox()
        rows_spin.setRange(1, 200)
        rows_spin.setValue(5)
        grid.addWidget(rows_spin, 2, 1)
        grid.addWidget(QLabel("列数"), 3, 0)
        cols_spin = QSpinBox()
        cols_spin.setRange(1, 30)
        cols_spin.setValue(4)
        grid.addWidget(cols_spin, 3, 1)
        grid.addWidget(QLabel("插在哪张表下面（Excel 与 Word 建表位置一致）"), 4, 0, 1, 2)
        pos_combo = QComboBox()
        pos_map: dict[str, str] = {}
        pos_combo.addItem("（最后一张表下面）")
        try:
            for item in api.list_table_blocks(self.ctx.docx, self.ctx.excel):
                short = item["name"].split(" · 章节：")[0]
                if len(short) > 24:
                    short = short[:23] + "…"
                label = f"{item['code']} · {short}" if short else item["code"]
                pos_map[label] = item["code"]
                pos_combo.addItem(label)
        except SyncError:
            pass
        grid.addWidget(pos_combo, 5, 0, 1, 2)

        def ok() -> None:
            after = pos_map.get(pos_combo.currentText())
            set_log_sink(lambda message: self.ctx.log(message))
            try:
                code = api.run_append_table_block(
                    self.ctx.docx,
                    self.ctx.excel,
                    name=name_edit.text(),
                    nrows=rows_spin.value(),
                    ncols=cols_spin.value(),
                    after=after,
                )
            except SyncError as exc:
                QMessageBox.critical(dlg, "追加失败", str(exc))
                return
            finally:
                set_log_sink(None)
            dlg.accept()
            self.ctx.log(f"新表块 {code} 已就绪：在 Excel 填好内容后点「开始同步」即可在 Word 里建表。", "ok")

        btns = QHBoxLayout()
        btns.addStretch(1)
        cancel = QPushButton("取消")
        cancel.clicked.connect(lambda _checked=False: dlg.reject())
        btns.addWidget(cancel)
        append = QPushButton("追加")
        append.setObjectName("Primary")
        append.clicked.connect(lambda _checked=False: ok())
        btns.addWidget(append)
        grid.addLayout(btns, 6, 0, 1, 2)
        name_edit.returnPressed.connect(ok)
        dlg.exec()

    def _start(self, dry_run: bool) -> None:
        if not self.ctx.require_docx():
            return
        tables = self.sync_tables.isChecked()
        marks = self.sync_marks.isChecked()
        placeholders = self.sync_placeholders.isChecked()
        if not (tables or marks or placeholders):
            self.ctx.error("未选择同步内容", "请至少勾选一项同步内容（表格/标注/占位符）。")
            return
        content_text = "＋".join(
            name for name, on in (("表格", tables), ("标注", marks), ("占位符", placeholders)) if on
        )
        excel = self.ctx.workbook_path()
        steps: list[tuple[str, object]] = []

        if not excel.exists():
            if dry_run:
                self.ctx.log(f"未找到工作簿：{excel}\n请先执行导出生成工作簿，再预演同步。", "error")
                return
            if not self.ctx.confirm(
                "未找到工作簿",
                f"未找到工作簿：\n{excel}\n\n是否现在从 Word 导出生成工作簿？\n生成后将继续执行同步。",
            ):
                self.ctx.log("已取消：缺少工作簿，未执行同步。", "muted")
                return
            steps.append(("导出到 Excel（生成工作簿）", lambda: api.run_export(self.ctx.docx, self.ctx.excel)))
        elif not dry_run:
            backup_note = (
                "同步前会自动备份 Word（若数据有变化）。" if self.backup.isChecked()
                else "未勾选备份：将直接覆盖原文件，不创建 _backup 备份。"
            )
            message = (
                f"将把以下内容同步到 Word：{content_text}\n\nWord：{self.ctx.docx}\nExcel：{excel}\n\n"
                + backup_note + " 确认开始同步？"
            )
            if not self.ctx.confirm("确认同步", message):
                self.ctx.log("已取消同步。", "muted")
                return

        force = self.force.isChecked()
        backup = self.backup.isChecked()
        changelog = self.changelog.isChecked()
        delete_empty = self.delete_empty.isChecked()
        keep_highlight = self.wb_keep_highlight.isChecked()
        label = ("同步到 Word（预演）" if dry_run else "同步到 Word") + f"·{content_text}"
        if tables and not marks and not placeholders:
            steps.append(
                (
                    label,
                    lambda: api.run_writeback(self.ctx.docx, self.ctx.excel, dry_run=dry_run, force=force, backup=backup, changelog=changelog, delete_empty_rows=delete_empty),
                )
            )
        else:
            def run_sync() -> int:
                api.run_sync_content(
                    self.ctx.docx, self.ctx.excel,
                    tables=tables, marks=marks, placeholders=placeholders,
                    keep_highlight=keep_highlight, dry_run=dry_run, force=force,
                    backup=backup, changelog=changelog, delete_empty_rows=delete_empty,
                )
                return 0

            steps.append((label, run_sync))
        self.ctx.run_steps(steps)

    # ---------- 排序设置 ----------

    def on_show(self) -> None:
        if not self._sort_items:
            self._reload_sorts(quiet=True)

    def _reload_sorts(self, quiet: bool = False) -> None:
        self._sort_items = []
        self.sort_table.clear()
        self.sort_column.clear()
        if self.ctx.docx is None or not self.ctx.workbook_path().exists():
            self.sort_status.setText("尚未加载。选择 Word 文件并生成工作簿后，点「刷新清单」。")
            return
        try:
            self._sort_items = api.list_table_sorts(self.ctx.docx, self.ctx.excel)
        except SyncError as exc:
            self.sort_status.setText(f"加载失败：{exc}")
            return
        for item in self._sort_items:
            short = item["name"].split(" · 章节：")[0]
            if len(short) > 24:
                short = short[:23] + "…"
            mark = f"（已设：{item['sort_column']} {'降序' if item['descending'] else '升序'}）" if item["has_sort"] else ""
            self.sort_table.addItem(f"{item['code']} · {short}{mark}")
        if self._sort_items:
            self.sort_table.setCurrentIndex(0)
            self._sort_table_changed()
            self.sort_status.setText(f"共 {len(self._sort_items)} 张表可设置排序；带「已设」标记的表同步时会自动排序。")
        else:
            self.sort_status.setText("工作簿里没有可设置排序的表块。")
        if not quiet:
            self.ctx.log("排序清单已刷新。", "muted")

    def _sort_table_changed(self) -> None:
        self.sort_column.clear()
        idx = self.sort_table.currentIndex()
        if not 0 <= idx < len(self._sort_items):
            return
        item = self._sort_items[idx]
        seen = []
        for h in item["headers"]:
            if h and h not in seen:
                seen.append(h)
        self.sort_column.addItems(seen)
        if item["has_sort"] and item["sort_column"] in seen:
            self.sort_column.setCurrentText(item["sort_column"])
            self.sort_dir.setCurrentIndex(0 if item["descending"] else 1)

    def _save_sort(self) -> None:
        idx = self.sort_table.currentIndex()
        if not 0 <= idx < len(self._sort_items):
            self.ctx.info("未选择", "请先点「刷新清单」并选择要排序的表。")
            return
        if not self.sort_column.currentText():
            self.ctx.info("未选择金额列", "请先选择按哪一列金额排序。")
            return
        item = self._sort_items[idx]
        set_log_sink(lambda message: self.ctx.log(message))
        try:
            api.set_table_sort(
                self.ctx.docx, self.ctx.excel,
                code=item["code"], column=self.sort_column.currentText(),
                descending=self.sort_dir.currentIndex() == 0,
            )
        except SyncError as exc:
            self.ctx.error("保存排序设置失败", str(exc))
            return
        finally:
            set_log_sink(None)
        self._reload_sorts(quiet=True)

    def _clear_sort(self) -> None:
        idx = self.sort_table.currentIndex()
        if not 0 <= idx < len(self._sort_items):
            self.ctx.info("未选择", "请先点「刷新清单」并选择表。")
            return
        item = self._sort_items[idx]
        set_log_sink(lambda message: self.ctx.log(message))
        try:
            api.clear_table_sort(self.ctx.docx, self.ctx.excel, code=item["code"])
        except SyncError as exc:
            self.ctx.error("清除排序设置失败", str(exc))
            return
        finally:
            set_log_sink(None)
        self._reload_sorts(quiet=True)

    # ---------- 差异比对 ----------

    def _build_diff_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        self.tip_label(
            tab,
            "差异比对：只读列出 Word 与工作簿当前不一致的单元格（旧值→新值），不写任何文件，"
            "供同步前复核。点「差异比对」刷新下表；可「导出差异表」存档。",
        )
        row = QHBoxLayout()
        row.setSpacing(12)
        refresh = QPushButton("差异比对")
        refresh.setObjectName("Primary")
        refresh.setToolTip("只读列出 Word 与工作簿不一致的单元格（旧值→新值），不写任何文件。")
        refresh.clicked.connect(lambda _checked=False: self._diff())
        row.addWidget(refresh)
        export = QPushButton("导出差异表")
        export.clicked.connect(lambda _checked=False: self._export_diff())
        row.addWidget(export)
        row.addStretch(1)
        layout.addLayout(row)
        self.diff_status = self.tip_label(tab, "尚未比对。选择 Word 文件并生成工作簿后，点「差异比对」。")
        self.diff_table = QTableWidget(0, 7)
        self.diff_table.setHorizontalHeaderLabels(["表号", "行", "列", "项目", "旧值", "新值", "类型"])
        self.diff_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.diff_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.diff_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.diff_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.diff_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.diff_table.setMinimumHeight(180)
        layout.addWidget(self.diff_table)
        self._diff_records: list[dict] = []
        return tab

    def _diff(self) -> None:
        if not self.ctx.require_docx():
            return
        if not self.ctx.workbook_path().exists():
            self.ctx.error("未找到工作簿", "请先执行「导出表格到 Excel」生成工作簿，再比对差异。")
            return
        set_log_sink(lambda message: self.ctx.log(message))
        try:
            records = api.scan_writeback_diff(self.ctx.docx, self.ctx.excel)
        except SyncError as exc:
            self.ctx.error("差异比对失败", str(exc))
            return
        finally:
            set_log_sink(None)
        self._diff_records = records
        self._fill_diff(records)

    def _fill_diff(self, records: list[dict]) -> None:
        self.diff_table.setRowCount(0)
        for rec in records:
            row = self.diff_table.rowCount()
            self.diff_table.insertRow(row)
            values = [rec.get("code", ""), rec.get("row", ""), rec.get("col", ""), rec.get("item", ""),
                      rec.get("old", ""), rec.get("new", ""), rec.get("kind", "")]
            for c, value in enumerate(values):
                self.diff_table.setItem(row, c, QTableWidgetItem(str(value)))
        cells = sum(1 for r in records if r.get("kind") == "单元格")
        if records:
            self.diff_status.setText(f"共 {len(records)} 条差异（其中单元格改动 {cells} 处）。确认无误后点「开始同步」。")
        else:
            self.diff_status.setText("Word 与工作簿一致，没有差异。")

    def _export_diff(self) -> None:
        if not self._diff_records:
            self.ctx.info("无差异可导出", "请先点「差异比对」；当前没有差异记录。")
            return
        set_log_sink(lambda message: self.ctx.log(message))
        try:
            path = api.export_diff_report(self.ctx.docx, self.ctx.excel, self._diff_records)
        except SyncError as exc:
            self.ctx.error("导出差异表失败", str(exc))
            return
        finally:
            set_log_sink(None)
        if path:
            self.ctx.log(f"差异表已导出：{path}", "ok")