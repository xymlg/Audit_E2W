"""桌面格式页：选择样板、预演格式方案、生成格式调整副本。"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import (
    QComboBox,
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from ... import api
from ...errors import SyncError
from ...report_format import DEFAULT_SCOPE, SCOPE_LABELS
from ._page import BasePage
from ..widgets.cards import make_checkbutton


class FormatPage(BasePage):
    title = "一键调整格式"
    subtitle = "按样板或已存方案统一字体、标题、正文和表格格式；只生成新副本，原报告不改动。"

    def build(self, body: QWidget) -> None:
        card = self.card(body)
        self.tip_label(
            card,
            "先在顶部选择要调整的报告，再到第 2 个 Tab“调整项目”勾选需要修改的内容。"
            "小报表字号保留，页面设置、分节、页眉页脚、表格结构和审计痕迹均不改动。",
        )
        self.tabs = QTabWidget()
        card.layout().addWidget(self.tabs)
        self.tabs.addTab(self._build_apply_tab(), "格式调整")
        self.tabs.addTab(self._build_scope_tab(), "调整项目")

        result_card = self.card(body)
        self.status_label = self.tip_label(result_card, "尚未检查。选择样板或已存方案后，先点“检查格式方案”。")
        self.details = QPlainTextEdit()
        self.details.setReadOnly(True)
        self.details.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        self.details.setMinimumHeight(240)
        result_card.layout().addWidget(self.details)

        self._waiting: str | None = None
        self._result: dict | None = None
        self._scan_key: tuple[str, str, str, tuple[bool, ...]] | None = None
        self._pending_key: tuple[str, str, str, tuple[bool, ...]] | None = None
        self._pending_output = ""
        self._output_auto = True
        self._last_target: Path | None = None
        self._reload_profiles()
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._check_done)
        self._timer.start(250)

    def _build_apply_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 4, 0, 0)
        layout.setSpacing(8)
        target_row = QHBoxLayout()
        target_row.addWidget(QLabel("待调整报告："))
        self.target_label = QLabel("尚未选择")
        self.target_label.setWordWrap(True)
        self.target_label.setMinimumWidth(0)
        target_row.addWidget(self.target_label, 1)
        layout.addLayout(target_row)

        profile_row = QHBoxLayout()
        profile_row.addWidget(QLabel("格式来源："))
        self.profile_combo = QComboBox()
        self.profile_combo.addItem("从样板报告提取", "")
        self.profile_combo.currentIndexChanged.connect(self._invalidate)
        self.profile_combo.currentIndexChanged.connect(self._update_profile_mode)
        profile_row.addWidget(self.profile_combo, 1)
        self.save_profile_button = QPushButton("保存当前样板为方案…")
        self.save_profile_button.clicked.connect(self._save_profile)
        profile_row.addWidget(self.save_profile_button)
        self.delete_profile_button = QPushButton("删除方案")
        self.delete_profile_button.clicked.connect(self._delete_profile)
        profile_row.addWidget(self.delete_profile_button)
        layout.addLayout(profile_row)

        template_row = QHBoxLayout()
        template_row.addWidget(QLabel("样板报告："))
        self.template_edit = QLineEdit()
        self.template_edit.setPlaceholderText("选择格式标准的 Word 报告；使用已存方案时无需选择")
        self.template_edit.textChanged.connect(self._invalidate)
        template_row.addWidget(self.template_edit, 1)
        self.template_button = QPushButton("浏览…")
        self.template_button.clicked.connect(self._browse_template)
        template_row.addWidget(self.template_button)
        layout.addLayout(template_row)

        output_row = QHBoxLayout()
        output_row.addWidget(QLabel("输出副本："))
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("默认保存在原报告旁：原名_格式调整.docx")
        self.output_edit.textChanged.connect(self._invalidate)
        self.output_edit.textEdited.connect(lambda _text: setattr(self, "_output_auto", False))
        output_row.addWidget(self.output_edit, 1)
        output_button = QPushButton("浏览…")
        output_button.clicked.connect(self._browse_output)
        output_row.addWidget(output_button)
        layout.addLayout(output_row)

        row = self.button_row(tab, self._start_scan, self._start_apply, "检查格式方案", "生成格式副本")
        self.apply_button = row.layout().itemAt(1).widget()
        self.apply_button.setEnabled(False)
        layout.addStretch(1)
        return tab

    def _build_scope_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(0, 4, 0, 0)
        layout.setSpacing(8)
        self.tip_label(tab, "只调整勾选的项目；未勾选项目的字体、字号、段落设置、对齐和加粗均保持原样。修改选择后需重新检查。")
        self.scope_buttons: dict[str, QCheckBox] = {}
        descriptions = {
            "cover": "封面与目录（封面标题、封面信息、目录标题和条目）",
            "headings": "一至四级标题",
            "statement": "报表名称、日期与表格标题",
            "body": "正文段落",
            "table_text": "表格字体与字号（小报表字号不放大）",
            "table_emphasis": "表格对齐、表头与合计行加粗",
            "table_borders": "表格框线（外框、内部横线和竖线）",
        }
        for key in DEFAULT_SCOPE:
            button = make_checkbutton(tab, descriptions[key], checked=True)
            button.toggled.connect(self._invalidate)
            layout.addWidget(button)
            self.scope_buttons[key] = button
        self.tip_label(
            tab,
            "智能保护：按报告角色区分处理，不把所有段落套成同一种格式；空段落和大纲级别异常段落保持原样。"
            "不改文字、数字、书签、超链接、图片、域代码、颜色、高亮、下划线、表格结构及页面设置。",
        )
        layout.addStretch(1)
        return tab

    def _scope(self) -> dict[str, bool]:
        return {key: button.isChecked() for key, button in self.scope_buttons.items()}

    def on_show(self) -> None:
        self._refresh_target()
        self._reload_profiles()

    def _reload_profiles(self) -> None:
        selected = self.profile_combo.currentData()
        self.profile_combo.blockSignals(True)
        self.profile_combo.clear()
        self.profile_combo.addItem("从样板报告提取", "")
        try:
            for item in api.list_format_profiles():
                self.profile_combo.addItem(item["name"], item["name"])
        except SyncError as exc:
            self.ctx.error("格式方案读取失败", str(exc))
        index = self.profile_combo.findData(selected)
        self.profile_combo.setCurrentIndex(max(index, 0))
        self.profile_combo.blockSignals(False)
        if (self.profile_combo.currentData() or None) != (selected or None):
            self._invalidate()
        self._update_profile_mode()

    def _update_profile_mode(self) -> None:
        saved = bool(self.profile_combo.currentData())
        enabled = not self.ctx.busy
        self.template_edit.setEnabled(not saved)
        self.template_button.setEnabled(not saved and enabled)
        self.save_profile_button.setEnabled(not saved and enabled)
        self.delete_profile_button.setEnabled(saved and enabled)

    def _save_profile(self) -> None:
        template_text = self.template_edit.text().strip()
        if not template_text:
            self.ctx.error("未选择样板", "请先选择样板报告，再保存纯格式方案。")
            return
        template = Path(template_text)
        name, ok = QInputDialog.getText(self, "保存格式方案", "方案名称：", QLineEdit.Normal, template.stem)
        if not ok or not name.strip():
            return
        try:
            if any(item["name"] == name.strip() for item in api.list_format_profiles()):
                self.ctx.error("方案已存在", "请换一个方案名称；已有方案不会被覆盖。")
                return
        except SyncError as exc:
            self.ctx.error("方案读取失败", str(exc))
            return
        self._result = None
        self._waiting = "save"
        self.status_label.setText("正在提取并保存纯格式方案…")

        def run() -> int:
            self._result = api.save_format_profile(template, name, source=self.ctx.docx)
            return 0

        self.ctx.run_steps([("保存纯格式方案", run)])

    def _delete_profile(self) -> None:
        name = self.profile_combo.currentData()
        if not name:
            return
        if not self.ctx.confirm("删除格式方案", f"确定删除格式方案“{name}”？\n\n只删除格式参数，不删除任何报告。"):
            return
        try:
            api.delete_format_profile(name)
        except SyncError as exc:
            self.ctx.error("删除失败", str(exc))
            return
        self._reload_profiles()
        self._invalidate()
        self.status_label.setText(f"已删除格式方案：{name}")

    def _refresh_target(self) -> None:
        target = self.ctx.docx
        self.target_label.setText(str(target) if target else "尚未选择（请在顶部选择 Word 文件）")
        if target != self._last_target:
            self._last_target = target
            self._scan_key = None
            self._result = None
            self.apply_button.setEnabled(False)
            if target is not None and self._output_auto:
                suffix = target.suffix if target.suffix.lower() in (".docx", ".docm") else ".docx"
                self.output_edit.setText(str(target.with_name(f"{target.stem}_格式调整{suffix}")))

    def _invalidate(self) -> None:
        self._scan_key = None
        self._result = None
        self.apply_button.setEnabled(False)

    def _browse_template(self) -> None:
        initial = self.template_edit.text() or str(self.ctx.docx.parent if self.ctx.docx else "")
        path, _ = QFileDialog.getOpenFileName(self, "选择样板报告", initial, "Word 文档 (*.docx *.docm)")
        if path:
            self.template_edit.setText(path)

    def _browse_output(self) -> None:
        target = self.ctx.docx
        suffix = target.suffix if target and target.suffix.lower() in (".docx", ".docm") else ".docx"
        initial = self.output_edit.text() or str(target.with_name(f"{target.stem}_格式调整{suffix}") if target else f"格式调整{suffix}")
        path, _ = QFileDialog.getSaveFileName(self, "选择格式副本保存位置", initial, f"Word 文档 (*{suffix})")
        if path:
            if not path.lower().endswith(suffix):
                path += suffix
            self._output_auto = False
            self.output_edit.setText(path)

    def _paths(self) -> tuple[Path, Path | None, Path, str | None] | None:
        self._refresh_target()
        if not self.ctx.require_docx() or self.ctx.docx is None:
            return None
        profile_name = self.profile_combo.currentData() or None
        template_text = self.template_edit.text().strip()
        if profile_name is None and not template_text:
            self.ctx.error("未选择格式来源", "请选择样板报告，或改用已保存的格式方案。")
            return None
        output_text = self.output_edit.text().strip()
        if not output_text:
            self.ctx.error("未选择输出位置", "请选择格式副本的保存位置。")
            return None
        return (
            self.ctx.docx,
            Path(template_text) if template_text and profile_name is None else None,
            Path(output_text),
            profile_name,
        )

    def _source_key(
        self,
        target: Path,
        template: Path | None,
        profile_name: str | None,
        scope: dict[str, bool],
    ) -> tuple[str, str, str, tuple[bool, ...]]:
        return (
            str(target),
            "profile" if profile_name else "template",
            profile_name or str(template),
            tuple(scope[key] for key in DEFAULT_SCOPE),
        )

    def _start_scan(self) -> None:
        paths = self._paths()
        if paths is None:
            return
        target, template, output, profile_name = paths
        scope = self._scope()
        self._pending_key = self._source_key(target, template, profile_name, scope)
        self._pending_output = str(output)
        self._waiting = "scan"
        self._result = None
        self.status_label.setText("正在读取格式来源和待调整报告…")

        def run() -> int:
            self._result = api.scan_report_format(target, template, profile_name=profile_name, scope=scope)
            return 0

        self.ctx.run_steps([("检查报告格式方案", run)])

    def _start_apply(self) -> None:
        paths = self._paths()
        if paths is None:
            return
        target, template, output, profile_name = paths
        scope = self._scope()
        key = self._source_key(target, template, profile_name, scope)
        if self._scan_key != key or self._result is None:
            self.ctx.error("尚未预演", "请先点“检查格式方案”，确认影响范围后再生成副本。")
            return
        self._pending_key = key
        self._pending_output = str(output)
        totals = self._result["totals"]
        source_label = f"已存方案：{profile_name}" if profile_name else f"样板：{template.name}"
        if not self.ctx.confirm(
            "确认生成格式副本",
            f"将按以下格式来源调整：\n\n{source_label}\n报告：{target.name}\n"
            f"段落：{totals['selected_paragraphs']} 个\n表格：{totals['tables']} 张\n\n"
            f"输出：{output}\n\n原报告不会改动。确认生成格式副本？",
        ):
            self.ctx.log("已取消格式调整。", "muted")
            return
        self._result = None
        self._waiting = "apply"
        self.status_label.setText("正在生成并核验格式副本…")

        def run() -> int:
            self._result = api.run_report_format(target, template, output=output, profile_name=profile_name, scope=scope)
            return 0

        self.ctx.run_steps([("生成报告格式副本", run)])

    def _check_done(self) -> None:
        if not self._waiting or self.ctx.busy:
            return
        mode, self._waiting = self._waiting, None
        pending_key, pending_output = self._pending_key, self._pending_output
        self._pending_key = None
        self._pending_output = ""
        if self._result is None:
            self.status_label.setText("执行失败，请查看右侧日志。")
            return
        if mode == "save":
            name = self._result["name"]
            self._reload_profiles()
            index = self.profile_combo.findData(name)
            if index >= 0:
                self.profile_combo.setCurrentIndex(index)
            self._invalidate()
            self.status_label.setText(f"已保存纯格式方案：{name}")
        elif mode == "scan":
            profile_name = self.profile_combo.currentData() or None
            template_text = self.template_edit.text().strip()
            template = Path(template_text) if template_text and profile_name is None else None
            current_key = (
                self._source_key(self.ctx.docx, template, profile_name, self._scope())
                if self.ctx.docx is not None
                else None
            )
            unchanged = pending_key == current_key
            self._scan_key = pending_key if unchanged else None
            self.apply_button.setEnabled(unchanged)
            if unchanged:
                self._render_scan(self._result, pending_output)
            else:
                self.status_label.setText("格式来源或调整项目已变化，请重新检查格式方案。")
        else:
            output = self._result.get("output", pending_output or self.output_edit.text())
            self._scan_key = None
            self.apply_button.setEnabled(False)
            self.status_label.setText(f"格式副本已生成并核验：{output}")
            self.details.appendPlainText(f"\n已生成：{output}\n原报告未改动。")
            review = self._result.get("visual_review", [])
            if review:
                self.details.appendPlainText("\nWord 目视验收（结构核验不代表排版已通过）：")
                self.details.appendPlainText("\n".join(f"· {item}" for item in review))
            self.ctx.info("格式副本已生成", f"格式副本已生成并结构核验：\n{output}\n\n原报告未改动。请在 Word 中按“目视验收清单”复核排版。")

    def _render_scan(self, result: dict, output: str = "") -> None:
        totals = result["totals"]
        selected_labels = [SCOPE_LABELS[key] for key, enabled in result["scope"].items() if enabled]
        lines = [
            f"格式来源：{result['template']}",
            f"待调整：{result['source']}",
            f"计划输出：{output or result['suggested_output']}",
            "",
            "调整项目：",
        ]
        lines.extend(f"· {label}" for label in selected_labels)
        lines.extend([
            "",
            f"将处理 {totals['selected_paragraphs']} 个非空段落、{totals['tables']} 张表（{totals['table_rows']} 行）。",
            f"未选项目跳过 {totals['skipped_paragraphs']} 个段落、{totals['skipped_tables']} 张表；"
            f"异常大纲段落保持原样 {totals['protected_paragraphs']} 个；空段落保留 {totals['empty_paragraphs']} 个。",
            "",
        ])
        selected_roles = [item for item in result["roles"] if item["selected"]]
        if selected_roles:
            lines.append("段落格式方案：")
            for item in selected_roles:
                lines.append(f"· {item['label']}：{item['selected_paragraphs']} 段；{item['format']}")
        if result["scope"]["table_text"]:
            lines.append(f"· 表格字体与字号：{result['table_format']}")
        if result["scope"]["table_emphasis"]:
            lines.append("· 表格强调：表头居中，数值右对齐，普通文字左对齐，表头和合计行加粗。")
        if result["scope"]["table_borders"]:
            lines.append(f"· 表格框线：{result['table_border_format']}")
        if result["scope"]["table_text"]:
            lines.append(f"· {result['small_table_policy']}")
        if result["warnings"]:
            lines.append("")
            lines.append("提示：")
            lines.extend(f"· {item}" for item in result["warnings"])
        self.details.setPlainText("\n".join(lines))
        self.status_label.setText("预演完成，未写入任何文件。确认无误后点“生成格式副本”。")

    def refresh_button_state(self) -> None:
        self.apply_button.setEnabled(self._scan_key is not None and self._result is not None and not self.ctx.busy)
        self._update_profile_mode()
