"""GUI offscreen 回归：九页、按钮状态、空白表页、批量页与格式页链路。"""
import time

import pytest
from helpers import ROOT, SHEET_NAME, Document, api, blocks_of, check, load_workbook, make_doc, word_tables, ws_dir


@pytest.fixture(scope="module")
def win(qapp):
    from wesync.gui.app import WesyncApp
    w = WesyncApp(base_dir=ROOT / "excel_word")
    w.ctx.info = lambda *a, **k: None
    w.ctx.error = lambda *a, **k: None
    w.ctx.confirm = lambda *a, **k: True
    yield w
    w.close()


def wait_idle(qapp, win, timeout=60):
    end = time.time() + timeout
    while time.time() < end:
        qapp.processEvents()
        if not win.runner.busy and win.runner.queue.empty():
            qapp.processEvents()
            return True
        time.sleep(0.02)
    return False


def test_g01_pages(win):
    check("九个页面", len(win._pages) == 9, str(list(win._pages)))
    names = list(win._pages)
    check("格式页在批量同步后", names.index("BatchPage") < names.index("FormatPage") < names.index("HistoryPage"), str(names))


def test_g02_button_state(win, qapp):
    blank = win._pages["BlankPage"]
    check("初始删除按钮禁用", not blank.btn_delete.isEnabled())
    check("初始批量删除按钮禁用", not blank.batch_btn_delete.isEnabled())
    win.ctx.run_steps([("空任务", lambda: 0)])
    wait_idle(qapp, win)
    check("任务后删除按钮仍禁用", not blank.btn_delete.isEnabled())
    check("任务后批量删除按钮仍禁用", not blank.batch_btn_delete.isEnabled())


def test_g03_blank_flow(win, qapp):
    d = ws_dir("g03")
    docx = make_doc(d / "g.docx", [
        ("表1", [["项目", "金额"], ["A", "1"], ["合计", "1"]]),
        ("表2 空", [["项目", "本期"], ["", ""], ["", ""]]),
    ])
    api.run_export(docx)
    win.set_docx(docx)
    blank = win._pages["BlankPage"]
    blank._scan()
    check("扫描出1候选", len(blank._rows) == 1, str(len(blank._rows)))
    check("扫描后删除按钮启用", blank.btn_delete.isEnabled())
    blank._toggle_all(True)
    blank._delete()
    check("任务完成", wait_idle(qapp, win))
    check("Word剩1张", len(word_tables(docx)) == 1)
    check("删除后按钮回到禁用", not blank.btn_delete.isEnabled())
    blank._refresh_deleted()
    check("恢复清单1条", len(blank._restore_rows) == 1)
    blank._restore_rows[0][1].setChecked(True)
    blank._restore()
    wait_idle(qapp, win)
    check("恢复后2张", len(word_tables(docx)) == 2)


def test_g04_batch_backfill(win, qapp):
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QTableWidgetItem
    d = ws_dir("g04")
    a = make_doc(d / "A.docx", [("表1", [["项目", "金额"], ["A", "1"], ["合计", "1"]])])
    api.run_export(a)
    b = make_doc(d / "B.docx", [("表1", [["项目", "金额"], ["B", "2"], ["合计", "2"]])])
    batch = win._pages["BatchPage"]
    for p in (a, b):
        row = batch.sync_table.rowCount()
        batch.sync_table.insertRow(row)
        item = QTableWidgetItem(p.name)
        item.setData(Qt.UserRole, str(p))
        item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
        item.setCheckState(Qt.Checked)
        batch.sync_table.setItem(row, 0, item)
    batch._start_sync(False)
    wait_idle(qapp, win)
    time.sleep(0.5)
    qapp.processEvents()
    s0 = batch.sync_table.item(0, 2).text()
    s1 = batch.sync_table.item(1, 2).text()
    check("状态回填", s0.startswith("✔") and s1.startswith("✘"), f"{s0} / {s1}")


def test_g05_template_placeholder(win, qapp):
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QComboBox, QTableWidgetItem

    from wesync.docx_placeholders import VALUES_SHEET
    d = ws_dir("g05")
    tpl = make_doc(d / "模板.docx", [("表1", [["项目", "金额"], ["A", "1"], ["合计", "1"]])])
    doc = Document(str(tpl))
    doc.add_paragraph("被审计单位：{公司名称}。")
    doc.save(str(tpl))
    api.run_export(tpl)
    xlsx = api.workbook_path(tpl)
    batch = win._pages["BatchPage"]
    batch.mode_single.setChecked(True)
    batch.tpl_label.setText(str(tpl))
    row = batch.tpl_table.rowCount()
    batch.tpl_table.insertRow(row)
    item = QTableWidgetItem(xlsx.name)
    item.setData(Qt.UserRole, str(xlsx))
    batch.tpl_table.setItem(row, 0, item)
    combo = QComboBox()
    combo.addItems(["模板1"])
    batch.tpl_table.setCellWidget(row, 1, combo)
    batch.tpl_table.setItem(row, 2, QTableWidgetItem("待执行"))
    batch._fill_placeholder_sheets()
    check("填表任务完成", wait_idle(qapp, win))
    wb = load_workbook(str(xlsx))
    check("正文信息表已建", VALUES_SHEET in wb.sheetnames, str(wb.sheetnames))
    wb[VALUES_SHEET].cell(2, 2).value = "甲公司"
    wb.save(str(xlsx))
    batch.out_edit.setText(str(d / "out"))
    batch._start_tpl(False)
    check("生成任务完成", wait_idle(qapp, win))
    time.sleep(0.4)
    qapp.processEvents()
    out = d / "out" / "模板_模板_表格数据.docx"
    check("新Word已生成", out.exists(), str(list((d / 'out').glob('*')) if (d / 'out').exists() else []))
    texts = [p.text for p in Document(str(out)).paragraphs]
    check("正文占位符已替换", "被审计单位：甲公司。" in texts, str(texts))


def test_g06_diff_and_changelog(win, qapp):
    d = ws_dir("g06")
    docx = make_doc(d / "报告.docx", [("表1", [["项目", "金额"], ["A", "1"], ["合计", "1"]])])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, order, blocks = blocks_of(xlsx)
    wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 4).value = 888
    wb.save(str(xlsx))
    win.set_docx(docx)
    page = win._pages["WritebackPage"]
    before = docx.read_bytes()
    page._diff()
    check("差异表有行", page.diff_table.rowCount() > 0, str(page.diff_table.rowCount()))
    check("差异只读不写Word", docx.read_bytes() == before)
    page.changelog.setChecked(True)
    page._start(False)
    check("同步任务完成", wait_idle(qapp, win))
    time.sleep(0.4)
    qapp.processEvents()
    check("改动清单已导出", len(list(d.glob("*_改动清单_*.xlsx"))) == 1, str(list(d.glob('*'))))
    check("数据已同步", word_tables(docx)[0][1][1] == "888")


def test_g07_history_restore(win, qapp):
    d = ws_dir("g07")
    docx = make_doc(d / "报告.docx", [("表1", [["项目", "金额"], ["A", "1"], ["合计", "1"]])])
    api.run_export(docx)
    xlsx = api.workbook_path(docx)
    wb, order, blocks = blocks_of(xlsx)
    wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 4).value = 888
    wb.save(str(xlsx))
    api.run_writeback(docx, backup=True)
    check("同步后888", word_tables(docx)[0][1][1] == "888")
    win.set_docx(docx)
    hist = win._pages["HistoryPage"]
    hist._refresh()
    check("Word备份列出", hist.word_table.rowCount() >= 1, str(hist.word_table.rowCount()))
    check("恢复按钮启用", hist.word_restore.isEnabled())
    hist.word_table.selectRow(hist.word_table.rowCount() - 1)
    hist._restore("word")
    check("回滚后回到1", word_tables(docx)[0][1][1] == "1", str(word_tables(docx)))


def test_g08_marks_placeholder_tab(win, qapp):
    from wesync.docx_placeholders import VALUES_SHEET
    d = ws_dir("g08")
    docx = make_doc(d / "报告.docx", [("表1", [["项目", "金额"], ["A", "1"], ["合计", "1"]])])
    doc = Document(str(docx))
    doc.add_paragraph("被审计单位：{公司名称}，期间：{审计期间}。")
    doc.save(str(docx))
    win.set_docx(docx)
    page = win._pages["MarksPage"]
    check("页面改名", page.title == "标注与占位", page.title)
    page._export_placeholders()  # 工作簿不存在 → 确认后完整导出+骨架
    check("导出任务完成", wait_idle(qapp, win))
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    check("正文信息表已建", VALUES_SHEET in wb.sheetnames, str(wb.sheetnames))
    wb[VALUES_SHEET].cell(2, 2).value = "甲公司"
    wb.save(str(xlsx))
    page._sync_placeholders(False)
    check("同步任务完成", wait_idle(qapp, win))
    texts = [p.text for p in Document(str(docx)).paragraphs]
    check("占位符已替换", any("甲公司" in t for t in texts), str(texts))
    check("未填保留", any("{审计期间}" in t for t in texts), str(texts))


def test_g09_writeback_contents(win, qapp):
    from wesync.docx_placeholders import VALUES_SHEET
    d = ws_dir("g09")
    docx = make_doc(d / "报告.docx", [("表1", [["项目", "金额"], ["A", "1"], ["合计", "1"]])])
    doc = Document(str(docx))
    doc.add_paragraph("单位：{公司名称}。")
    doc.save(str(docx))
    api.run_export(docx)
    api.run_placeholder_export(docx)
    xlsx = api.workbook_path(docx)
    wb, order, blocks = blocks_of(xlsx)
    wb[SHEET_NAME].cell(blocks["T01"]["grid_rows"][1], 4).value = 999
    wb[VALUES_SHEET].cell(2, 2).value = "乙公司"
    wb.save(str(xlsx))
    win.set_docx(docx)
    page = win._pages["WritebackPage"]
    check("默认只勾表格", page.sync_tables.isChecked() and not page.sync_marks.isChecked()
          and not page.sync_placeholders.isChecked())
    check("保留高亮默认隐藏", page.wb_keep_highlight.isHidden())
    page.sync_marks.setChecked(True)
    qapp.processEvents()
    check("勾标注后显示保留高亮", not page.wb_keep_highlight.isHidden())
    page.sync_marks.setChecked(False)
    page.sync_placeholders.setChecked(True)
    sync_errors = []
    old_error = win.ctx.error
    win.ctx.error = lambda *a, **k: sync_errors.append(a)
    page._start(False)
    check("组合同步完成", wait_idle(qapp, win))
    check("组合同步无报错弹窗", not sync_errors, str(sync_errors))
    check("表格已同步", word_tables(docx)[0][1][1] == "999", str(word_tables(docx)))
    check("占位符已同步", any("乙公司" in p.text for p in Document(str(docx)).paragraphs))
    # 全不勾 → 报错不执行
    errors = []
    win.ctx.error = lambda *a, **k: errors.append(a)
    page.sync_tables.setChecked(False)
    page.sync_placeholders.setChecked(False)
    page._start(False)
    win.ctx.error = old_error
    check("空组合报错", errors and "至少勾选" in str(errors[0]), str(errors))


def test_g10_batch_contents(win, qapp):
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QTableWidgetItem

    from wesync.docx_placeholders import VALUES_SHEET
    d = ws_dir("g10")
    docx = make_doc(d / "报告.docx", [("表1", [["项目", "金额"], ["A", "1"], ["合计", "1"]])])
    doc = Document(str(docx))
    doc.add_paragraph("单位：{公司名称}。")
    doc.save(str(docx))
    api.run_export(docx)
    api.run_placeholder_export(docx)
    xlsx = api.workbook_path(docx)
    wb = load_workbook(str(xlsx))
    wb[VALUES_SHEET].cell(2, 2).value = "丙公司"
    wb.save(str(xlsx))
    batch = win._pages["BatchPage"]
    check("批量默认只勾表格", batch.batch_tables.isChecked() and not batch.batch_marks.isChecked()
          and not batch.batch_placeholders.isChecked())
    row = batch.sync_table.rowCount()
    batch.sync_table.insertRow(row)
    item = QTableWidgetItem(docx.name)
    item.setData(Qt.UserRole, str(docx))
    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
    item.setCheckState(Qt.Checked)
    batch.sync_table.setItem(row, 0, item)
    batch.batch_tables.setChecked(False)
    batch.batch_placeholders.setChecked(True)
    batch._start_sync(False)
    check("批量任务完成", wait_idle(qapp, win))
    time.sleep(0.4)
    qapp.processEvents()
    check("仅占位符已同步", any("丙公司" in p.text for p in Document(str(docx)).paragraphs))
    check("表格未动", word_tables(docx)[0][1][1] == "1", str(word_tables(docx)))

def test_g16_incremental_export_button(win, qapp):
    from PySide6.QtWidgets import QPushButton
    from wesync.config import SHEET_NEW_TABLES

    d = ws_dir("g16")
    docx = make_doc(d / "增量.docx", [("表1", [["项目", "金额"], ["A", "1"]])])
    api.run_export(docx)
    doc = Document(str(docx))
    doc.add_table(rows=2, cols=2).cell(0, 0).text = "新增项目"
    doc.save(str(docx))
    win.set_docx(docx)
    page = win._pages["ExportPage"]
    button = next(button for button in page.findChildren(QPushButton) if button.text() == "增量导出")
    check("增量导出按钮存在", button.isEnabled())
    button.click()
    check("GUI增量任务完成", wait_idle(qapp, win))
    wb = load_workbook(str(api.workbook_path(docx)))
    check("GUI增量表已登记", SHEET_NEW_TABLES in wb.sheetnames and wb[SHEET_NEW_TABLES].max_row > 1, str(wb.sheetnames))


def test_g11_runner_tolerates_dict_result(win, qapp):
    errors = []
    old_error = win.ctx.error
    win.ctx.error = lambda *a, **k: errors.append(a)
    try:
        win.ctx.run_steps([("返回摘要字典", lambda: {"tables": 1, "marks": 0})])
        check("字典返回值任务完成", wait_idle(qapp, win))
        check("字典返回值不误报失败", not errors, str(errors))
    finally:
        win.ctx.error = old_error


def test_g12_manual_and_icon(win):
    from wesync.gui.manual import load_manual_text
    from wesync.gui.resources import icon_path

    text = load_manual_text()
    check("内置说明书包含启动说明", "一、怎么启动" in text)
    check("内置说明书包含备份说明", "八、备份与回滚" in text)
    check("图标资源存在", icon_path().is_file(), str(icon_path()))
    check("窗口图标已设置", not win.windowIcon().isNull())
    help_page = win._pages["HelpPage"]
    check("说明书章节已生成", len(help_page.sections) >= 10, str(len(help_page.sections)))
    check("说明书目录已生成", help_page.chapter_list.count() == len(help_page.sections))
    check("数字步骤未误判为章节", any(title == "一、怎么启动" for title, _ in help_page.sections)
          and not any(title.startswith("1.") for title, _ in help_page.sections))


def test_g13_format_page_flow(win, qapp):
    d = ws_dir("g13")
    template = d / "样板.docx"
    doc = Document()
    title = doc.add_paragraph("审计报告")
    title.runs[0].font.size = 24
    doc.add_paragraph("目录")
    doc.add_paragraph("一、审计意见", style="Heading 1")
    doc.add_paragraph("样板正文格式。")
    doc.add_table(rows=2, cols=2).cell(0, 0).text = "项目"
    doc.save(str(template))

    target = make_doc(d / "乱版.docx", [("表1", [["项目", "金额"], ["A", "1"], ["合计", "1"]])])
    output = d / "乱版_格式调整.docx"
    before = target.read_bytes()
    win.set_docx(target)
    page = win._pages["FormatPage"]
    page.template_edit.setText(str(template))
    page.output_edit.setText(str(output))
    page._start_scan()
    check("格式扫描完成", wait_idle(qapp, win))
    page._check_done()
    check("预演后启用生成按钮", page.apply_button.isEnabled())
    check("预演不写文件", not output.exists() and target.read_bytes() == before)
    check("预演显示所选输出路径", str(output) in page.details.toPlainText(), page.details.toPlainText())
    page._start_apply()
    check("格式生成完成", wait_idle(qapp, win))
    page._check_done()
    check("格式副本已生成", output.exists())
    check("原报告未改动", target.read_bytes() == before)


def test_g14_saved_format_profile_gui_flow(win, qapp, monkeypatch):
    from PySide6.QtWidgets import QInputDialog
    from wesync.errors import SyncError

    errors = []
    monkeypatch.setattr(win.ctx, "error", lambda *args, **kwargs: errors.append(args))
    d = ws_dir("g14")
    template = d / "样板.docx"
    doc = Document()
    title = doc.add_paragraph("审计报告")
    title.runs[0].font.size = 24
    doc.add_paragraph("目录")
    doc.add_paragraph("一、审计意见", style="Heading 1")
    doc.add_paragraph("只保存格式，不保存这段正文。")
    doc.add_table(rows=2, cols=2).cell(0, 0).text = "项目"
    doc.save(str(template))
    target = make_doc(d / "乱版.docx", [("表1", [["项目", "金额"], ["A", "1"], ["合计", "1"]])])
    output = d / "乱版_方案格式调整.docx"
    page = win._pages["FormatPage"]
    monkeypatch.setattr(QInputDialog, "getText", lambda *args, **kwargs: ("GUI方案", True))
    win.set_docx(target)
    page.on_show()
    page.template_edit.setText(str(template))
    page._save_profile()
    check("保存方案完成", wait_idle(qapp, win))
    page._check_done()
    check("保存后选中新方案", page.profile_combo.currentData() == "GUI方案", page.profile_combo.currentData())
    check("选用方案后样板输入禁用", not page.template_edit.isEnabled())
    check("保存后需重新预演", not page.apply_button.isEnabled())

    template.unlink()
    page.template_edit.setText(str(template))
    page.output_edit.setText(str(output))
    check("输出路径保持所选值", page.output_edit.text() == str(output), page.output_edit.text())
    page._start_scan()
    check("方案预演完成", wait_idle(qapp, win))
    page._check_done()
    check("预演来源为已存方案", "格式来源：已存方案：GUI方案" in page.details.toPlainText(), page.details.toPlainText())
    check("生成路径来自输入框", page._paths()[2] == output, str(page._paths()))
    page._start_apply()
    check("已启动方案生成", page._waiting == "apply", f"scan={page._scan_key}；result={page._result}；errors={errors}")
    check("方案生成完成", wait_idle(qapp, win))
    page._check_done()
    check("生成回调已完成", page._waiting is None and not win.ctx.busy,
          f"waiting={page._waiting}；busy={win.ctx.busy}；result={page._result}；status={page.status_label.text()}；errors={errors}")
    check("无样板仍生成副本", output.exists(), f"status={page.status_label.text()}；result={page._result}；errors={errors}")
    check("生成后需重新预演", not page.apply_button.isEnabled())

    failed_output = d / "乱版_生成失败.docx"
    page.output_edit.setText(str(failed_output))
    page._start_scan()
    check("失败用例预演完成", wait_idle(qapp, win))
    page._check_done()
    def fail_run(*args, **kwargs):
        raise SyncError("模拟生成失败")

    monkeypatch.setattr(api, "run_report_format", fail_run)
    page._start_apply()
    check("失败用例任务结束", wait_idle(qapp, win))
    page._check_done()
    check("生成失败不误报成功", page.status_label.text().startswith("执行失败") and not failed_output.exists(), page.status_label.text())

    confirmations = []
    monkeypatch.setattr(win.ctx, "confirm", lambda *args: confirmations.append(args) or True)
    page._delete_profile()
    check("删除前有确认", len(confirmations) == 1, str(confirmations))
    check("方案已从下拉删除", page.profile_combo.findData("GUI方案") < 0)
    check("删除后回到样板模式", page.profile_combo.currentData() in (None, ""))


def test_g15_format_scope_tab_invalidates_preview_and_limits_apply(win, qapp, monkeypatch):
    from docx.shared import Pt
    from wesync.report_format import DEFAULT_SCOPE

    errors = []
    monkeypatch.setattr(win.ctx, "error", lambda *args, **kwargs: errors.append(args))
    d = ws_dir("g15")
    template_doc = Document()
    template_doc.add_paragraph("审计报告")
    template_doc.add_paragraph("目录")
    template_doc.add_paragraph("一、审计意见", style="Heading 1")
    template_table = template_doc.add_table(rows=2, cols=2)
    template_table.cell(0, 0).text = "项目"
    template_table.cell(1, 0).text = "A"
    template_table.cell(1, 1).text = "100"
    for row in template_table.rows:
        for cell in row.cells:
            for run in cell.paragraphs[0].runs:
                run.font.size = Pt(8)
    template_body = template_doc.add_paragraph("样板正文")
    template_body.runs[0].font.size = Pt(14)
    template = d / "样板.docx"
    template_doc.save(str(template))

    target_doc = Document()
    target_doc.add_paragraph("乱版报告")
    target_doc.add_paragraph("目录")
    target_doc.add_paragraph("一、审计意见", style="Heading 1")
    target_table = target_doc.add_table(rows=2, cols=2)
    target_table.cell(0, 0).text = "项目"
    target_table.cell(1, 0).text = "A"
    target_table.cell(1, 1).text = "100"
    target_body = target_doc.add_paragraph("未选择时保持原样的正文")
    target_body.runs[0].font.size = Pt(30)
    target = d / "乱版.docx"
    target_doc.save(str(target))

    page = win._pages["FormatPage"]
    for button in page.scope_buttons.values():
        button.setChecked(True)
    check("格式页恰有两个Tab", page.tabs.count() == 2, str(page.tabs.count()))
    check("第二个Tab为调整项目", page.tabs.tabText(1) == "调整项目", page.tabs.tabText(1))
    check("七项默认全部勾选", set(page.scope_buttons) == set(DEFAULT_SCOPE)
          and all(button.isChecked() for button in page.scope_buttons.values()), str(page._scope()))
    check("表格框线可单独选择", "table_borders" in page.scope_buttons)

    win._select_page("FormatPage")
    win.set_docx(target)
    check("顶部选文件后当前格式页立即刷新", page.target_label.text() == str(target), page.target_label.text())
    check("顶部选文件后自动带出副本路径", page.output_edit.text() == str(target.with_name("乱版_格式调整.docx")), page.output_edit.text())
    page.template_edit.setText(str(template))
    page._start_scan()
    check("默认范围预演完成", wait_idle(qapp, win))
    page._check_done()
    check("默认范围可生成", page.apply_button.isEnabled())
    for key in ("cover", "headings", "statement", "body"):
        page.scope_buttons[key].setChecked(False)
    check("改变范围使旧预演失效", not page.apply_button.isEnabled() and page._result is None)
    page._start_apply()
    check("旧预演不能直接生成", any(args[0] == "尚未预演" for args in errors), str(errors))

    output = d / "仅表格.docx"
    page.output_edit.setText(str(output))
    page._start_scan()
    check("仅表格预演完成", wait_idle(qapp, win))
    page._check_done()
    details = page.details.toPlainText()
    check("预演显示未选正文", "将处理 0 个非空段落、1 张表" in details and "未选项目跳过" in details, details)
    page._start_apply()
    check("仅表格生成完成", wait_idle(qapp, win))
    page._check_done()
    formatted = Document(str(output))
    formatted_body = next(p for p in formatted.paragraphs if "未选择时保持原样" in p.text)
    check("取消正文后正文保持原样", formatted_body.runs[0].font.size == Pt(30))
    check("仍调整所选表格", formatted.tables[0].cell(1, 1).paragraphs[0].runs[0].font.size == Pt(8))
    check("生成后显示 Word 目视验收清单", "Word 目视验收" in page.details.toPlainText()
          and "宽表格" in page.details.toPlainText(), page.details.toPlainText())

    page.scope_buttons["body"].setChecked(True)
    check("恢复正文后需重新预演", not page.apply_button.isEnabled())
    second_output = d / "正文与表格.docx"
    page.output_edit.setText(str(second_output))
    page._start_scan()
    check("恢复正文后预演完成", wait_idle(qapp, win))
    page._check_done()
    page._start_apply()
    check("恢复正文后生成完成", wait_idle(qapp, win))
    page._check_done()
    formatted = Document(str(second_output))
    formatted_body = next(p for p in formatted.paragraphs if "未选择时保持原样" in p.text)
    check("恢复正文后应用样板正文格式", formatted_body.runs[0].font.size == Pt(14))

    page._start_scan()
    page.scope_buttons["body"].setChecked(False)
    check("预演中改变选择任务已启动", wait_idle(qapp, win))
    page._check_done()
    check("预演中改变选择不启用旧结果", not page.apply_button.isEnabled()
          and page.status_label.text().startswith("格式来源或调整项目已变化"), page.status_label.text())
    for button in page.scope_buttons.values():
        button.setChecked(True)
