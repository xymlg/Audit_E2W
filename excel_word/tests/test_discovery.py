"""启动器页面自动发现、坏包隔离、_目录过滤。"""
import importlib.util
import os
import shutil
import sys
import time

from helpers import ROOT, api, check, make_doc, ws_dir


def test_discovery(qapp):
    good = ROOT / "zz_test_feature"
    bad = ROOT / "zz_bad_feature"
    good.mkdir(exist_ok=True)
    bad.mkdir(exist_ok=True)
    (good / "gui_pages.py").write_text(
        "from PySide6.QtWidgets import QLabel\n"
        "from wesync.gui.pages import BasePage, register_page\n"
        "class DummyPage(BasePage):\n"
        "    title = '测试功能'\n"
        "    def build(self, body):\n"
        "        self.card(body).layout().addWidget(QLabel('ok'))\n"
        "register_page(DummyPage)\n", encoding="utf-8")
    (bad / "gui_pages.py").write_text("raise RuntimeError('故意坏掉')\n", encoding="utf-8")
    import wesync.gui.pages as pages_mod
    try:
        spec = importlib.util.spec_from_file_location("toolbox_launcher", ROOT / "审计工具箱.pyw")
        launcher = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(launcher)
        names = [c.__name__ for c in pages_mod.iter_pages()]
        check("自动发现新页面", "DummyPage" in names, str(names))
        own_dir = str(ROOT / "excel_word" / "tests")
        bad_paths = [p for p in sys.path if p.startswith(str(ROOT) + os.sep + "_") and p != own_dir]
        check("_目录不进搜索路径", not bad_paths, str(bad_paths))
        from wesync.gui.app import WesyncApp
        win = WesyncApp(base_dir=ROOT / "excel_word")
        win.ctx.info = lambda *a, **k: None
        check("导航含测试功能", "DummyPage" in win._pages, str(list(win._pages)))
        d = ws_dir("f03")
        docx = make_doc(d / "f.docx", [("表1 空", [["项目", "本期"], ["", ""], ["", ""]])])
        api.run_export(docx)
        win.set_docx(docx)
        blank = win._pages["BlankPage"]
        blank._scan()
        check("扫描后删除按钮启用", blank.btn_delete.isEnabled())
        win.ctx.run_steps([("空任务", lambda: 0)])
        end = time.time() + 30
        while time.time() < end and (win.runner.busy or not win.runner.queue.empty()):
            qapp.processEvents()
            time.sleep(0.02)
        qapp.processEvents()
        check("任务后仍启用", blank.btn_delete.isEnabled())
        win.close()
    finally:
        pages_mod._PAGES[:] = [c for c in pages_mod._PAGES if c.__name__ != "DummyPage"]
        shutil.rmtree(good, ignore_errors=True)
        shutil.rmtree(bad, ignore_errors=True)