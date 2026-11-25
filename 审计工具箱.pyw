#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""审计工具箱 · 全局界面入口（双击运行，.pyw 不弹黑窗）。

启动后左侧导航直接列出各功能：
  Word表格同步（导出 / 回写 / 链接标注 / 书签）——位于 excel_word 子功能；
  后续新增功能（如多文件合并）会自动出现在导航中。

也可命令行启动：
    ../.venv/Scripts/python.exe 审计工具箱.pyw
    ../.venv/Scripts/python.exe -c "import sys; sys.path.insert(0, r'excel_word'); from wesync.gui import run; run()"
"""

import importlib.util
import sys
from pathlib import Path

# 打包成 exe 后 __file__ 指向临时解包目录，改用 exe 自身所在目录。
FROZEN = getattr(sys, "frozen", False)
SCRIPT_DIR = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent

# 工具箱下各子功能代码以子目录形式存在（excel_word、file_merge …），
# 把各功能目录加入 sys.path，全局界面即可 import 到各功能包（如 wesync）。
# 跳过隐藏目录、_ 开头的工作目录（_backup/_scratch 等）与 venv/build/dist 等非功能目录。
FEATURE_DIRS = []
for sub in sorted(SCRIPT_DIR.iterdir()):
    if not sub.is_dir():
        continue
    name = sub.name.lower()
    if name.startswith((".", "_")) or name in {"venv", "build", "dist", "packaging", "__pycache__"}:
        continue
    FEATURE_DIRS.append(sub)
    sys.path.insert(0, str(sub))

from wesync.fsops import default_base_dir  # noqa: E402
from wesync.gui import run  # noqa: E402

# 页面自动发现：各功能目录放一个 gui_pages.py，里面继承 BasePage 并 register_page()，
# 启动时自动加载，新功能无需改启动器或 wesync 代码（见《功能规划与文件布局.md》第六节）。
for sub in FEATURE_DIRS:
    hook = sub / "gui_pages.py"
    if not hook.is_file():
        continue
    try:
        spec = importlib.util.spec_from_file_location(f"_toolbox_pages_{sub.name}", hook)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
    except Exception as exc:
        print(f"警告：功能 {sub.name} 的页面加载失败，已跳过：{exc}", file=sys.stderr)

if __name__ == "__main__":
    # frozen 时 __file__ 在 _internal 里，不能靠包内路径推工作目录；与命令行入口同口径：
    # 当前目录有 docx 用当前目录，否则用入口（exe）所在目录。
    raise SystemExit(run(base_dir=default_base_dir(SCRIPT_DIR)))
