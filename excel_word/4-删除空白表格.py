#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""删除空白表格：找出整表全空或仅表头有字的表格，按候选编号删除，可连表标题一起删。

用法：
    .venv\\Scripts\\python.exe "4-删除空白表格.py"                  # 只扫描列出候选，不删除
    .venv\\Scripts\\python.exe "4-删除空白表格.py" --pick 1,3       # 删除候选 1、3（连表标题）
    .venv\\Scripts\\python.exe "4-删除空白表格.py" --all            # 删除全部候选
    .venv\\Scripts\\python.exe "4-删除空白表格.py" --all --keep-caption   # 只删表格，保留标题
    .venv\\Scripts\\python.exe "4-删除空白表格.py" --pick 1 --reversible  # 可逆删除：Excel 表块保留待恢复
    .venv\\Scripts\\python.exe "4-删除空白表格.py" --restore 1,3    # 恢复已可逆删除的表（all=全部）
    .venv\\Scripts\\python.exe "4-删除空白表格.py" "报告.docx" --dry-run  # 指定文件，只预演
"""

import sys
from pathlib import Path

# 打包成 exe 后 __file__ 指向临时解包目录，改用 exe 自身所在目录。
FROZEN = getattr(sys, "frozen", False)
SCRIPT_DIR = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent
if not FROZEN:
    sys.path.insert(0, str(SCRIPT_DIR))

from wesync.fsops import default_base_dir  # noqa: E402
from wesync.cli import main  # noqa: E402

BASE_DIR = default_base_dir(SCRIPT_DIR)

if __name__ == "__main__":
    raise SystemExit(main(["blank", *sys.argv[1:]], base_dir=BASE_DIR))