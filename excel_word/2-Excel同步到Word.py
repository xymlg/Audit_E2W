#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Excel 同步到 Word：把工作簿里的改动（含增行删行、增列删列）回写到 Word。

用法：
    .venv\\Scripts\\python.exe "2-Excel同步到Word.py" --dry-run   # 预演，只显示要改哪些单元格
    .venv\\Scripts\\python.exe "2-Excel同步到Word.py"             # 正式回写（先自动备份）
    .venv\\Scripts\\python.exe "2-Excel同步到Word.py" "报告.docx"  # 指定文件

回写前请先关闭该 Word 文件与 Excel 工作簿。
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
    raise SystemExit(main(["writeback", *sys.argv[1:]], base_dir=BASE_DIR))
