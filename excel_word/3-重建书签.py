#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""重建书签：只处理 Word 书签与正文遗留文字，并同步 Excel 里的书签名与跳转链接，不动数据。

用法：
    .venv\\Scripts\\python.exe "3-重建书签.py"                  # 自动识别同目录唯一的 docx
    .venv\\Scripts\\python.exe "3-重建书签.py" --dry-run         # 只看将处理什么
    .venv\\Scripts\\python.exe "3-重建书签.py" "报告.docx"       # 指定文件
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
    raise SystemExit(main(["bookmark", *sys.argv[1:]], base_dir=BASE_DIR))
