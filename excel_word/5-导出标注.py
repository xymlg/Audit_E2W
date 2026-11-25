#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""导出标注：把 Word 里的高亮标记导出/刷新到工作簿的「链接标注」工作表。

用法：
    .venv\\Scripts\\python.exe "5-导出标注.py"                 # 自动识别同目录唯一的 docx
    .venv\\Scripts\\python.exe "5-导出标注.py" --dry-run        # 只看会导出哪些标记
    .venv\\Scripts\\python.exe "5-导出标注.py" "报告.docx"      # 指定文件

只给标记插入不可见书签，不改 Word 文字；已填写的新内容/链接地址会被保留。
"""

import sys
from pathlib import Path

FROZEN = getattr(sys, "frozen", False)
SCRIPT_DIR = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent
if not FROZEN:
    sys.path.insert(0, str(SCRIPT_DIR))

from wesync.fsops import default_base_dir  # noqa: E402
from wesync.cli import main  # noqa: E402

BASE_DIR = default_base_dir(SCRIPT_DIR)

if __name__ == "__main__":
    raise SystemExit(main(["marks", *sys.argv[1:]], base_dir=BASE_DIR))
