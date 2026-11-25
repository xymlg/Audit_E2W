#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""同步标注：把「链接标注」工作表里填写的「新内容／链接地址」写回 Word。

用法：
    .venv\\Scripts\\python.exe "6-同步标注.py" --dry-run   # 预演
    .venv\\Scripts\\python.exe "6-同步标注.py"             # 正式同步（先自动备份）
    .venv\\Scripts\\python.exe "6-同步标注.py" "报告.docx"  # 指定文件

同步后该处高亮会被去掉（黄色=待处理，无色=已处理）；Word 中已删除的标记会跳过并记录。
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
    raise SystemExit(main(["marks-sync", *sys.argv[1:]], base_dir=BASE_DIR))
