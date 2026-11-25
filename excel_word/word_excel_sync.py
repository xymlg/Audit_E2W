#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""兼容入口：等价于 ``python -m wesync.cli``，子命令随 wesync.cli 自动更新
（export / rejudge / bookmark / writeback / marks / marks-sync / placeholders /
placeholders-sync / blank / clear-bookmarks）。

核心代码已拆分到 wesync 包；本文件保留是为了让既有命令继续可用。
注意：本入口**不启用**“自动识别同目录 docx”，必须显式给出文件，避免误触 export 覆盖工作簿。
如果不想每次输入文件名，请直接使用中文入口脚本：1-导出表格到Excel.py / 2-Excel同步到Word.py / 3-重建书签.py。
"""

import sys
from pathlib import Path

# 打包成 exe 后 __file__ 指向临时解包目录，改用 exe 自身所在目录。
FROZEN = getattr(sys, "frozen", False)
SCRIPT_DIR = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent
if not FROZEN:
    sys.path.insert(0, str(SCRIPT_DIR))

from wesync.cli import main  # noqa: E402  （放在 sys.path 调整之后）

if __name__ == "__main__":
    raise SystemExit(main())
