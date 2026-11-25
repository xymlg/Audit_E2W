#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""导出表格到 Excel：把 Word 里的全部表格导出为工作簿，并创建/归一化书签。

用法：
    .venv\\Scripts\\python.exe "1-导出表格到Excel.py"                 # 自动识别同目录唯一的 docx
    .venv\\Scripts\\python.exe "1-导出表格到Excel.py" "报告.docx"      # 指定文件
    .venv\\Scripts\\python.exe "1-导出表格到Excel.py" --dry-run        # 只看将导出什么

注意：导出会按当前 Word 重新生成工作簿（覆盖前自动备份到 _backup）。
"""

import sys
from pathlib import Path

# 打包成 exe 后 __file__ 指向临时解包目录，这里改用 exe 自身所在目录，
# 保证“自动识别同目录 docx”在免安装版里同样有效。
FROZEN = getattr(sys, "frozen", False)
SCRIPT_DIR = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent
if not FROZEN:
    sys.path.insert(0, str(SCRIPT_DIR))

from wesync.fsops import default_base_dir  # noqa: E402
from wesync.cli import main  # noqa: E402

BASE_DIR = default_base_dir(SCRIPT_DIR)

if __name__ == "__main__":
    raise SystemExit(main(["export", *sys.argv[1:]], base_dir=BASE_DIR))
