"""支持 ``python -m wesync <子命令> "<Word文件.docx>"``。"""

from .cli import main

if __name__ == "__main__":
    raise SystemExit(main())
