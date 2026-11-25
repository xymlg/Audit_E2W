"""Word 表格 ⇄ Excel 双向同步工具（结构化包）。

模块职责：
    config / errors / console   —— 常量、异常、中文日志
    fsops / valuefmt            —— 文件操作、数值与公式处理
    docx_*                      —— Word 侧：网格解析、书签、单元格写入、行列结构
    excel_*                     —— Excel 侧：导出工作簿、解析表块与元数据
    commands / cli              —— 各子命令与命令行入口

对外入口：``from wesync import main`` 或 ``python -m wesync.cli <子命令> ...``
"""

from __future__ import annotations

__version__ = "1.11.1"

from .cli import main

__all__ = ["main", "__version__"]
