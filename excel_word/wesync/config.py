"""全局常量与正则：工作表名、书签前缀、字体与配色、边框/对齐映射、各种上限。"""

import re

from openpyxl.styles import PatternFill


SHEET_NAME = "表格汇总"


CODE_RE = re.compile(r"^T\d{2,}$")


DATE_RE = re.compile(r"^\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日$")


NUMBER_RE = re.compile(r"^-?\d[\d,]*(\.\d+)?$")


EXCEL_LINK_TEXT = "〔→Excel〕"


BOOKMARK_PREFIX = "Table_"


# 新表标记：表块 C 列写「书签：新增」表示要在 Word 里新建这张表
NEW_BOOKMARK = "新增"

# 可逆删除标记：表块 C 列写「书签：已删除」表示 Word 里的表已删、表块保留待恢复
DELETED_BOOKMARK = "已删除"


LEGACY_BOOKMARK_RE = re.compile(r"^TBL_\d+$")


META_SCAN_COLUMNS = 16


# B 列行标记：表头行写「表头」；数据行除首列外任一列有内容=✓、全空=✕；
# 勾选「删除✕行」同步后，被删行在 Excel 保留为「已删」（改回✓可恢复）
MARK_HEADER = "表头"
MARK_YES = "✓"
MARK_NO = "✕"
MARK_DELETED = "已删"
MARK_VALUES = (MARK_HEADER, MARK_YES, MARK_NO, MARK_DELETED)


MAX_COLUMN_SHIFT = 4


SKIP_COLUMN = -1


MAX_LOG_LINES = 200


FONT_NAME = "宋体"


FONT_SIZE = 9


FILL_CODE_COLUMN = PatternFill("solid", fgColor="F2F2F2")


FILL_ANCHOR = PatternFill("solid", fgColor="FFF2CC")


BORDER_MAP = {
    "single": "thin",
    "double": "double",
    "dotted": "dotted",
    "dashed": "dashed",
    "dashDot": "dashDot",
    "dashDotDot": "dashDotDot",
    "thick": "medium",
    "wave": "thin",
}


JC_MAP = {"center": "center", "right": "right", "both": "justify", "distribute": "justify", "left": "left"}


VA_MAP = {"center": "center", "bottom": "bottom", "top": "top"}


CELL_REF_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


RANGE_REF_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)\s*:\s*\$?([A-Za-z]{1,3})\$?(\d+)$")


SUM_RE = re.compile(r"^SUM\s*\((.*)\)$", re.IGNORECASE)


ARITH_OK_RE = re.compile(r"^[0-9+\-*/(). ]+$")
