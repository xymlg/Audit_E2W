"""导出时给合计行写求和公式：改明细后合计自动更新。

口径：首列以「合计」开头的行视为合计行；对每个数值列写 =SUM(数据区首行:数据区末行)，
数据区 = 表头之后、该合计行之前。表头按 w:tblHeader 标记识别，未标记时按首行。
文字列（如备注）与「—」等占位文字不写公式，保留原样。
"""

from __future__ import annotations

from openpyxl.utils import get_column_letter

from .valuefmt import to_excel_value

TOTAL_PREFIX = "合计"


def _is_number(text: str) -> bool:
    value, _ = to_excel_value(text or "")
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def total_formula_cells(info: dict, grid_start: int, col0: int = 2) -> dict[tuple[int, int], str]:
    """返回合计行应写 SUM 公式的单元格 {(Excel行, Excel列): 公式}。

    grid_start 为该表网格首行（表头）在工作表中的行号；col0 为网格首列列号。
    仅对合计行里本身是数字或空、且数据区出现过数字的列写公式；合并掉的非 owner 格、文字占位格跳过。
    """
    text = info["text"]
    nrows, ncols = info["nrows"], info["ncols"]
    cellmap = info["cellmap"]
    header = info.get("header_rows") or set()
    data_start = (max(header) + 1) if header else 1
    formulas: dict[tuple[int, int], str] = {}
    for t in range(nrows):
        if not (text[t][0] or "").strip().startswith(TOTAL_PREFIX):
            continue
        data_end = t - 1
        if data_end < data_start:
            continue
        for c in range(1, ncols):
            pos = cellmap[t][c]
            if pos is None or not pos.owner:
                continue
            total_value, _ = to_excel_value(text[t][c] or "")
            if isinstance(total_value, str):  # 「—」等占位文字，保留原样
                continue
            if not any(_is_number(text[r][c]) for r in range(data_start, data_end + 1)):
                continue
            letter = get_column_letter(col0 + c)
            formulas[(grid_start + t, col0 + c)] = (
                f"=SUM({letter}{grid_start + data_start}:{letter}{grid_start + data_end})"
            )
    return formulas