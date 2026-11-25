"""Excel 导出层：把 Word 表格生成为单工作表堆叠结构（样式、合并、跳转链接、定义名称）。"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.hyperlink import Hyperlink
from .config import FILL_ANCHOR, FILL_CODE_COLUMN, FONT_NAME, FONT_SIZE, MARK_HEADER, MARK_NO, MARK_YES, SHEET_NAME
from .console import log
from .docx_grid import cell_alignment, explicit_borders, position_sides
from .excel_totals import total_formula_cells
from .valuefmt import approx_width, display_text, to_excel_value


def make_border(sides: dict) -> Border:
    def side(name):
        style = sides.get(name)
        return Side(style=style) if style else Side()

    return Border(left=side("left"), right=side("right"), top=side("top"), bottom=side("bottom"))


def export_workbook(infos: list[dict], docx_name: str, with_links: bool = True, totals_formula: bool = True) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.column_dimensions["A"].width = 7
    ws.freeze_panes = "A1"

    widths: dict[int, float] = {}
    row = 1
    defined: list[tuple[str, str]] = []
    formula_count = 0

    for info in infos:
        anchor_row = row
        code = info["code"]

        ws.cell(anchor_row, 1, code)
        ws.cell(anchor_row, 3, f"【表{info['index'] + 1}】{info['name']} · 章节：{info['section']}")
        meta_text = f"书签：{info['bookmark']}"
        if info.get("old_bookmark"):
            meta_text += f"（原审计软件书签：{info['old_bookmark']}）"
        ws.cell(anchor_row, 4, meta_text)
        link_cell = ws.cell(anchor_row, 5, "→定位Word表")
        if with_links:
            link_cell.hyperlink = Hyperlink(
                ref=link_cell.coordinate, target=docx_name, location=info["bookmark"], display="→定位Word表"
            )
        ws.cell(anchor_row, 6, f"列数：{info['ncols']}")
        for col in range(1, 7):
            cell = ws.cell(anchor_row, col)
            cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True, color="1F4E79")
            cell.fill = FILL_ANCHOR
            cell.alignment = Alignment(horizontal="left" if col > 1 else "center", vertical="center")

        grid_start = anchor_row + 1
        nrows, ncols = info["nrows"], info["ncols"]
        borders = info["borders"]
        sides_cache: dict[tuple[int, int], dict] = {}
        formulas = total_formula_cells(info, grid_start, col0=3) if totals_formula else {}
        formula_count += len(formulas)

        for r in range(nrows):
            xlsx_row = grid_start + r
            code_cell = ws.cell(xlsx_row, 1, code)
            code_cell.fill = FILL_CODE_COLUMN
            code_cell.font = Font(name=FONT_NAME, size=FONT_SIZE, color="808080")
            code_cell.alignment = Alignment(horizontal="center", vertical="center")
            if r == 0:
                mark, mark_color = MARK_HEADER, "808080"
            else:
                has_content = any((info["text"][r][c] or "").strip() for c in range(1, ncols)) or any(
                    (xlsx_row, 3 + c) in formulas for c in range(1, ncols)
                )
                mark, mark_color = (MARK_YES, "008000") if has_content else (MARK_NO, "C00000")
            mark_cell = ws.cell(xlsx_row, 2, mark)
            mark_cell.font = Font(name=FONT_NAME, size=FONT_SIZE, color=mark_color)
            mark_cell.alignment = Alignment(horizontal="center", vertical="center")
            for c in range(ncols):
                pos = info["cellmap"][r][c]
                sides_cache[(r, c)] = position_sides(borders, r, c, nrows, ncols)
                if pos is None or not pos.owner:
                    continue
                cell = ws.cell(xlsx_row, 3 + c)
                value, number_format = to_excel_value(info["text"][r][c] or "")
                horizontal, vertical, _ = cell_alignment(pos.tc)
                formula = formulas.get((xlsx_row, 3 + c))
                if formula is not None:
                    cell.value = formula
                elif value is not None:
                    cell.value = value
                cell.number_format = number_format or ("@" if isinstance(value, str) else "General")
                cell.alignment = Alignment(
                    horizontal=horizontal, vertical=vertical, wrap_text=True if isinstance(value, str) and "\n" in value else False
                )
                cell.font = Font(name=FONT_NAME, size=FONT_SIZE)
                sides = dict(sides_cache[(r, c)])
                sides.update(explicit_borders(pos.tc))
                cell.border = make_border(sides)
                text_for_width = display_text(value, number_format) if value is not None else ""
                longest = max((approx_width(part) for part in text_for_width.split("\n")), default=0.0)
                widths[3 + c] = max(widths.get(3 + c, 8.0), min(longest + 2.0, 45.0))

        for (r, c, rowspan, colspan) in info["merges"]:
            if rowspan == 1 and colspan == 1:
                continue
            ws.merge_cells(
                start_row=grid_start + r,
                start_column=3 + c,
                end_row=grid_start + r + rowspan - 1,
                end_column=3 + c + colspan - 1,
            )
            top = sides_cache[(r, c)]["top"]
            left = sides_cache[(r, c)]["left"]
            bottom = sides_cache[(r + rowspan - 1, c + colspan - 1)]["bottom"]
            right = sides_cache[(r + rowspan - 1, c + colspan - 1)]["right"]
            anchor_cell = ws.cell(grid_start + r, 3 + c)
            anchor_cell.border = make_border({"top": top, "left": left, "bottom": bottom, "right": right})

        defined.append((f"EX_{code}", f"'{SHEET_NAME}'!$C${grid_start}"))
        row = grid_start + nrows + 1  # 表间空一行

    ws.column_dimensions["B"].width = 6
    for column, width in widths.items():
        ws.column_dimensions[get_column_letter(column)].width = width

    for name, ref in defined:
        definition = DefinedName(name, attr_text=ref)
        try:
            wb.defined_names.add(definition)
        except AttributeError:  # 兼容旧版 openpyxl
            wb.defined_names[name] = definition
    if formula_count:
        log(f"[导出] 合计行已写 {formula_count} 个求和公式（在 Excel 改明细后合计自动更新；同步时按公式结果回写 Word）")
    return wb
