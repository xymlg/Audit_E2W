"""Excel 解析层：识别表块与锚点元数据、刷新书签名/链接/列数、刷新 EX_ 定义名称。"""

import re
from pathlib import Path
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.hyperlink import Hyperlink
from .config import CODE_RE, DELETED_BOOKMARK, FILL_ANCHOR, FILL_CODE_COLUMN, FONT_NAME, FONT_SIZE, MARK_DELETED, MARK_HEADER, MARK_NO, MARK_VALUES, MARK_YES, META_SCAN_COLUMNS, NEW_BOOKMARK, SHEET_NAME
from .console import log, warn
from .fsops import backup_file, load_workbook_for_save


def _cell_str(ws, row: int, col: int) -> str:
    value = ws.cell(row, col).value
    if value is None:
        return ""
    return str(value).strip()


def _merged_ranges(ws) -> list[tuple[int, int, int, int]]:
    return [(r.min_row, r.min_col, r.max_row, r.max_col) for r in ws.merged_cells.ranges]


def _in_merged(ranges, row: int, col: int) -> bool:
    return any(r1 <= row <= r2 and c1 <= col <= c2 for (r1, c1, r2, c2) in ranges)


def row_grid_has_content(ws, row: int, max_col: int, ranges, start_col: int = 2) -> bool:
    for col in range(start_col, max(max_col, start_col) + 1):
        cell = ws.cell(row, col)
        if cell.value is not None and str(cell.value).strip() != "":
            return True
        if _in_merged(ranges, row, col):
            return True
    return False


def anchor_metadata(ws, row: int, max_col: int) -> dict:
    """按内容特征定位锚点行的元数据，兼容插入列后整体右移的情况。"""
    meta: dict = {"name": "", "bookmark": "", "ncols": None, "link_col": None, "after": "", "restore_after": "", "caption": "", "sort": "", "cols": {}}
    for col in range(1, min(max_col, META_SCAN_COLUMNS) + 1):
        cell = ws.cell(row, col)
        text = "" if cell.value is None else str(cell.value).strip()
        if col > 1 and cell.hyperlink is not None and meta["link_col"] is None:
            meta["link_col"] = col
            meta["cols"]["link"] = col
        if text.startswith("书签："):
            meta["cols"]["bookmark"] = col
            meta["bookmark"] = text.split("：", 1)[1].split("（", 1)[0].strip()
        elif col >= 2 and text in (NEW_BOOKMARK, "新表"):
            meta["cols"]["bookmark"] = col
            meta["bookmark"] = NEW_BOOKMARK
        elif text.startswith("列数："):
            meta["cols"]["ncols"] = col
            match = re.search(r"(\d+)", text)
            if match:
                meta["ncols"] = int(match.group(1))
        elif col >= 2 and text.startswith("位置："):
            meta["cols"]["pos"] = col
            value = text.split("：", 1)[1].strip()
            match = re.match(r"T\d{2,}", value)
            meta["after"] = match.group(0) if match else "末尾"
        elif col >= 2 and text.startswith("原位置："):
            meta["cols"]["restore_pos"] = col
            value = text.split("：", 1)[1].strip()
            meta["restore_after"] = value[:-2] if value.endswith("之后") else "末尾"
        elif col >= 2 and text.startswith("原标题："):
            meta["cols"]["caption"] = col
            meta["caption"] = text.split("：", 1)[1].strip()
        elif col >= 2 and text.startswith("排序："):
            meta["cols"]["sort"] = col
            meta["sort"] = text
        elif col >= 2 and not meta["name"] and ("【表" in text or text.startswith("【新表】")):
            meta["cols"]["name"] = col
            meta["name"] = text
    return meta


def parse_blocks(ws) -> tuple[list[str], dict[str, dict]]:
    """解析表块。

    块归属规则：A 列为表编号的行、以及 A 列为空但网格区有内容的行（例如用户在块中间
    插入行后直接填数据）都计入当前表块；整行空白视为分隔行，不计入行数。
    """
    order: list[str] = []
    blocks: dict[str, dict] = {}
    max_row = ws.max_row or 0
    max_col = ws.max_column or 0
    ranges = _merged_ranges(ws)
    current: dict | None = None

    for row in range(1, max_row + 1):
        code = _cell_str(ws, row, 1)
        if CODE_RE.match(code):
            if current is not None and current["code"] == code:
                if current["col0"] == 3 and _cell_str(ws, row, 2) == MARK_DELETED:
                    continue  # 「已删」墓碑行不计入网格
                current["grid_rows"].append(row)  # 块内数据行同样带表编号
                continue
            meta = anchor_metadata(ws, row, max_col)
            current = {
                "code": code,
                "anchor_row": row,
                "grid_rows": [],
                "meta": meta,
                "col0": 3 if meta["cols"].get("name") == 3 else 2,
            }
            blocks[code] = current
            order.append(code)
            continue
        if current is None or code != "":
            continue
        if current["col0"] == 3 and _cell_str(ws, row, 2) == MARK_DELETED:
            continue
        if row_grid_has_content(ws, row, max_col, ranges, start_col=current["col0"]):
            current["grid_rows"].append(row)

    for block in blocks.values():
        head_mark = _cell_str(ws, block["grid_rows"][0], 2) if block["grid_rows"] else ""
        second_mark = _cell_str(ws, block["grid_rows"][1], 2) if len(block["grid_rows"]) > 1 else ""
        if block["col0"] == 2 and head_mark in MARK_VALUES and second_mark in (MARK_YES, MARK_NO, MARK_DELETED):
            block["col0"] = 3
        col0 = block["col0"]
        inferred = 0
        for grid_row in block["grid_rows"]:
            for col in range(max(max_col, 2), 1, -1):
                cell = ws.cell(grid_row, col)
                filled = cell.value is not None and str(cell.value).strip() != ""
                # 空格子但带格式，说明它属于导出的网格（例如整列为空的表格列）
                if filled or cell.has_style or _in_merged(ranges, grid_row, col):
                    inferred = max(inferred, col - col0 + 1)
                    break
        block["ncols"] = inferred
        block["bookmark"] = block["meta"].get("bookmark", "")
        block["is_new"] = block["bookmark"] == NEW_BOOKMARK
        if block["is_new"]:
            # 新表块要按"连续网格"取行：空但有样式的行（模板/Excel 插入行）也算表格行
            declared = block["meta"].get("ncols")
            ncols = declared or inferred or 1
            rows = []
            row = block["anchor_row"] + 1
            while row <= max_row:
                code_text = _cell_str(ws, row, 1)
                if code_text and code_text != block["code"]:
                    break
                if _cell_str(ws, row, 2) == MARK_DELETED:
                    break
                has = row_grid_has_content(ws, row, max_col, ranges, start_col=col0) or any(
                    ws.cell(row, col).has_style for col in range(col0, col0 + ncols)
                )
                if not has:
                    break
                rows.append(row)
                row += 1
            block["grid_rows"] = rows
            block["ncols"] = declared or max(inferred, 1)
    return order, blocks


def refresh_anchor_metadata(ws, block: dict, info: dict, docx_name: str, with_link: bool = True, ncols: int | None = None) -> None:
    """把锚点行元数据写回规范位置（col0 名称 / +1 书签 / +2 链接 / +3 列数），并清理错位残留。"""
    row = block["anchor_row"]
    col0 = block.get("col0", 2)
    for col in range(col0, META_SCAN_COLUMNS + 1):
        if col in (col0, col0 + 1, col0 + 2, col0 + 3):
            continue
        cell = ws.cell(row, col)
        text = "" if cell.value is None else str(cell.value).strip()
        if text.startswith(("书签：", "列数：", "位置：", "原位置：", "原标题：")) or text == "→定位Word表":
            cell.value = None
            cell.hyperlink = None

    ws.cell(row, col0, f"【表{info['index'] + 1}】{info['name']} · 章节：{info['section']}")
    ws.cell(row, col0 + 1, f"书签：{info['bookmark']}")
    link_cell = ws.cell(row, col0 + 2, "→定位Word表")
    if with_link:
        link_cell.hyperlink = Hyperlink(
            ref=link_cell.coordinate, target=docx_name, location=info["bookmark"], display="→定位Word表"
        )
    else:
        link_cell.hyperlink = None
    ws.cell(row, col0 + 3, f"列数：{ncols if ncols is not None else info['ncols']}")
    for col in range(1, col0 + 4):
        cell = ws.cell(row, col)
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True, color="1F4E79")
        cell.fill = FILL_ANCHOR
        cell.alignment = Alignment(horizontal="left" if col > 1 else "center", vertical="center")


def refresh_defined_names(wb, blocks: dict[str, dict]) -> list[str]:
    """把 EX_Tnn 重新指向各表数据区首行（用户增删行后仍能准确定位）。"""
    changed = []
    for code, block in blocks.items():
        letter = get_column_letter(block.get("col0", 2))
        ref = f"'{SHEET_NAME}'!${letter}${block['anchor_row'] + 1}"
        name = f"EX_{code}"
        existing = wb.defined_names.get(name) if hasattr(wb.defined_names, "get") else None
        if existing is not None and existing.attr_text == ref:
            continue
        definition = DefinedName(name, attr_text=ref)
        try:
            wb.defined_names.add(definition)
        except AttributeError:
            wb.defined_names[name] = definition
        changed.append(f"{name} → {ref}")
    return changed


def read_recorded_bookmarks(xlsx_path: Path) -> dict[str, str]:
    """读取现有工作簿里记录的表编号 → 书签名，用于复用既有书签。"""
    if not xlsx_path.exists():
        return {}
    try:
        wb = load_workbook_for_save(xlsx_path)
    except Exception as exc:  # 文件被占用或损坏时不阻断主流程
        warn(f"读取现有工作簿失败（{exc}），本次不沿用旧书签名")
        return {}
    if SHEET_NAME not in wb.sheetnames:
        return {}
    _, blocks = parse_blocks(wb[SHEET_NAME])
    return {code: block.get("bookmark", "") for code, block in blocks.items() if block.get("bookmark")}


def block_merges(ws, block: dict) -> list[tuple[int, int, int, int]]:
    """表块网格内的 Excel 合并区 → (起始行, 起始列, 行跨度, 列跨度)，坐标相对块内网格。"""
    rows = block["grid_rows"]
    if not rows:
        return []
    row_index = {r: i for i, r in enumerate(rows)}
    ncols = block["ncols"]
    col0 = block.get("col0", 2)
    result = []
    for (r1, c1, r2, c2) in _merged_ranges(ws):
        if c1 < col0 or c1 > col0 - 1 + ncols or r1 not in row_index or r2 not in row_index:
            continue
        c2 = min(c2, col0 - 1 + ncols)
        result.append((row_index[r1], c1 - col0, row_index[r2] - row_index[r1] + 1, c2 - c1 + 1))
    return result


_CELL_REF_IN_FORMULA_RE = re.compile(r"(?<![A-Za-z0-9_$])(\$?[A-Za-z]{1,3}\$?)(\d+)")


def _shift_formula(text: str, idx: int, amount: int) -> str:
    """按 Excel 插行语义平移公式里的行号：引用行 >= idx 的加 amount（引号内文本不动）。"""

    parts = re.split(r'("[^"]*")', text)
    for i, part in enumerate(parts):
        if part.startswith('"'):
            continue

        def repl(match: re.Match, part: str = part) -> str:
            if part[match.end():match.end() + 1] == "(":  # LOG10( 这类函数名，不是单元格引用
                return match.group(0)
            row = int(match.group(2))
            return f"{match.group(1)}{row + amount if row >= idx else row}"

        parts[i] = _CELL_REF_IN_FORMULA_RE.sub(repl, part)
    return "".join(parts)


def _insert_rows_safe(ws, idx: int, amount: int) -> None:
    """在 idx 行前插入 amount 行，并自行搬运合并区/超链接/公式引用/行高。

    openpyxl 的 insert_rows 只搬单元格本体，这些附件不搬或搬错，必须手工处理。
    """
    from openpyxl.worksheet.dimensions import RowDimension

    merges = _merged_ranges(ws)
    for (r1, c1, r2, c2) in merges:
        ws.unmerge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    heights = {
        r: dim.height for r, dim in ws.row_dimensions.items() if r >= idx and dim.height is not None
    }
    ws.insert_rows(idx, amount)
    for (r1, c1, r2, c2) in merges:
        if r1 >= idx:
            r1 += amount
            r2 += amount
        elif r2 >= idx:
            r2 += amount
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    for r in [r for r in ws.row_dimensions if r >= idx]:
        del ws.row_dimensions[r]
    for r, height in heights.items():
        ws.row_dimensions[r + amount] = RowDimension(ws, index=r + amount, height=height, customHeight=True)
    for row in ws.iter_rows():
        for cell in row:
            if cell.hyperlink is not None:
                cell.hyperlink.ref = cell.coordinate
            if isinstance(cell.value, str) and cell.value.startswith("="):
                cell.value = _shift_formula(cell.value, idx, amount)


def append_new_block(wb, name: str, nrows: int, ncols: int, after: str | None = None) -> str:
    """在指定表块下面物理插入一个规范的新表块（C 列带「书签：新增」标记），返回表编号。

    ``after`` 指定插到哪张表下面（如 "T02"），不指定时插到最后一张表下面；
    后续表块整体下移，全部表块按上下顺序重新编号（T01、T02…依次顺延）。
    Word 建表位置跟随工作簿顺序，即插在哪张表下面就建在哪张表之后。
    网格预置标准边框与字体：空行也带样式，填数后即可被回写识别为新建 Word 表。
    """
    from openpyxl.utils import get_column_letter
    from .docx_grid import position_sides
    from .errors import SyncError
    from .excel_export import make_border

    name = (name or "").strip()
    if not name:
        raise SyncError("请填写新表名称")
    if not 1 <= nrows <= 200 or not 1 <= ncols <= 30:
        raise SyncError("行数限 1～200、列数限 1～30")
    if SHEET_NAME not in wb.sheetnames:
        raise SyncError(f"工作簿中找不到工作表“{SHEET_NAME}”，请确认该文件由本工具导出")
    ws = wb[SHEET_NAME]
    order, blocks = parse_blocks(ws)
    after = (after or "").strip()
    if after and after not in order:
        raise SyncError(f"插入位置 {after} 不在工作簿的表编号里，请重新选择")
    if not after and order:
        after = order[-1]
    next_num = max([int(code[1:]) for code in order if code[1:].isdigit()] + [0]) + 1
    code = f"T{next_num:02d}"

    if after:
        prev = blocks[after]
        anchor = max([prev["anchor_row"]] + prev["grid_rows"]) + 2  # 与上一块空一行
        if anchor <= (ws.max_row or 0):
            _insert_rows_safe(ws, anchor, nrows + 2)  # 锚点行 + 网格行 + 与下一块的空行
    else:
        anchor = (ws.max_row or 0) + 2
    ws.cell(anchor, 1, code)
    ws.cell(anchor, 3, f"【新表】{name}")
    ws.cell(anchor, 4, f"书签：{NEW_BOOKMARK}")
    ws.cell(anchor, 5, "→定位Word表")
    ws.cell(anchor, 6, f"列数：{ncols}")
    for col in range(1, 7):
        cell = ws.cell(anchor, col)
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True, color="1F4E79")
        cell.fill = FILL_ANCHOR
        cell.alignment = Alignment(horizontal="left" if col > 1 else "center", vertical="center")

    borders = {"top": "thin", "bottom": "thin", "left": None, "right": None, "insideH": "dotted", "insideV": "dotted"}
    for r in range(nrows):
        grid_row = anchor + 1 + r
        code_cell = ws.cell(grid_row, 1, code)
        code_cell.fill = FILL_CODE_COLUMN
        code_cell.font = Font(name=FONT_NAME, size=FONT_SIZE, color="808080")
        code_cell.alignment = Alignment(horizontal="center", vertical="center")
        mark_cell = ws.cell(grid_row, 2, MARK_HEADER if r == 0 else MARK_NO)
        mark_cell.font = Font(name=FONT_NAME, size=FONT_SIZE, color="808080" if r == 0 else "C00000")
        mark_cell.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(ncols):
            cell = ws.cell(grid_row, 3 + c)
            cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=(r == 0))
            cell.border = make_border(position_sides(borders, r, c, nrows, ncols))
            cell.alignment = Alignment(horizontal="center" if r == 0 else "left", vertical="center", wrap_text=True)
    for c in range(ncols):
        letter = get_column_letter(3 + c)
        if letter not in ws.column_dimensions or not ws.column_dimensions[letter].width:
            ws.column_dimensions[letter].width = 12.0

    # 物理插入后按上下顺序重新编号：A 列、【表N】标签、EX_ 定义名称一起顺延
    order2, blocks2 = parse_blocks(ws)
    renamed: dict[str, str] = {}
    final_blocks: dict[str, dict] = {}
    for i, old_code in enumerate(order2, 1):
        new_code = f"T{i:02d}"
        block = blocks2[old_code]
        renamed[old_code] = new_code
        if old_code != new_code:
            ws.cell(block["anchor_row"], 1, new_code)
            for r in block["grid_rows"]:
                ws.cell(r, 1, new_code)
            label = ws.cell(block["anchor_row"], block["col0"]).value
            if isinstance(label, str):
                ws.cell(block["anchor_row"], block["col0"]).value = re.sub(r"【表\d+】", f"【表{i}】", label, count=1)
        final_blocks[new_code] = block
    for nm in list(wb.defined_names.keys()):
        if nm.startswith("EX_") and nm[3:] not in final_blocks:
            del wb.defined_names[nm]
    refresh_defined_names(wb, final_blocks)
    return renamed[code]


def rejudge_marks(ws, blocks: dict) -> dict:
    """按工作簿当前内容重新判定 B 列标记（表头/✓/✕），只改标记不改数据。

    仅新布局表块（col0=3）参与：网格首行=表头；其余行除首列外任一列有内容
    （含公式、「—」等占位文字）=✓，否则=✕，口径与导出时一致。
    「已删」行不在 grid_rows 里，天然保持原样。旧布局块记入 old 列表提示重新导出。
    """
    changed = 0
    rows = 0
    old: list[str] = []
    for code, block in blocks.items():
        if block.get("col0", 2) != 3:
            old.append(code)
            continue
        col0, ncols = block["col0"], block["ncols"]
        for i, row in enumerate(block["grid_rows"]):
            rows += 1
            if i == 0:
                mark, color = MARK_HEADER, "808080"
            else:
                has = False
                for c in range(1, ncols):
                    cell = ws.cell(row, col0 + c)
                    value = cell.value
                    if isinstance(value, str) and value.startswith("="):
                        has = True
                        break
                    if value is not None and str(value).strip():
                        has = True
                        break
                mark, color = (MARK_YES, "008000") if has else (MARK_NO, "C00000")
            mark_cell = ws.cell(row, 2)
            if mark_cell.value != mark:
                changed += 1
            mark_cell.value = mark
            mark_cell.font = Font(name=FONT_NAME, size=FONT_SIZE, color=color)
            mark_cell.alignment = Alignment(horizontal="center", vertical="center")
    return {"changed": changed, "rows": rows, "old": old}


def carry_over_deleted_blocks(wb_new, ws_old, blocks_old: dict) -> list[str]:
    """导出重建工作簿时，把「已可逆删除」的表块原样保留到新工作簿末尾。

    复制值、样式、合并区与锚点记录（原位置/原标题），编号顺延；返回保留的表编号。
    """
    from copy import copy
    from openpyxl.utils import get_column_letter

    deleted = [b for b in blocks_old.values() if b.get("bookmark") == DELETED_BOOKMARK]
    if not deleted:
        return []
    ws = wb_new[SHEET_NAME]
    order, _ = parse_blocks(ws)
    next_num = max([int(code[1:]) for code in order if code[1:].isdigit()] + [0]) + 1
    carried: list[str] = []
    new_blocks: dict[str, dict] = {}
    for block in deleted:
        code = f"T{next_num:02d}"
        next_num += 1
        old_anchor = block["anchor_row"]
        old_col0 = block.get("col0", 2)
        anchor = (ws.max_row or 0) + 2
        for col in range(1, META_SCAN_COLUMNS + 1):
            old = ws_old.cell(old_anchor, col)
            if old.value is None and not old.has_style:
                continue
            new_col = 1 if col == 1 else (col if col < old_col0 else col + 3 - old_col0)
            new = ws.cell(anchor, new_col)
            new.value = old.value
            if old.has_style:
                new._style = copy(old._style)
        ws.cell(anchor, 1, code)
        ws.cell(anchor, 4, f"书签：{DELETED_BOOKMARK}")
        link = ws.cell(anchor, 5)
        link.value = None
        link.hyperlink = None
        grid_rows = []
        for i, old_row in enumerate(block["grid_rows"]):
            new_row = anchor + 1 + i
            grid_rows.append(new_row)
            old = ws_old.cell(old_row, 1)
            new = ws.cell(new_row, 1)
            new.value = code
            if old.has_style:
                new._style = copy(old._style)
            if old_col0 == 3:
                old = ws_old.cell(old_row, 2)
                new = ws.cell(new_row, 2)
                new.value = old.value
                if old.has_style:
                    new._style = copy(old._style)
            for c in range(block["ncols"]):
                old = ws_old.cell(old_row, old_col0 + c)
                new = ws.cell(new_row, 3 + c)
                new.value = old.value
                if old.has_style:
                    new._style = copy(old._style)
        for (r, c, rs, cs) in block_merges(ws_old, block):
            if rs == 1 and cs == 1:
                continue
            ws.merge_cells(
                start_row=anchor + 1 + r, start_column=3 + c,
                end_row=anchor + r + rs, end_column=2 + c + cs,
            )
        for c in range(block["ncols"]):
            old_dim = ws_old.column_dimensions.get(get_column_letter(old_col0 + c))
            if old_dim is not None and old_dim.width:
                letter = get_column_letter(3 + c)
                cur = ws.column_dimensions[letter].width or 0
                ws.column_dimensions[letter].width = max(cur, old_dim.width)
        new_blocks[code] = {"code": code, "anchor_row": anchor, "grid_rows": grid_rows, "ncols": block["ncols"], "col0": 3}
        carried.append(code)
    refresh_defined_names(wb_new, new_blocks)
    return carried


def refresh_workbook_metadata(
    xlsx_path: Path,
    docx_name: str,
    infos_by_code: dict[str, dict],
    with_links: bool = True,
) -> int:
    """刷新工作簿里的书签名、跳转链接与列数（不动任何数据单元格）。"""
    wb = load_workbook_for_save(xlsx_path)
    if SHEET_NAME not in wb.sheetnames:
        return 0
    ws = wb[SHEET_NAME]
    _, blocks = parse_blocks(ws)
    changed = 0
    for code, block in blocks.items():
        info = infos_by_code.get(code)
        if info is None:
            continue
        row = block["anchor_row"]
        col0 = block.get("col0", 2)
        before = [
            ws.cell(row, col).value for col in range(1, col0 + 4)
        ] + [
            (ws.cell(row, col0 + 2).hyperlink.location if ws.cell(row, col0 + 2).hyperlink else None)
        ]
        refresh_anchor_metadata(ws, block, info, docx_name, with_link=with_links, ncols=block["ncols"])
        after = [
            ws.cell(row, col).value for col in range(1, col0 + 4)
        ] + [
            (ws.cell(row, col0 + 2).hyperlink.location if ws.cell(row, col0 + 2).hyperlink else None)
        ]
        if before != after:
            changed += 1
            log(f"[定位] {code}：书签/链接已刷新为 {info['bookmark']}")
    names = refresh_defined_names(wb, blocks)
    if changed or names:
        backup = backup_file(xlsx_path)
        wb.save(str(xlsx_path))
        log(f"[定位] 已保存 {xlsx_path.name}" + (f"（原文件已备份：{backup.name}）" if backup else ""))
    else:
        log("[定位] Excel 中的书签名与链接已是目标状态，无需改动")
    return changed + len(names)
