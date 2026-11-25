"""行结构同步：表头/合计行识别、按 Excel 行数增行删行，并返回被删行的首列文本。"""

import copy
from docx.oxml.ns import qn
from .docx_grid import cell_text
from .errors import SyncError


def first_cell_text(tr) -> str:
    tc = tr.find(qn("w:tc"))
    return cell_text(tc) if tc is not None else ""


def leading_header_count(trs) -> int:
    count = 0
    for tr in trs:
        trPr = tr.find(qn("w:trPr"))
        if trPr is not None and trPr.find(qn("w:tblHeader")) is not None:
            count += 1
        else:
            break
    return count


def split_rows(trs):
    header = leading_header_count(trs)
    head_rows = trs[:header]
    rest = trs[header:]
    total_index = None
    for idx, tr in enumerate(rest):
        if first_cell_text(tr).strip().startswith("合计"):
            total_index = idx
            break
    if total_index is None:
        return head_rows, list(rest), []
    return head_rows, list(rest[:total_index]), list(rest[total_index:])


def align_rows(tbl, target_total_rows: int) -> tuple[int, list[str]]:
    """让数据区行数对齐 Excel。返回 (增删行数, 被删除行的首列文本)。"""
    trs = tbl.findall(qn("w:tr"))
    head_rows, body_rows, total_rows = split_rows(trs)
    target_body = target_total_rows - len(head_rows) - len(total_rows)
    if target_body < 0:
        raise SyncError(f"Excel 中的行数（{target_total_rows}）少于 Word 表头行数与合计行数之和，无法对齐")
    delta = target_body - len(body_rows)
    if delta == 0:
        return 0, []
    if delta > 0:
        if not body_rows:
            raise SyncError("Word 表格没有可用于复制格式的数据行，无法新增行，请手工调整")
        template = body_rows[-1]
        if template.findall(f".//{qn('w:vMerge')}"):
            raise SyncError("模板数据行含纵向合并单元格，自动增行不安全，请手工在 Word 中调整后重试")
        anchor_node = total_rows[0] if total_rows else body_rows[-1]
        for _ in range(delta):
            clone = copy.deepcopy(template)
            anchor_node.addprevious(clone)
        return delta, []
    deleted: list[str] = []
    for _ in range(-delta):
        if not body_rows:
            raise SyncError("Excel 中删除了过多数据行，Word 表格已无对应行可删，请手工调整")
        row = body_rows.pop()
        deleted.append((first_cell_text(row).strip() or "（空行）").replace("\n", " "))
        row.getparent().remove(row)
    return delta, deleted
