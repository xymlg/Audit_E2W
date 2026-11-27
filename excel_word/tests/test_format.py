import base64
import copy
import io
import json
import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Inches, Pt, RGBColor

from helpers import ws_dir
from wesync import api
from wesync.errors import SyncError
from wesync.report_format import DEFAULT_SCOPE, _document_snapshot, build_profile


def _heading(doc, text, level=1):
    paragraph = doc.add_paragraph(text, style=f"Heading {level}")
    return paragraph


def _bookmark(paragraph, name, bookmark_id):
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def _hyperlink(paragraph, text, url):
    relationship = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship)
    run = OxmlElement("w:r")
    text_element = OxmlElement("w:t")
    text_element.text = text
    run.append(text_element)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def _template(path: Path) -> Path:
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(10.5)
    title = doc.add_paragraph("审计报告")
    title.alignment = 1
    for run in title.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(26)
        run.font.bold = True
    doc.add_paragraph("目录")
    heading = _heading(doc, "一、审计意见", 1)
    heading.runs[0].font.size = Pt(12)
    body = doc.add_paragraph("这是样板正文，用于提取字体和段落格式。")
    body.paragraph_format.first_line_indent = Pt(21)
    body.paragraph_format.line_spacing = Pt(20)
    doc.add_paragraph("金额单位：元")
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    table.cell(0, 0).text = "项目"
    table.cell(0, 1).text = "金额"
    table.cell(1, 0).text = "合计"
    table.cell(1, 1).text = "100.00"
    for row in table.rows:
        for cell in row.cells:
            for run in cell.paragraphs[0].runs:
                run.font.size = Pt(9)
    doc.save(str(path))
    return path


def _messy(path: Path) -> Path:
    doc = Document()
    title = doc.add_paragraph("乱版审计报告")
    title.alignment = 3
    for run in title.runs:
        run.font.name = "微软雅黑"
        run.font.size = Pt(36)
    doc.add_paragraph("目录")
    heading = _heading(doc, "一、审计意见", 1)
    numbering = OxmlElement("w:numPr")
    level = OxmlElement("w:ilvl")
    level.set(qn("w:val"), "0")
    number = OxmlElement("w:numId")
    number.set(qn("w:val"), "1")
    numbering.extend((level, number))
    heading._p.get_or_add_pPr().append(numbering)
    body = doc.add_paragraph()
    run = body.add_run("保留文字和审计痕迹")
    run.font.name = "微软雅黑"
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.highlight_color = 7
    run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    run.font.underline = True
    _bookmark(body, "CPAH_KEEP", 1)
    _hyperlink(body, "链接文字", "https://example.invalid/audit")
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    field_run = OxmlElement("w:r")
    field_text = OxmlElement("w:t")
    field_text.text = "1"
    field_run.append(field_text)
    field.append(field_run)
    body._p.append(field)
    doc.add_paragraph("金额单位：元")

    png = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
    doc.add_picture(io.BytesIO(png), width=Inches(0.1))

    table = doc.add_table(rows=3, cols=2)
    table.style = "Table Grid"
    table.cell(0, 0).merge(table.cell(0, 1)).text = "合并表头"
    table.cell(1, 0).text = "小报表项目"
    small = table.cell(1, 1).paragraphs[0].add_run("1,234.00")
    small.font.size = Pt(6.5)
    table.cell(2, 0).text = "合计"
    table.cell(2, 1).text = "1,234.00"

    section = doc.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    doc.add_paragraph("横向页正文保持不变。")
    doc.save(str(path))
    return path


def test_format_scan_is_read_only_and_reports_roles():
    d = ws_dir("format_scan")
    template = _template(d / "样板.docx")
    target = _messy(d / "乱版.docx")
    before = target.read_bytes()
    result = api.scan_report_format(target, template)
    assert target.read_bytes() == before
    assert result["totals"]["tables"] == 1
    assert result["totals"]["table_rows"] == 3
    labels = {item["role"] for item in result["roles"]}
    assert {"cover_title", "toc_title", "heading_1", "body", "table_title"} <= labels


def _outline(paragraph, level):
    element = OxmlElement("w:outlineLvl")
    element.set(qn("w:val"), str(level))
    paragraph._p.get_or_add_pPr().append(element)


def _notes_template(path: Path) -> Path:
    doc = Document()
    defaults = doc.styles.element.find(qn("w:docDefaults"))
    default_size = defaults.find(qn("w:rPrDefault")).find(qn("w:rPr")).find(qn("w:sz"))
    default_size.set(qn("w:val"), "18")
    normal = doc.styles["Normal"]
    normal.font.size = Pt(10.5)
    normal.paragraph_format.alignment = 3
    audit_title = doc.add_paragraph("审计报告")
    audit_title.runs[0].font.size = Pt(26)
    audit_heading = doc.add_paragraph("一、审计意见")
    audit_heading.alignment = 1
    audit_heading.runs[0].font.size = Pt(14)
    audit_body = doc.add_paragraph("审计报告正文不应混入附注格式。")
    audit_body.runs[0].font.size = Pt(8)
    audit_table = doc.add_table(rows=1, cols=1)
    audit_table.cell(0, 0).text = "审计报表"
    audit_table.cell(0, 0).paragraphs[0].runs[0].font.size = Pt(7)

    for text, size in (
        ("样例公司", 16),
        ("财务报表附注", 16),
        ("2026年1-6月、2025年度", 16),
        ("（除特别说明外，金额单位为人民币元）", 12),
    ):
        paragraph = doc.add_paragraph(text)
        paragraph.alignment = 1
        paragraph.runs[0].font.size = Pt(size)
        paragraph.runs[0].bold = size == 16
    heading = doc.add_paragraph("一、公司基本情况")
    _outline(heading, 0)
    heading.alignment = 0
    heading.runs[0].font.size = Pt(12)
    heading.runs[0].bold = True
    body = doc.add_paragraph("附注正文继承 Normal 的 10.5 磅字号。")
    doc.add_paragraph("各类固定资产折旧年限如下：")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "项目"
    table.cell(0, 1).text = "金额"
    table.cell(1, 0).text = "合计"
    table.cell(1, 1).text = "100.00"
    for row in table.rows:
        for cell in row.cells:
            cell.paragraphs[0].runs[0].font.size = Pt(9)
    caption = doc.add_paragraph("金额单位：元")
    caption.alignment = 1
    caption.runs[0].font.size = Pt(9)
    caption.runs[0].bold = True
    doc.add_table(rows=1, cols=1).cell(0, 0).text = "单位表"
    continuation = doc.add_paragraph("（续上表）")
    continuation.alignment = 2
    continuation.runs[0].font.size = Pt(10.5)
    doc.add_table(rows=1, cols=1).cell(0, 0).text = "第二张表"
    doc.save(str(path))
    return path


def _notes_target(path: Path, combined_title: bool = False) -> Path:
    doc = Document()
    opening = (
        ("样例公司", "2026年1-6月财务报表附注", "（除特别说明外，金额单位为人民币元）")
        if combined_title
        else ("样例公司", "财务报表附注", "2026年1-6月、2025年度", "（除特别说明外，金额单位为人民币元）")
    )
    for text in (
        *opening, "一、公司基本情况",
        "附注正文继承 Normal 的 10.5 磅字号。", "各类固定资产折旧年限如下：",
    ):
        paragraph = doc.add_paragraph(text)
        paragraph.runs[0].font.size = Pt(9)
        if text == "一、公司基本情况":
            _outline(paragraph, 0)
        if text.endswith("如下："):
            doc.add_table(rows=1, cols=1).cell(0, 0).text = "引导正文表"
    doc.add_paragraph("金额单位：元").runs[0].font.size = Pt(9)
    doc.add_table(rows=1, cols=1).cell(0, 0).text = "单位表"
    doc.add_paragraph("（续上表）").runs[0].font.size = Pt(9)
    doc.add_table(rows=1, cols=1).cell(0, 0).text = "附注表格"
    doc.save(str(path))
    return path


def test_notes_profile_uses_only_notes_region_and_resolves_style_defaults():
    d = ws_dir("format_notes_region")
    template = _notes_template(d / "完整样板.docx")
    target = _notes_target(d / "附注.docx")
    profile = build_profile(template, target)
    assert profile.paragraphs["cover_title"].size == 16
    assert profile.paragraphs["cover_info"].size == 12
    assert profile.paragraphs["heading_1"].size == 12
    assert profile.paragraphs["heading_1"].align == "left"
    assert profile.paragraphs["body"].size == 10.5
    assert profile.paragraphs["body"].align == "both"
    assert profile.paragraphs["table_title"].size == 9
    assert profile.paragraphs["table_continuation"].size == 10.5
    assert profile.paragraphs["table_continuation"].align == "right"
    assert profile.table_size == 9

    output = d / "附注_格式调整.docx"
    before = _document_snapshot(Document(str(target)))
    result = api.run_report_format(target, template, output=output)
    assert result["applied"] is True
    assert _document_snapshot(Document(str(output))) == before
    paragraphs = Document(str(output)).paragraphs
    assert [paragraph.runs[0].font.size for paragraph in paragraphs[:4]] == [Pt(16), Pt(16), Pt(16), Pt(12)]
    heading = next(paragraph for paragraph in paragraphs if paragraph.text == "一、公司基本情况")
    assert heading.alignment == 0
    assert heading.runs[0].font.size == Pt(12)
    body = next(paragraph for paragraph in paragraphs if "继承 Normal" in paragraph.text)
    assert body.runs[0].font.size == Pt(10.5)
    lead_in = next(paragraph for paragraph in paragraphs if paragraph.text.endswith("如下："))
    assert lead_in.runs[0].font.size == Pt(10.5)
    caption = next(paragraph for paragraph in paragraphs if paragraph.text == "金额单位：元")
    assert caption.runs[0].font.size == Pt(9)
    assert caption.runs[0].bold is True
    continuation = next(paragraph for paragraph in paragraphs if paragraph.text == "（续上表）")
    assert continuation.runs[0].font.size == Pt(10.5)
    assert continuation.alignment == 2


def test_combined_notes_title_keeps_three_line_opening():
    d = ws_dir("format_notes_combined_title")
    template = _notes_template(d / "完整样板.docx")
    target = _notes_target(d / "合并标题附注.docx", combined_title=True)
    output = d / "合并标题附注_格式调整.docx"
    api.run_report_format(target, template, output=output)
    paragraphs = Document(str(output)).paragraphs
    assert [paragraph.runs[0].font.size for paragraph in paragraphs[:3]] == [Pt(16), Pt(16), Pt(12)]


def _set_table_borders(table, values: dict[str, tuple[str, int, str]]) -> None:
    tbl_pr = table._tbl.tblPr
    previous = tbl_pr.find(qn("w:tblBorders"))
    if previous is not None:
        tbl_pr.remove(previous)
    borders = OxmlElement("w:tblBorders")
    for edge, (val, size, color) in values.items():
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), val)
        element.set(qn("w:sz"), str(size))
        element.set(qn("w:color"), color)
        borders.append(element)
    tbl_pr.insert(0, borders)


def _border_snapshot(table) -> tuple[str, tuple[str, ...]]:
    table_borders = table._tbl.tblPr.find(qn("w:tblBorders"))
    cell_borders = tuple(
        cell.find(qn("w:tcPr")).find(qn("w:tcBorders")).xml
        for cell in table._tbl.xpath(".//w:tc")
        if cell.find(qn("w:tcPr")) is not None
        and cell.find(qn("w:tcPr")).find(qn("w:tcBorders")) is not None
    )
    return table_borders.xml if table_borders is not None else "", cell_borders


def _table_without_borders(table) -> str:
    element = copy.deepcopy(table._tbl)
    for container in [element.tblPr, *element.xpath(".//w:tcPr")]:
        borders = container.find(qn("w:tblBorders"))
        if borders is None:
            borders = container.find(qn("w:tcBorders"))
        if borders is not None:
            container.remove(borders)
    return element.xml


def test_format_scope_borders_does_not_change_text_or_emphasis():
    d = ws_dir("format_scope_borders")
    template = _template(d / "样板.docx")
    target = _messy(d / "乱版.docx")
    template_doc = Document(str(template))
    target_doc = Document(str(target))
    expected = {
        "top": ("single", 12, "112233"),
        "left": ("nil", 0, "auto"),
        "bottom": ("double", 16, "112233"),
        "right": ("single", 12, "112233"),
        "insideH": ("dotted", 6, "445566"),
        "insideV": ("nil", 0, "auto"),
    }
    _set_table_borders(template_doc.tables[0], expected)
    _set_table_borders(target_doc.tables[0], {edge: ("single", 4, "FF0000") for edge in expected})
    cell = target_doc.tables[0].cell(1, 1)
    cell_borders = OxmlElement("w:tcBorders")
    right = OxmlElement("w:right")
    right.set(qn("w:val"), "double")
    right.set(qn("w:sz"), "24")
    right.set(qn("w:color"), "0000FF")
    cell_borders.append(right)
    cell._tc.get_or_add_tcPr().append(cell_borders)
    template_doc.save(str(template))
    target_doc.save(str(target))

    source_doc = Document(str(target))
    source_body = next(p for p in source_doc.paragraphs if "保留文字和审计痕迹" in p.text)
    source_table = source_doc.tables[0]
    scope = _scope(cover=False, headings=False, statement=False, body=False, table_text=False, table_emphasis=False, table_borders=True)
    preview = api.scan_report_format(target, template, scope=scope)
    assert preview["totals"]["selected_paragraphs"] == 0
    assert preview["totals"]["tables"] == 1
    assert "内部横线" in preview["table_border_format"]
    output = d / "仅框线.docx"
    result = api.run_report_format(target, template, output=output, scope=scope)
    formatted = Document(str(output))
    output_body = next(p for p in formatted.paragraphs if "保留文字和审计痕迹" in p.text)
    output_table = formatted.tables[0]
    assert output_body._p.xml == source_body._p.xml
    assert _table_without_borders(output_table) == _table_without_borders(source_table)
    assert result["counts"]["table_borders"] == 1
    assert result["counts"]["table_paragraphs"] == 0
    assert output_table.cell(1, 1).paragraphs[0].runs[0].font.size == Pt(6.5)
    borders = output_table._tbl.tblPr.find(qn("w:tblBorders"))
    assert borders is not None
    for edge, (val, size, color) in expected.items():
        element = borders.find(qn(f"w:{edge}"))
        assert element.get(qn("w:val")) == val
        if val not in ("nil", "none"):
            assert int(element.get(qn("w:sz"))) == size
            assert element.get(qn("w:color")) == color
    assert not output_table._tbl.xpath(".//w:tcBorders")


def test_format_scope_borders_disabled_preserves_existing_borders():
    d = ws_dir("format_scope_no_borders")
    template = _template(d / "样板.docx")
    target = _messy(d / "乱版.docx")
    output = d / "关闭框线.docx"
    scope = _scope(cover=False, headings=False, statement=False, body=True, table_text=False, table_emphasis=False, table_borders=False)
    api.run_report_format(target, template, output=output, scope=scope)
    assert _border_snapshot(Document(str(output)).tables[0]) == _border_snapshot(Document(str(target)).tables[0])


def _heading_candidate_report(path: Path) -> Path:
    doc = Document()
    doc.add_paragraph("审计报告")
    doc.add_paragraph("报告日期")
    doc.add_paragraph("目录")
    first = doc.add_paragraph("一、总体情况")
    first.runs[0].bold = True
    doc.add_paragraph("这是普通正文，用于确认模型候选之外的内容不会发送，也不会被改成标题。")
    second = doc.add_paragraph("（一）具体原因")
    second.runs[0].bold = True
    short = doc.add_paragraph("独立短标题")
    short.runs[0].bold = True
    short.runs[0].font.size = Pt(16)
    _heading(doc, "已有二级标题", 2)
    doc.add_paragraph("另一段普通正文，保留原来的字体和段落格式。")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "项目"
    table.cell(0, 1).text = "金额"
    table.cell(1, 0).text = "合计"
    table.cell(1, 1).text = "100.00"
    protected = doc.add_paragraph("异常大纲段落保持原样")
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), "9")
    protected._p.get_or_add_pPr().append(outline)
    doc.add_paragraph("不应进入候选包的隔离段落")
    doc.save(str(path))
    return path


def test_heading_candidates_are_local_minimal_and_read_only():
    d = ws_dir("format_heading_candidates")
    target = _heading_candidate_report(d / "候选标题.docx")
    before = target.read_bytes()
    package = api.scan_format_heading_candidates(target)
    assert target.read_bytes() == before
    assert set(package) == {"batch_id", "candidates", "role_candidates", "boundary_candidates", "candidate_count", "truncated", "privacy", "apply_mode"}
    assert package["candidate_count"] == 4
    assert package["truncated"] is False
    assert {item["text"] for item in package["candidates"]} == {"一、总体情况", "（一）具体原因", "独立短标题"}
    assert all(set(item) == {"id", "element_index", "text", "before", "after", "style", "bold", "size", "align", "numbered"} for item in package["candidates"])
    assert all(len(item["before"]) <= 100 and len(item["after"]) <= 100 for item in package["candidates"])
    assert [item["text"] for item in package["role_candidates"]] == ["另一段普通正文，保留原来的字体和段落格式。"]
    assert package["role_candidates"][0]["options"] == ["body", "table_title"]
    assert package["role_candidates"][0]["next_element"].startswith("表格")
    assert "不应进入候选包" not in str(package)
    assert "100.00" not in str(package)


def test_confirmed_heading_levels_apply_only_direct_formatting():
    d = ws_dir("format_heading_apply")
    template = _template(d / "样板.docx")
    target = _heading_candidate_report(d / "候选标题.docx")
    package = api.scan_format_heading_candidates(target)
    levels = {package["candidates"][0]["id"]: 1, package["candidates"][1]["id"]: 2}
    scope = _scope(cover=False, headings=True, statement=False, body=False, table_text=False, table_emphasis=False, table_borders=False)
    output = d / "候选标题格式.docx"
    before = _document_snapshot(Document(str(target)))
    preview = api.scan_report_format(target, template, scope=scope, heading_batch_id=package["batch_id"], heading_levels=levels)
    assert preview["totals"]["selected_paragraphs"] == 3
    result = api.run_report_format(
        target, template, output=output, scope=scope,
        heading_batch_id=package["batch_id"], heading_levels=levels,
    )
    assert result["assigned_headings"] == 2
    assert _document_snapshot(Document(str(output))) == before
    source_doc = Document(str(target))
    formatted = Document(str(output))
    for text, size in (("一、总体情况", Pt(12)), ("（一）具体原因", Pt(10.5))):
        source = next(p for p in source_doc.paragraphs if p.text == text)
        after = next(p for p in formatted.paragraphs if p.text == text)
        assert after.style.style_id == source.style.style_id
        assert after.style.name == source.style.name
        assert after._p.pPr.find(qn("w:outlineLvl")) is None
        assert after.runs[0].font.size == size
    source_body = next(p for p in source_doc.paragraphs if "这是普通正文" in p.text)
    output_body = next(p for p in formatted.paragraphs if "这是普通正文" in p.text)
    assert output_body._p.xml == source_body._p.xml
    unchanged_heading = next(p for p in formatted.paragraphs if p.text == "独立短标题")
    assert unchanged_heading.runs[0].font.size == Pt(16)


def test_heading_candidates_reject_stale_unknown_and_unselected_levels():
    d = ws_dir("format_heading_invalid")
    template = _template(d / "样板.docx")
    target = _heading_candidate_report(d / "候选标题.docx")
    package = api.scan_format_heading_candidates(target)
    levels = {package["candidates"][0]["id"]: 1}
    with pytest.raises(SyncError, match="同时提供"):
        api.run_report_format(target, template, heading_levels=levels)
    with pytest.raises(SyncError, match="不存在"):
        api.run_report_format(target, template, heading_batch_id=package["batch_id"], heading_levels={"unknown": 1})
    with pytest.raises(SyncError, match="1～4"):
        api.run_report_format(target, template, heading_batch_id=package["batch_id"], heading_levels={package["candidates"][0]["id"]: 5})
    scope = _scope(headings=False)
    with pytest.raises(SyncError, match="必须勾选"):
        api.run_report_format(target, template, scope=scope, heading_batch_id=package["batch_id"], heading_levels=levels)
    changed = Document(str(target))
    changed.paragraphs[3].add_run("变化")
    changed.save(str(target))
    with pytest.raises(SyncError, match="重新扫描"):
        api.run_report_format(target, template, heading_batch_id=package["batch_id"], heading_levels=levels)


def test_confirmed_table_role_and_body_override_are_applied():
    d = ws_dir("format_role_apply")
    template = _notes_template(d / "完整样板.docx")
    target = _notes_target(d / "附注.docx")
    package = api.scan_format_heading_candidates(target)
    lead = next(item for item in package["role_candidates"] if item["text"].endswith("如下："))
    scope = _scope(cover=False, headings=False, statement=True, body=False, table_text=False, table_emphasis=False, table_borders=False)
    output = d / "表题角色.docx"
    before = _document_snapshot(Document(str(target)))
    preview = api.scan_report_format(
        target, template, scope=scope, role_batch_id=package["batch_id"],
        role_overrides={lead["id"]: "table_title"},
    )
    assert preview["totals"]["selected_paragraphs"] == 3
    result = api.run_report_format(
        target, template, output=output, scope=scope,
        role_batch_id=package["batch_id"], role_overrides={lead["id"]: "table_title"},
    )
    assert result["assigned_roles"] == 1
    assert result["assigned_headings"] == 0
    assert _document_snapshot(Document(str(output))) == before
    paragraph = next(p for p in Document(str(output)).paragraphs if p.text.endswith("如下："))
    assert paragraph.runs[0].font.size == Pt(9)
    assert paragraph.runs[0].font.bold is True
    assert paragraph.alignment == 1

    body_output = d / "保持正文角色.docx"
    body_result = api.run_report_format(
        target, template, output=body_output, scope=scope,
        role_batch_id=package["batch_id"], role_overrides={lead["id"]: "body"},
    )
    assert body_result["assigned_roles"] == 1
    source_lead = next(p for p in Document(str(target)).paragraphs if p.text.endswith("如下："))
    output_lead = next(p for p in Document(str(body_output)).paragraphs if p.text.endswith("如下："))
    assert output_lead._p.xml == source_lead._p.xml


def test_role_overrides_reject_unknown_unconfirmed_and_unselected_items():
    d = ws_dir("format_role_invalid")
    template = _notes_template(d / "完整样板.docx")
    target = _notes_target(d / "附注.docx")
    package = api.scan_format_heading_candidates(target)
    candidate = package["role_candidates"][0]
    with pytest.raises(SyncError, match="同时提供角色批次"):
        api.scan_report_format(target, template, role_overrides={candidate["id"]: "body"})
    with pytest.raises(SyncError, match="不存在"):
        api.scan_report_format(target, template, role_batch_id=package["batch_id"], role_overrides={"unknown": "body"})
    with pytest.raises(SyncError, match="只能选择"):
        api.scan_report_format(target, template, role_batch_id=package["batch_id"], role_overrides={candidate["id"]: "statement"})
    with pytest.raises(SyncError, match="必须勾选"):
        api.scan_report_format(
            target, template, scope=_scope(statement=False),
            role_batch_id=package["batch_id"], role_overrides={candidate["id"]: "table_title"},
        )


def test_confirmed_notes_boundary_limits_cover_formatting():
    d = ws_dir("format_boundary_apply")
    template = _notes_template(d / "完整样板.docx")
    target = _notes_target(d / "附注.docx")
    package = api.scan_format_heading_candidates(target)
    title = next(item for item in package["boundary_candidates"] if item["text"] == "财务报表附注")
    current = next(item for item in package["boundary_candidates"] if item["current_boundary"])
    assert current["text"] == "样例公司"
    scope = _scope(cover=True, headings=False, statement=False, body=False, table_text=False, table_emphasis=False, table_borders=False)
    output = d / "边界调整.docx"
    preview = api.scan_report_format(
        target, template, scope=scope, role_batch_id=package["batch_id"], region_start_id=title["id"],
    )
    assert preview["totals"]["selected_paragraphs"] == 3
    result = api.run_report_format(
        target, template, output=output, scope=scope,
        role_batch_id=package["batch_id"], region_start_id=title["id"],
    )
    assert result["region_start"] == title["element_index"]
    paragraphs = Document(str(output)).paragraphs
    assert paragraphs[0].runs[0].font.size == Pt(9)
    assert [p.runs[0].font.size for p in paragraphs[1:4]] == [Pt(16), Pt(16), Pt(12)]
    with pytest.raises(SyncError, match="边界编号不存在"):
        api.scan_report_format(target, template, role_batch_id=package["batch_id"], region_start_id="unknown")
    with pytest.raises(SyncError, match="必须同时提供角色批次"):
        api.scan_report_format(target, template, region_start_id=title["id"])


def test_format_apply_preserves_content_and_creates_copy():
    d = ws_dir("format_apply")
    template = _template(d / "样板.docx")
    target = _messy(d / "乱版.docx")
    output = d / "乱版_格式调整.docx"
    before_bytes = target.read_bytes()
    before = _document_snapshot(Document(str(target)))
    result = api.run_report_format(target, template, output=output, dry_run=True)
    assert result["dry_run"] is True
    assert not output.exists()
    assert target.read_bytes() == before_bytes

    result = api.run_report_format(target, template, output=output)
    assert result["applied"] is True
    assert len(result["visual_review"]) == 6
    assert output.exists()
    assert target.read_bytes() == before_bytes
    assert _document_snapshot(Document(str(output))) == before

    formatted = Document(str(output))
    body = next(paragraph for paragraph in formatted.paragraphs if "保留文字和审计痕迹" in paragraph.text)
    traced = next(run for run in body.runs if run.text == "保留文字和审计痕迹")
    assert traced.font.size == Pt(10.5)
    assert traced.font.bold is None
    assert traced._r.rPr.find(qn("w:rFonts")) is not None
    assert traced.font.highlight_color is not None
    assert traced.font.underline is True
    assert traced.font.color.rgb == RGBColor(0xC0, 0x00, 0x00)
    assert body._p.xpath(".//w:fldSimple/@w:instr") == ["PAGE"]
    source_heading = next(p for p in Document(str(target)).paragraphs if p.text == "一、审计意见")
    output_heading = next(p for p in formatted.paragraphs if p.text == "一、审计意见")
    assert output_heading._p.pPr.find(qn("w:numPr")).xml == source_heading._p.pPr.find(qn("w:numPr")).xml
    assert any(paragraph.text == "横向页正文保持不变。" for paragraph in formatted.paragraphs)
    assert formatted.sections[-1].orientation == WD_ORIENT.LANDSCAPE
    assert formatted.tables[0].cell(1, 1).paragraphs[0].runs[0].font.size == Pt(6.5)


def _scope(**changes):
    result = dict(DEFAULT_SCOPE)
    result.update(changes)
    return result


def test_format_scope_table_text_does_not_change_paragraphs_or_emphasis():
    d = ws_dir("format_scope_table_text")
    template = _template(d / "样板.docx")
    target = _messy(d / "乱版.docx")
    output = d / "仅表格字体.docx"
    scope = _scope(cover=False, headings=False, statement=False, body=False, table_text=True, table_emphasis=False, table_borders=False)
    preview = api.scan_report_format(target, template, scope=scope)
    assert preview["totals"]["selected_paragraphs"] == 0
    assert preview["totals"]["tables"] == 1
    assert all(not role["selected"] for role in preview["roles"])

    result = api.run_report_format(target, template, output=output, scope=scope)
    source_doc = Document(str(target))
    formatted = Document(str(output))
    source_body = next(p for p in source_doc.paragraphs if "保留文字和审计痕迹" in p.text)
    output_body = next(p for p in formatted.paragraphs if "保留文字和审计痕迹" in p.text)
    assert output_body._p.xml == source_body._p.xml
    assert formatted.tables[0].cell(0, 0).paragraphs[0].alignment == source_doc.tables[0].cell(0, 0).paragraphs[0].alignment
    assert formatted.tables[0].cell(1, 1).paragraphs[0].runs[0].font.size == Pt(6.5)
    assert formatted.tables[0].cell(2, 1).paragraphs[0].runs[0].font.size == Pt(9)
    assert not formatted.tables[0].cell(0, 0).paragraphs[0].runs[0].font.bold
    assert result["counts"]["skipped_paragraphs"] == result["totals"]["paragraphs"]
    assert "body" not in result["counts"]


def test_format_scope_body_does_not_change_tables():
    d = ws_dir("format_scope_body")
    template = _template(d / "样板.docx")
    target = _messy(d / "乱版.docx")
    output = d / "仅正文.docx"
    scope = _scope(cover=False, headings=False, statement=False, body=True, table_text=False, table_emphasis=False, table_borders=False)
    result = api.run_report_format(target, template, output=output, scope=scope)
    source_doc = Document(str(target))
    formatted = Document(str(output))
    output_body = next(p for p in formatted.paragraphs if "横向页正文" in p.text)
    assert output_body.runs[0].font.size == Pt(10.5)
    assert formatted.tables[0]._tbl.xml == source_doc.tables[0]._tbl.xml
    assert result["counts"]["body"] > 0
    assert result["counts"]["skipped_tables"] == 1
    assert result["counts"]["table_paragraphs"] == 0


def test_format_scope_table_emphasis_does_not_change_fonts_or_sizes():
    d = ws_dir("format_scope_emphasis")
    template = _template(d / "样板.docx")
    target = _messy(d / "乱版.docx")
    output = d / "仅表格强调.docx"
    scope = _scope(cover=False, headings=False, statement=False, body=False, table_text=False, table_emphasis=True, table_borders=False)
    api.run_report_format(target, template, output=output, scope=scope)
    source_doc = Document(str(target))
    formatted = Document(str(output))
    for row in range(3):
        for col in range(2):
            before = source_doc.tables[0].cell(row, col).paragraphs[0]
            after = formatted.tables[0].cell(row, col).paragraphs[0]
            assert [run.font.size for run in after.runs] == [run.font.size for run in before.runs]
            assert [run.font.name for run in after.runs] == [run.font.name for run in before.runs]
    assert formatted.tables[0].cell(0, 0).paragraphs[0].alignment == 1
    assert formatted.tables[0].cell(0, 0).paragraphs[0].runs[0].font.bold is True
    assert formatted.tables[0].cell(1, 1).paragraphs[0].alignment == 2
    assert formatted.tables[0].cell(2, 0).paragraphs[0].runs[0].font.bold is True


def test_format_scope_preserves_abnormal_outline_paragraph():
    d = ws_dir("format_scope_protected")
    template = _template(d / "样板.docx")
    target = _messy(d / "乱版.docx")
    source_doc = Document(str(target))
    protected = next(p for p in source_doc.paragraphs if "保留文字和审计痕迹" in p.text)
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), "9")
    protected._p.get_or_add_pPr().append(outline)
    source_doc.save(str(target))
    protected_xml = protected._p.xml
    output = d / "保护异常段落.docx"
    result = api.run_report_format(target, template, output=output)
    formatted = Document(str(output))
    output_protected = next(p for p in formatted.paragraphs if "保留文字和审计痕迹" in p.text)
    assert output_protected._p.xml == protected_xml
    assert result["totals"]["protected_paragraphs"] == 1
    assert result["counts"]["protected_paragraphs"] == 1


@pytest.mark.parametrize("scope", [
    {},
    {**DEFAULT_SCOPE, "unknown": True},
    {**DEFAULT_SCOPE, "body": 1},
    {key: False for key in DEFAULT_SCOPE},
])
def test_format_scope_rejects_invalid_selection(scope):
    d = ws_dir("format_scope_invalid")
    template = _template(d / "样板.docx")
    target = _messy(d / "乱版.docx")
    with pytest.raises(SyncError, match="调整项目|至少选择"):
        api.scan_report_format(target, template, scope=scope)


def test_format_apply_rejects_existing_output_and_same_file():
    d = ws_dir("format_output")
    template = _template(d / "样板.docx")
    target = _messy(d / "乱版.docx")
    with pytest.raises(SyncError, match="只生成新副本"):
        api.run_report_format(target, template, output=target)
    existing = d / "已存在.docx"
    existing.write_bytes(b"keep")
    with pytest.raises(SyncError, match="已存在"):
        api.run_report_format(target, template, output=existing)
    assert existing.read_bytes() == b"keep"


def _docm_with_vba(path: Path) -> Path:
    doc = Document()
    doc.add_paragraph("宏报告正文")
    doc.save(str(path))
    with zipfile.ZipFile(path) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    files["word/vbaProject.bin"] = b"fake-vba-project"
    rel_path = "word/_rels/document.xml.rels"
    rel_xml = files[rel_path].decode("utf-8")
    rel_xml = rel_xml.replace(
        "</Relationships>",
        '<Relationship Id="rIdVba1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
    )
    files[rel_path] = rel_xml.encode("utf-8")
    ct_path = "[Content_Types].xml"
    ct_xml = files[ct_path].decode("utf-8")
    ct_xml = ct_xml.replace(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "application/vnd.ms-word.document.macroEnabled.main+xml",
    )
    ct_xml = ct_xml.replace(
        "</Types>",
        '<Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
    )
    files[ct_path] = ct_xml.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in files.items():
            archive.writestr(name, data)
    return path


def test_saved_format_profile_is_content_free_and_reusable_without_template():
    d = ws_dir("format_profile")
    template = _template(d / "样板.docx")
    target = _messy(d / "乱版.docx")
    saved = api.save_format_profile(template, "纯格式测试")
    profile_path = Path(saved["path"])
    raw = profile_path.read_text(encoding="utf-8")
    data = json.loads(raw)
    assert str(template) not in raw
    assert template.name not in raw
    assert "这是样板正文" not in raw
    assert "金额" not in raw
    assert "100.00" not in raw
    assert "template" not in data
    assert "table_title" in data["paragraphs"]
    assert api.list_format_profiles()[0]["name"] == "纯格式测试"
    with pytest.raises(SyncError, match="已存在"):
        api.save_format_profile(template, "纯格式测试")

    template.unlink()
    preview = api.scan_report_format(target, profile_name="纯格式测试")
    assert preview["template"] == "已存方案：纯格式测试"
    before = _document_snapshot(Document(str(target)))
    output = d / "乱版_方案格式调整.docx"
    result = api.run_report_format(target, output=output, profile_name="纯格式测试")
    assert result["applied"] is True
    assert _document_snapshot(Document(str(output))) == before

    assert api.delete_format_profile("纯格式测试")["deleted"] is True
    assert api.list_format_profiles() == []
    with pytest.raises(SyncError, match="找不到格式方案"):
        api.load_format_profile("纯格式测试")


def test_legacy_format_profile_preserves_borders_when_field_is_missing():
    d = ws_dir("format_profile_legacy_borders")
    template = _template(d / "样板.docx")
    target = _messy(d / "乱版.docx")
    saved = api.save_format_profile(template, "旧框线方案")
    profile_path = Path(saved["path"])
    data = json.loads(profile_path.read_text(encoding="utf-8"))
    del data["table_borders"]
    profile_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    scope = _scope(cover=False, headings=False, statement=False, body=False, table_text=False, table_emphasis=False, table_borders=True)
    preview = api.scan_report_format(target, profile_name="旧框线方案", scope=scope)
    assert preview["table_border_format"] == "旧方案未保存框线数据，保持原样"
    output = d / "旧方案输出.docx"
    result = api.run_report_format(target, output=output, profile_name="旧框线方案", scope=scope)
    assert "table_borders" not in result["counts"]
    assert _border_snapshot(Document(str(output)).tables[0]) == _border_snapshot(Document(str(target)).tables[0])


def test_saved_notes_profile_uses_source_region_and_loads_legacy_continuation():
    d = ws_dir("format_profile_legacy_continuation")
    template = _notes_template(d / "完整样板.docx")
    target = _notes_target(d / "附注.docx")
    saved = api.save_format_profile(template, "旧附注方案", source=target)
    profile_path = Path(saved["path"])
    data = json.loads(profile_path.read_text(encoding="utf-8"))
    assert data["paragraphs"]["cover_title"]["size"] == 16
    assert data["paragraphs"]["body"]["size"] == 10.5
    assert str(template) not in profile_path.read_text(encoding="utf-8")
    del data["paragraphs"]["table_continuation"]
    profile_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    template.unlink()

    output = d / "旧附注方案输出.docx"
    before = _document_snapshot(Document(str(target)))
    result = api.run_report_format(target, output=output, profile_name="旧附注方案")
    assert result["applied"] is True
    assert _document_snapshot(Document(str(output))) == before
    continuation = next(p for p in Document(str(output)).paragraphs if p.text == "（续上表）")
    assert continuation.runs[0].font.size == Pt(10.5)
    assert continuation.alignment == 2
    assert api.delete_format_profile("旧附注方案")["deleted"] is True


def test_notes_source_without_notes_template_stops_instead_of_using_whole_template():
    d = ws_dir("format_notes_missing_region")
    template = _template(d / "无附注样板.docx")
    target = _notes_target(d / "附注.docx")
    with pytest.raises(SyncError, match="未找到.*附注区"):
        api.scan_report_format(target, template)


@pytest.mark.parametrize("edge_values", [
    {"top": {"val": "single", "size": 1, "color": None, "extra": 1}},
    {"top": {"val": "invalid", "size": 1, "color": None}},
    {"top": {"val": "single", "size": True, "color": None}},
    {"top": {"val": "single", "size": 1, "color": "red"}},
])
def test_corrupt_format_profile_borders_are_rejected(edge_values):
    d = ws_dir("format_profile_bad_borders")
    template = _template(d / "样板.docx")
    saved = api.save_format_profile(template, "损坏框线方案")
    profile_path = Path(saved["path"])
    data = json.loads(profile_path.read_text(encoding="utf-8"))
    data["table_borders"].update(edge_values)
    profile_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(SyncError, match="框线"):
        api.load_format_profile("损坏框线方案")


def test_corrupt_format_profile_is_reported():
    d = ws_dir("format_profile_corrupt")
    template = _template(d / "样板.docx")
    saved = api.save_format_profile(template, "损坏方案")
    Path(saved["path"]).write_text("{不是JSON", encoding="utf-8")
    with pytest.raises(SyncError, match="无法读取"):
        api.load_format_profile("损坏方案")
    with pytest.raises(SyncError, match="已停止加载"):
        api.list_format_profiles()


def test_format_apply_preserves_docm_vba():
    d = ws_dir("format_docm")
    template = _template(d / "样板.docx")
    target = _docm_with_vba(d / "宏报告.docm")
    output = d / "宏报告_格式调整.docm"
    api.run_report_format(target, template, output=output)
    with zipfile.ZipFile(target) as source, zipfile.ZipFile(output) as result:
        assert result.read("word/vbaProject.bin") == source.read("word/vbaProject.bin") == b"fake-vba-project"
