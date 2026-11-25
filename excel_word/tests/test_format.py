import base64
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
from wesync.report_format import DEFAULT_SCOPE, _document_snapshot


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
    title = doc.add_paragraph("审计报告")
    title.alignment = 1
    for run in title.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(26)
        run.font.bold = True
    doc.add_paragraph("目录")
    _heading(doc, "一、审计意见", 1)
    body = doc.add_paragraph("这是样板正文，用于提取字体和段落格式。")
    body.paragraph_format.first_line_indent = Pt(21)
    body.paragraph_format.line_spacing = Pt(20)
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    table.cell(0, 0).text = "项目"
    table.cell(0, 1).text = "金额"
    table.cell(1, 0).text = "合计"
    table.cell(1, 1).text = "100.00"
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
    _heading(doc, "一、审计意见", 1)
    body = doc.add_paragraph()
    run = body.add_run("保留文字和审计痕迹")
    run.font.highlight_color = 7
    run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    run.font.underline = True
    _bookmark(body, "CPAH_KEEP", 1)
    _hyperlink(body, "链接文字", "https://example.invalid/audit")

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
    assert output.exists()
    assert target.read_bytes() == before_bytes
    assert _document_snapshot(Document(str(output))) == before

    formatted = Document(str(output))
    body = next(paragraph for paragraph in formatted.paragraphs if "保留文字和审计痕迹" in paragraph.text)
    traced = next(run for run in body.runs if run.text == "保留文字和审计痕迹")
    assert traced.font.highlight_color is not None
    assert traced.font.underline is True
    assert traced.font.color.rgb == RGBColor(0xC0, 0x00, 0x00)
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
    scope = _scope(cover=False, headings=False, statement=False, body=False, table_text=True, table_emphasis=False)
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
    scope = _scope(cover=False, headings=False, statement=False, body=True, table_text=False, table_emphasis=False)
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
    scope = _scope(cover=False, headings=False, statement=False, body=False, table_text=False, table_emphasis=True)
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
