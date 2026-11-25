"""审计报告格式方案：从样板提取角色格式，预演或生成格式调整副本。"""

from __future__ import annotations

import hashlib
import json
import math
import os
import re
import tempfile
import zipfile
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from pathlib import Path
from weakref import WeakKeyDictionary

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.styles import BabelFish
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .config import DATE_RE
from .console import log
from .errors import SyncError
from .fsops import project_root, resolve_docx, tmp_root


DOCX_MAIN_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
DOCM_MAIN_TYPE = "application/vnd.ms-word.document.macroEnabled.main+xml"


def _rewrite_main_type(path: Path, macro_enabled: bool) -> None:
    with zipfile.ZipFile(path) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    content = files["[Content_Types].xml"].decode("utf-8")
    if macro_enabled:
        content = content.replace(DOCX_MAIN_TYPE, DOCM_MAIN_TYPE)
    else:
        content = content.replace(DOCM_MAIN_TYPE, DOCX_MAIN_TYPE)
    files["[Content_Types].xml"] = content.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in files.items():
            archive.writestr(name, data)


def _open_document(path: Path):
    if path.suffix.lower() != ".docm":
        return Document(str(path))
    fd, temp_name = tempfile.mkstemp(prefix="wesync_docm_", suffix=".docx", dir=tmp_root())
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(path) as source, zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
            for name in source.namelist():
                data = source.read(name)
                if name == "[Content_Types].xml":
                    data = data.replace(DOCM_MAIN_TYPE.encode(), DOCX_MAIN_TYPE.encode())
                target.writestr(name, data)
        return Document(str(temp_path))
    finally:
        temp_path.unlink(missing_ok=True)


ROLE_LABELS = {
    "cover_title": "封面标题",
    "cover_info": "封面信息",
    "toc_title": "目录标题",
    "toc_entry": "目录条目",
    "heading_1": "一级标题",
    "heading_2": "二级标题",
    "heading_3": "三级标题",
    "heading_4": "四级标题",
    "statement": "报表名称/日期",
    "table_title": "表格标题",
    "body": "正文",
    "table_header": "表头",
    "table_total": "合计行",
    "table_data": "表格数据",
}

SCOPE_LABELS = {
    "cover": "封面与目录",
    "headings": "一至四级标题",
    "statement": "报表名称、日期与表格标题",
    "body": "正文段落",
    "table_text": "表格字体与字号",
    "table_emphasis": "表格对齐、表头与合计行加粗",
}
DEFAULT_SCOPE = {key: True for key in SCOPE_LABELS}
PARAGRAPH_SCOPE_ROLES = {
    "cover": ("cover_title", "cover_info", "toc_title", "toc_entry"),
    "headings": ("heading_1", "heading_2", "heading_3", "heading_4"),
    "statement": ("statement", "table_title"),
    "body": ("body",),
}

ALIGN_VALUES = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "distribute": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

ALIGN_LABELS = {
    None: "默认",
    "left": "左对齐",
    "center": "居中",
    "right": "右对齐",
    "both": "两端对齐",
    "distribute": "分散对齐",
}


@dataclass
class ParagraphFormat:
    ascii_font: str | None = None
    east_asia_font: str | None = None
    size: float | None = None
    bold: bool | None = None
    align: str | None = None
    line_value: float | None = None
    line_rule: str | None = None
    before: float | None = None
    after: float | None = None
    first_line: float | None = None
    left: float | None = None


@dataclass
class FormatProfile:
    paragraphs: dict[str, ParagraphFormat] = field(default_factory=dict)
    table_ascii_font: str = "Times New Roman"
    table_east_asia_font: str = "宋体"
    table_size: float = 9.0
    template: str = ""


def _profiles_dir() -> Path:
    path = project_root() / "format_profiles"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _validated_scope(scope: dict | None) -> dict[str, bool]:
    if scope is None:
        return dict(DEFAULT_SCOPE)
    if not isinstance(scope, dict) or set(scope) != set(DEFAULT_SCOPE):
        raise SyncError("调整项目必须完整包含六项，且不能包含未知项目")
    if any(type(scope[key]) is not bool for key in DEFAULT_SCOPE):
        raise SyncError("调整项目的选择必须为“是”或“否”")
    if not any(scope.values()):
        raise SyncError("请至少选择一项需要调整的项目")
    return {key: scope[key] for key in DEFAULT_SCOPE}


def _selected_paragraph_roles(scope: dict[str, bool]) -> set[str]:
    return {
        role
        for key, enabled in scope.items()
        if enabled and key in PARAGRAPH_SCOPE_ROLES
        for role in PARAGRAPH_SCOPE_ROLES[key]
    }


def _profile_path(name: str) -> Path:
    name = name.strip()
    if (
        not name
        or len(name) > 80
        or name.endswith((".", " "))
        or re.search(r'[<>:"/\\|?*\x00-\x1f]', name)
        or re.match(r"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)", name, re.IGNORECASE)
    ):
        raise SyncError("方案名称不能为空，且不能超过 80 个字符或包含文件名禁用字符")
    return _profiles_dir() / f"{name}.json"


def _profile_data(profile: FormatProfile, name: str) -> dict:
    return {
        "version": 1,
        "name": name,
        "paragraphs": {role: asdict(fmt) for role, fmt in profile.paragraphs.items()},
        "table_ascii_font": profile.table_ascii_font,
        "table_east_asia_font": profile.table_east_asia_font,
        "table_size": profile.table_size,
    }


def _validated_profile(data: dict) -> FormatProfile:
    if not isinstance(data, dict) or data.get("version") != 1:
        raise SyncError("格式方案文件版本不受支持")
    paragraphs = data.get("paragraphs")
    if not isinstance(paragraphs, dict):
        raise SyncError("格式方案缺少段落格式数据")
    fields = set(ParagraphFormat.__dataclass_fields__)
    parsed = {}
    for role, values in paragraphs.items():
        if role not in ROLE_LABELS or role in ("table_header", "table_total", "table_data") or not isinstance(values, dict):
            raise SyncError("格式方案包含不支持的段落角色")
        if set(values) - fields:
            raise SyncError("格式方案包含不支持的格式字段")
        try:
            fmt = ParagraphFormat(**values)
        except TypeError as exc:
            raise SyncError("格式方案字段类型不正确") from exc
        for key in ("ascii_font", "east_asia_font", "align", "line_rule"):
            value = getattr(fmt, key)
            if value is not None and not isinstance(value, str):
                raise SyncError("格式方案字体或段落属性类型不正确")
        if fmt.align is not None and fmt.align not in ALIGN_LABELS:
            raise SyncError("格式方案段落对齐属性不正确")
        if fmt.line_rule is not None and fmt.line_rule not in ("auto", "exact", "atLeast"):
            raise SyncError("格式方案行距属性不正确")
        for key in ("size", "line_value", "before", "after", "first_line", "left"):
            value = getattr(fmt, key)
            if value is not None and (isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value < 0):
                raise SyncError("格式方案数值属性不正确")
        if fmt.bold is not None and not isinstance(fmt.bold, bool):
            raise SyncError("格式方案加粗属性类型不正确")
        parsed[role] = fmt
    required_roles = set(ROLE_LABELS) - {"table_header", "table_total", "table_data"}
    if set(parsed) != required_roles:
        raise SyncError("格式方案段落角色不完整")
    fonts = (data.get("table_ascii_font"), data.get("table_east_asia_font"))
    size = data.get("table_size")
    if any(not isinstance(font, str) or not font for font in fonts):
        raise SyncError("格式方案缺少表格字体")
    if isinstance(size, bool) or not isinstance(size, (int, float)) or not math.isfinite(size) or size <= 0:
        raise SyncError("格式方案表格字号不正确")
    return FormatProfile(parsed, fonts[0], fonts[1], float(size))


def list_format_profiles() -> list[dict]:
    result = []
    for path in sorted(_profiles_dir().glob("*.json")):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            name = data.get("name")
            if not isinstance(name, str) or not name:
                raise ValueError("缺少方案名称")
            result.append({"name": name, "path": str(path), "size": path.stat().st_size})
        except (OSError, ValueError, AttributeError) as exc:
            raise SyncError(f"格式方案文件无法读取，已停止加载：{path.name}：{exc}") from exc
    return result


def save_format_profile(template: str | Path, name: str) -> dict:
    path = _profile_path(name)
    if path.exists():
        raise SyncError(f"格式方案已存在，未覆盖：{name}")
    profile = build_profile(template)
    data = _profile_data(profile, name.strip())
    fd, temp_name = tempfile.mkstemp(prefix=f".{path.stem}.", suffix=".json", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
        os.replace(temp_name, path)
    except Exception:
        Path(temp_name).unlink(missing_ok=True)
        raise
    log(f"[格式] 已保存纯格式方案：{path}（不含报告文字、表格内容或样板路径）")
    return {"name": name.strip(), "path": str(path)}


def load_format_profile(name: str) -> FormatProfile:
    path = _profile_path(name)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise SyncError(f"找不到格式方案：{name}") from exc
    except (OSError, ValueError) as exc:
        raise SyncError(f"格式方案无法读取：{name}：{exc}") from exc
    if not isinstance(data, dict) or data.get("name") != name.strip():
        raise SyncError("格式方案文件名称不一致")
    return _validated_profile(data)


def delete_format_profile(name: str) -> dict:
    path = _profile_path(name)
    try:
        path.unlink()
    except FileNotFoundError as exc:
        raise SyncError(f"找不到格式方案：{name}") from exc
    log(f"[格式] 已删除格式方案：{name}")
    return {"name": name.strip(), "deleted": True}


def _number(element, name: str, scale: float = 1.0) -> float | None:
    if element is None:
        return None
    value = element.get(qn(f"w:{name}"))
    if value is None or not value.lstrip("-").isdigit():
        return None
    return int(value) / scale


def _on_off(element) -> bool | None:
    if element is None:
        return None
    return element.get(qn("w:val")) not in ("0", "false", "off")


_STYLE_CACHE: WeakKeyDictionary = WeakKeyDictionary()


def _style_data(paragraph: Paragraph):
    part = paragraph.part
    cached = _STYLE_CACHE.get(part)
    if cached is not None:
        return cached
    styles = part.styles
    by_id = {style.style_id: style for style in styles}
    defaults = {
        WD_STYLE_TYPE.PARAGRAPH: styles.default(WD_STYLE_TYPE.PARAGRAPH),
        WD_STYLE_TYPE.CHARACTER: styles.default(WD_STYLE_TYPE.CHARACTER),
    }
    cached = (by_id, defaults, {})
    _STYLE_CACHE[part] = cached
    return cached


def _paragraph_style(paragraph: Paragraph):
    by_id, defaults, _ = _style_data(paragraph)
    style_id = paragraph._p.style
    style = by_id.get(style_id)
    if style is None or style.type != WD_STYLE_TYPE.PARAGRAPH:
        style = defaults[WD_STYLE_TYPE.PARAGRAPH]
    return style


def _run_style(paragraph: Paragraph, run: Run):
    by_id, defaults, _ = _style_data(paragraph)
    style_id = run._r.rPr.rStyle.val if run._r.rPr is not None and run._r.rPr.rStyle is not None else None
    style = by_id.get(style_id)
    if style is None or style.type != WD_STYLE_TYPE.CHARACTER:
        style = defaults[WD_STYLE_TYPE.CHARACTER]
    return style


def _style_chain(paragraph: Paragraph, style) -> list:
    _, _, chains = _style_data(paragraph)
    key = style.style_id if style is not None else None
    if key not in chains:
        result = []
        seen = set()
        while style is not None and style.style_id not in seen:
            seen.add(style.style_id)
            result.append(style.element)
            style = style.base_style
        chains[key] = result
    return chains[key]


def _find_property(paragraph: Paragraph, name: str, run: Run | None = None):
    containers = []
    if run is not None:
        containers.append(run._r.rPr)
        containers.extend(_style_chain(paragraph, _run_style(paragraph, run)))
    containers.append(paragraph._p.pPr)
    containers.extend(_style_chain(paragraph, _paragraph_style(paragraph)))
    for element in containers:
        if element is None:
            continue
        container = element
        if element.tag == qn("w:style") and name in ("spacing", "ind", "jc", "outlineLvl"):
            container = element.find(qn("w:pPr"))
            if container is None:
                continue
        found = container.find(qn(f"w:{name}"))
        if found is not None:
            return found
    return None


def _effective_paragraph_format(paragraph: Paragraph) -> ParagraphFormat:
    run = next((item for item in paragraph.runs if item.text.strip()), None)
    fonts = _find_property(paragraph, "rFonts", run)
    ascii_font = None
    east_asia_font = None
    if fonts is not None:
        ascii_font = fonts.get(qn("w:ascii")) or fonts.get(qn("w:hAnsi"))
        east_asia_font = fonts.get(qn("w:eastAsia"))
    spacing = _find_property(paragraph, "spacing")
    line_rule = spacing.get(qn("w:lineRule")) if spacing is not None else None
    line_value = _number(spacing, "line")
    if line_rule == "auto" and line_value is not None:
        line_value /= 240.0
    else:
        line_value = _number(spacing, "line", 20.0)
    indent = _find_property(paragraph, "ind")
    alignment = _find_property(paragraph, "jc")
    return ParagraphFormat(
        ascii_font=ascii_font,
        east_asia_font=east_asia_font,
        size=_number(_find_property(paragraph, "sz", run), "val", 2.0),
        bold=_on_off(_find_property(paragraph, "b", run)),
        align=alignment.get(qn("w:val")) if alignment is not None else None,
        line_value=line_value,
        line_rule=line_rule,
        before=_number(spacing, "before", 20.0),
        after=_number(spacing, "after", 20.0),
        first_line=_number(indent, "firstLine", 20.0),
        left=_number(indent, "left", 20.0),
    )


def _effective_run_size(paragraph: Paragraph, run: Run) -> float | None:
    return _number(_find_property(paragraph, "sz", run), "val", 2.0)


def _outline_level(paragraph: Paragraph) -> int | None:
    element = _find_property(paragraph, "outlineLvl")
    value = element.get(qn("w:val")) if element is not None else None
    if value is None or not value.lstrip("-").isdigit():
        return None
    return int(value)


def _style_heading_level(paragraph: Paragraph) -> int | None:
    style = _paragraph_style(paragraph)
    internal_name = style.element.name_val if style is not None else None
    name = BabelFish.internal2ui(internal_name) if internal_name else ""
    match = re.search(r"(?:标题|heading)\s*\[?\s*([1-5])", name, re.IGNORECASE)
    if match:
        return int(match.group(1))
    match = re.search(r"([一二三四五])级标题", name)
    if match:
        return "一二三四五".index(match.group(1)) + 1
    return None


def _heading_role(paragraph: Paragraph) -> str | None:
    level = _outline_level(paragraph)
    if level is not None and 0 <= level <= 4:
        return "statement" if level == 4 else f"heading_{level + 1}"
    if level is not None and level > 4:
        return None
    level = _style_heading_level(paragraph)
    if level is None:
        return None
    return "statement" if level == 5 else f"heading_{level}"


def _body_items(doc):
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def _normalized_text(paragraph: Paragraph) -> str:
    return re.sub(r"\s+", "", paragraph.text)


def _classify_paragraphs(doc) -> tuple[list[tuple[Paragraph, str]], dict[str, int]]:
    paragraphs = [item for item in _body_items(doc) if isinstance(item, Paragraph)]
    nonempty = [item for item in paragraphs if item.text.strip()]
    heading_roles = [_heading_role(item) for item in nonempty]
    heading_indexes = [index for index, role in enumerate(heading_roles) if role]
    toc_index = next((index for index, item in enumerate(nonempty) if _normalized_text(item) == "目录"), None)
    first_heading = heading_indexes[0] if heading_indexes else len(nonempty)
    if toc_index is not None:
        cover_end = toc_index
    else:
        cover_end = min(first_heading, 3)
    toc_end = len(nonempty)
    if toc_index is not None:
        later = [index for index in heading_indexes if index > toc_index]
        toc_end = later[0] if later else len(nonempty)

    table_titles: set[int] = set()
    previous: Paragraph | None = None
    previous_index = -1
    position = -1
    for item in _body_items(doc):
        if isinstance(item, Paragraph):
            if item.text.strip():
                position += 1
                previous = item
                previous_index = position
            continue
        if previous is None or not previous.text.strip() or heading_roles[previous_index]:
            continue
        if DATE_RE.match(previous.text.strip()):
            continue
        index = previous_index
        if index < cover_end:
            continue
        if toc_index is not None and toc_index <= index < toc_end:
            continue
        table_titles.add(index)

    result: list[tuple[Paragraph, str]] = []
    ambiguous = 0
    for index, paragraph in enumerate(nonempty):
        role = heading_roles[index]
        if role is None and (_outline_level(paragraph) or 0) > 4:
            ambiguous += 1
        if role is None and index in table_titles:
            role = "table_title"
        if role is None and toc_index is not None:
            if index == toc_index:
                role = "toc_title"
            elif toc_index < index < toc_end:
                role = "toc_entry"
        if role is None and index < cover_end:
            role = "cover_title" if index < min(3, cover_end) else "cover_info"
        result.append((paragraph, role or "body"))
    return result, {"ambiguous": ambiguous, "empty": len(paragraphs) - len(nonempty)}


def _mode(values, *, ignore_none: bool = False):
    usable = [value for value in values if not ignore_none or value is not None]
    if not usable:
        return None
    return Counter(usable).most_common(1)[0][0]


def _default_formats() -> dict[str, ParagraphFormat]:
    body = ParagraphFormat("Times New Roman", "宋体", 10.5, None, "left", 20.0, "exact", 0.0, 0.0, 21.0, 0.0)
    return {
        "cover_title": ParagraphFormat("Times New Roman", "宋体", 26.0, True, "center", 47.0, "atLeast", 0.0, 0.0, 0.0, 0.0),
        "cover_info": ParagraphFormat("Times New Roman", "宋体", 14.0, None, "center", 30.0, "atLeast", 0.0, 0.0, 0.0, 0.0),
        "toc_title": ParagraphFormat("Times New Roman", "宋体", 18.0, True, "center", 24.0, "exact", 0.0, 0.0, 0.0, 0.0),
        "toc_entry": ParagraphFormat("Times New Roman", "宋体", 12.0, False, "left", 40.0, "atLeast", 0.0, 0.0, 0.0, 26.25),
        "heading_1": ParagraphFormat("Times New Roman", "宋体", 12.0, True, "left", 20.0, "exact", 12.0, 0.0, 24.1, 0.0),
        "heading_2": ParagraphFormat("Times New Roman", "宋体", 10.5, True, "left", 20.0, "exact", 6.0, 0.0, 21.1, 0.0),
        "heading_3": ParagraphFormat("Times New Roman", "宋体", 10.5, False, "left", 18.0, "exact", 6.0, 0.0, 21.1, 0.0),
        "heading_4": ParagraphFormat("Times New Roman", "宋体", 10.5, False, "left", 18.0, "exact", 6.0, 0.0, 31.5, 0.0),
        "statement": ParagraphFormat("Times New Roman", "宋体", 10.5, False, "left", 18.0, "exact", 6.0, 0.0, 31.5, 0.0),
        "table_title": ParagraphFormat("Times New Roman", "宋体", 9.0, True, "center", 18.0, "exact", 6.0, 0.0, 0.0, 0.0),
        "body": body,
    }


def _merge_format(default: ParagraphFormat, samples: list[ParagraphFormat]) -> ParagraphFormat:
    if not samples:
        return default
    bold_values = [item.bold for item in samples]
    bold = _mode(bold_values)
    if bold is None and default.bold is not None:
        bold = default.bold
    return ParagraphFormat(
        ascii_font=_mode([item.ascii_font for item in samples], ignore_none=True) or default.ascii_font,
        east_asia_font=_mode([item.east_asia_font for item in samples], ignore_none=True) or default.east_asia_font,
        size=_mode([item.size for item in samples], ignore_none=True) or default.size,
        bold=bold,
        align=_mode([item.align for item in samples], ignore_none=True) or default.align,
        line_value=_mode([item.line_value for item in samples], ignore_none=True) or default.line_value,
        line_rule=_mode([item.line_rule for item in samples if item.line_value is not None], ignore_none=True) or default.line_rule,
        before=_mode([item.before for item in samples]) if any(item.before is not None for item in samples) else default.before,
        after=_mode([item.after for item in samples]) if any(item.after is not None for item in samples) else default.after,
        first_line=_mode([item.first_line for item in samples], ignore_none=True) or default.first_line,
        left=_mode([item.left for item in samples], ignore_none=True) or default.left,
    )


def _table_run_formats(doc) -> tuple[Counter, Counter]:
    fonts = Counter()
    sizes = Counter()
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    run = next((item for item in paragraph.runs if item.text.strip()), None)
                    if run is None:
                        continue
                    fmt = _effective_paragraph_format(paragraph)
                    if fmt.ascii_font:
                        fonts[fmt.ascii_font] += len(run.text)
                    if fmt.east_asia_font:
                        fonts[fmt.east_asia_font] += len(run.text)
                    if fmt.size is not None:
                        sizes[fmt.size] += len(run.text)
    return fonts, sizes


def build_profile(template: str | Path) -> FormatProfile:
    path = resolve_docx(str(template))
    doc = _open_document(path)
    classified, _ = _classify_paragraphs(doc)
    grouped: dict[str, list[ParagraphFormat]] = defaultdict(list)
    for paragraph, role in classified:
        grouped[role].append(_effective_paragraph_format(paragraph))
    formats = _default_formats()
    for role, samples in grouped.items():
        formats[role] = _merge_format(formats.get(role, ParagraphFormat()), samples)
    fonts, sizes = _table_run_formats(doc)
    ascii_fonts = [name for name in fonts if name not in ("宋体", "SimSun", "微软雅黑")]
    return FormatProfile(
        paragraphs=formats,
        table_ascii_font=_mode(ascii_fonts) or "Times New Roman",
        table_east_asia_font=_mode([name for name in fonts if name in ("宋体", "SimSun")]) or "宋体",
        table_size=_mode([size for size in sizes if size <= 9.0]) or _mode(list(sizes)) or 9.0,
        template=str(path),
    )


def _suggested_output(source: Path) -> Path:
    candidate = source.with_name(f"{source.stem}_格式调整{source.suffix}")
    counter = 2
    while candidate.exists():
        candidate = source.with_name(f"{source.stem}_格式调整_{counter}{source.suffix}")
        counter += 1
    return candidate


def _format_summary(fmt: ParagraphFormat) -> str:
    font = " / ".join(dict.fromkeys(item for item in (fmt.ascii_font, fmt.east_asia_font) if item))
    parts = [font or "继承字体"]
    if fmt.size is not None:
        parts.append(f"{fmt.size:g}磅")
    if fmt.bold is not None:
        parts.append("加粗" if fmt.bold else "不加粗")
    parts.append(ALIGN_LABELS.get(fmt.align, fmt.align or "默认对齐"))
    if fmt.line_value is not None:
        if fmt.line_rule == "auto":
            parts.append(f"{fmt.line_value:g}倍行距")
        else:
            rule = "固定行距" if fmt.line_rule == "exact" else "最小行距"
            parts.append(f"{rule}{fmt.line_value:g}磅")
    if fmt.first_line:
        parts.append(f"首行缩进{fmt.first_line:g}磅")
    if fmt.left:
        parts.append(f"左缩进{fmt.left:g}磅")
    return "；".join(parts)


def _table_row_roles(table: Table) -> list[str]:
    roles = []
    header_seen = False
    for row in table.rows:
        first_text = next((cell.text.strip() for cell in row.cells if cell.text.strip()), "")
        if first_text.startswith(("合计", "小计", "总计")):
            roles.append("table_total")
        elif not header_seen and any(cell.text.strip() for cell in row.cells):
            roles.append("table_header")
            header_seen = True
        else:
            roles.append("table_data")
    return roles


def _paragraph_runs(paragraph: Paragraph) -> list[Run]:
    return [Run(element, paragraph) for element in paragraph._p.xpath(".//w:r") if element.xpath(".//w:t")]


def _set_run_font(run: Run, fmt: ParagraphFormat, *, size: float | None = None, bold: bool | None = None) -> None:
    if fmt.ascii_font:
        run.font.name = fmt.ascii_font
        rpr = run._r.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        rfonts.set(qn("w:ascii"), fmt.ascii_font)
        rfonts.set(qn("w:hAnsi"), fmt.ascii_font)
        if fmt.east_asia_font:
            rfonts.set(qn("w:eastAsia"), fmt.east_asia_font)
        rfonts.set(qn("w:cs"), fmt.ascii_font)
    target_size = fmt.size if size is None else size
    if target_size is not None:
        run.font.size = Pt(target_size)
    target_bold = fmt.bold if bold is None else bold
    if target_bold is not None:
        run.font.bold = target_bold


def _apply_paragraph_format(paragraph: Paragraph, fmt: ParagraphFormat) -> None:
    for run in _paragraph_runs(paragraph):
        _set_run_font(run, fmt)
    pf = paragraph.paragraph_format
    if fmt.align in ALIGN_VALUES:
        pf.alignment = ALIGN_VALUES[fmt.align]
    if fmt.line_value is not None:
        pf.line_spacing = fmt.line_value if fmt.line_rule == "auto" else Pt(fmt.line_value)
    pf.space_before = Pt(fmt.before or 0.0)
    pf.space_after = Pt(fmt.after or 0.0)
    pf.first_line_indent = Pt(fmt.first_line or 0.0)
    pf.left_indent = Pt(fmt.left or 0.0)


def _looks_numeric(text: str) -> bool:
    value = text.strip().replace(",", "").replace("，", "")
    if value in ("", "-", "—", "–", "（", "）", "()"):
        return False
    return bool(re.fullmatch(r"[￥¥$]?-?\d[\d,]*(?:\.\d+)?%?\)?", value))


def _apply_table_format(table: Table, profile: FormatProfile, scope: dict[str, bool]) -> tuple[int, Counter]:
    apply_text = scope["table_text"]
    apply_emphasis = scope["table_emphasis"]
    if not apply_text and not apply_emphasis:
        return 0, Counter()
    base = ParagraphFormat(profile.table_ascii_font, profile.table_east_asia_font, profile.table_size, None)
    formatted = 0
    roles = _table_row_roles(table)
    for row_index, row in enumerate(table.rows):
        role = roles[row_index]
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                if not paragraph.text.strip():
                    continue
                if apply_emphasis:
                    align = "center" if role == "table_header" else ("right" if _looks_numeric(paragraph.text) else "left")
                    paragraph.alignment = ALIGN_VALUES[align]
                for run in _paragraph_runs(paragraph):
                    if not apply_text:
                        if role in ("table_header", "table_total"):
                            run.font.bold = True
                        continue
                    current_size = _effective_run_size(paragraph, run)
                    size = min(current_size, profile.table_size) if current_size is not None else profile.table_size
                    bold = True if apply_emphasis and role in ("table_header", "table_total") else None
                    _set_run_font(run, base, size=size, bold=bold)
                formatted += 1
    return formatted, Counter(roles)


def _apply_profile(doc, profile: FormatProfile, scope: dict[str, bool]) -> dict[str, int]:
    counts: Counter[str] = Counter()
    selected_roles = _selected_paragraph_roles(scope)
    classified, meta = _classify_paragraphs(doc)
    skipped_paragraphs = 0
    protected_paragraphs = 0
    for paragraph, role in classified:
        if (_outline_level(paragraph) or 0) > 4:
            protected_paragraphs += 1
            continue
        if role not in selected_roles:
            skipped_paragraphs += 1
            continue
        fmt = profile.paragraphs.get(role) or profile.paragraphs["body"]
        _apply_paragraph_format(paragraph, fmt)
        counts[role] += 1
    table_paragraphs = 0
    skipped_tables = 0
    table_enabled = scope["table_text"] or scope["table_emphasis"]
    for table in doc.tables:
        if not table_enabled:
            skipped_tables += 1
            continue
        formatted, roles = _apply_table_format(table, profile, scope)
        table_paragraphs += formatted
        counts.update(roles)
    counts["table_paragraphs"] = table_paragraphs
    counts["ambiguous"] = meta["ambiguous"]
    counts["empty"] = meta["empty"]
    counts["protected_paragraphs"] = protected_paragraphs
    counts["skipped_paragraphs"] = skipped_paragraphs
    counts["skipped_tables"] = skipped_tables
    return dict(counts)


def _document_snapshot(doc) -> dict:
    body = doc.element.body
    section_values = []
    for section in doc.sections:
        section_values.append((
            section.page_width, section.page_height, section.left_margin, section.right_margin,
            section.top_margin, section.bottom_margin, section.orientation,
        ))
    merges = []
    for table in body.xpath(".//w:tbl"):
        for cell in table.xpath(".//w:tc"):
            tcpr = cell.find(qn("w:tcPr"))
            grid = tcpr.find(qn("w:gridSpan")) if tcpr is not None else None
            vertical = tcpr.find(qn("w:vMerge")) if tcpr is not None else None
            merges.append((grid.get(qn("w:val")) if grid is not None else None, vertical.get(qn("w:val")) if vertical is not None else None))
    links = sorted(rel.target_ref for rel in doc.part.rels.values() if rel.is_external)
    return {
        "text": tuple(body.xpath(".//w:t/text()")),
        "bookmarks": tuple(body.xpath(".//w:bookmarkStart/@w:name")),
        "tables": len(doc.tables),
        "merges": tuple(merges),
        "sections": tuple(section_values),
        "links": tuple(links),
    }


def _binary_parts(path: Path) -> dict[str, str]:
    prefixes = ("word/media/", "word/embeddings/", "word/fonts/")
    result = {}
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.startswith(prefixes) or name == "word/vbaProject.bin":
                result[name] = hashlib.sha256(archive.read(name)).hexdigest()
    return result


def _validate_output(source: Path, output: Path, before: dict) -> None:
    check = _open_document(output)
    after = _document_snapshot(check)
    for key in before:
        if before[key] != after[key]:
            raise SyncError(f"格式副本校验失败：{key} 与原文不一致，已删除未通过的副本")
    if _binary_parts(source) != _binary_parts(output):
        raise SyncError("格式副本校验失败：图片、嵌入对象或 VBA 部分与原文不一致，已删除未通过的副本")


def _format_source(source: Path, template: str | Path | None, profile_name: str | None) -> tuple[FormatProfile, str, bool]:
    if profile_name is not None:
        if template is not None:
            raise SyncError("样板报告与已存格式方案只能二选一")
        name = profile_name.strip()
        return load_format_profile(name), f"已存方案：{name}", False
    if template is None:
        return build_profile(source), str(source), True
    path = resolve_docx(str(template))
    return build_profile(path), str(path), False


def _scan_opened_report(source: Path, template_label: str, profile: FormatProfile, target, self_sample: bool = False, scope: dict[str, bool] | None = None) -> dict:
    scope = _validated_scope(scope)
    selected_roles = _selected_paragraph_roles(scope)
    classified, meta = _classify_paragraphs(target)
    counts = Counter(role for _, role in classified)
    protected_roles = Counter(role for paragraph, role in classified if (_outline_level(paragraph) or 0) > 4)
    protected = sum(protected_roles.values())
    selected_paragraphs = sum(counts[role] - protected_roles[role] for role in selected_roles)
    roles = []
    for role in ROLE_LABELS:
        if role in ("table_header", "table_total", "table_data"):
            continue
        if counts[role]:
            roles.append({
                "role": role,
                "label": ROLE_LABELS[role],
                "paragraphs": counts[role],
                "selected_paragraphs": counts[role] - protected_roles[role] if role in selected_roles else 0,
                "selected": role in selected_roles and counts[role] > protected_roles[role],
                "format": _format_summary(profile.paragraphs[role]),
            })
    table_enabled = scope["table_text"] or scope["table_emphasis"]
    table_rows = sum(len(table.rows) for table in target.tables) if table_enabled else 0
    warnings = []
    if meta["ambiguous"]:
        warnings.append(f"有 {meta['ambiguous']} 个大纲级别异常的段落保持原样，不会强行套用格式。")
    if source.suffix.lower() == ".docm":
        warnings.append("将校验 Word 宏（VBA）是否完整保留。")
    if self_sample:
        warnings.append("未指定样板，已按原文件自身格式提取。")
    return {
        "source": str(source),
        "template": template_label,
        "suggested_output": str(_suggested_output(source)),
        "scope": dict(scope),
        "totals": {
            "paragraphs": len(classified),
            "tables": len(target.tables) if table_enabled else 0,
            "table_rows": table_rows,
            "empty_paragraphs": meta["empty"],
            "ambiguous_paragraphs": meta["ambiguous"],
            "selected_paragraphs": selected_paragraphs,
            "protected_paragraphs": protected,
            "skipped_paragraphs": len(classified) - selected_paragraphs - protected,
            "skipped_tables": 0 if table_enabled else len(target.tables),
        },
        "roles": roles,
        "table_format": _format_summary(ParagraphFormat(profile.table_ascii_font, profile.table_east_asia_font, profile.table_size, None)),
        "small_table_policy": "不大于样板表格字号的字号保留；较大字号统一为样板字号。",
        "warnings": warnings,
    }


def scan_report_format(
    docx: str | Path,
    template: str | Path | None = None,
    *,
    profile_name: str | None = None,
    scope: dict | None = None,
) -> dict:
    scope = _validated_scope(scope)
    source = resolve_docx(str(docx))
    profile, template_label, self_sample = _format_source(source, template, profile_name)
    target = _open_document(source)
    return _scan_opened_report(source, template_label, profile, target, self_sample, scope)


def run_report_format(
    docx: str | Path,
    template: str | Path | None = None,
    *,
    output: str | Path | None = None,
    dry_run: bool = False,
    profile_name: str | None = None,
    scope: dict | None = None,
) -> dict:
    scope = _validated_scope(scope)
    source = resolve_docx(str(docx))
    output_path = _suggested_output(source) if output is None else Path(output).expanduser().resolve()
    if output_path.suffix.lower() != source.suffix.lower():
        raise SyncError("格式副本必须与原报告使用同一种文件类型（.docx 或 .docm）")
    if output_path == source:
        raise SyncError("为保护原报告，格式调整只生成新副本；请选择其他输出路径")
    if not output_path.parent.is_dir():
        raise SyncError(f"输出目录不存在：{output_path.parent}")
    if output_path.exists():
        raise SyncError(f"输出文件已存在，未覆盖：{output_path}")
    profile, template_label, self_sample = _format_source(source, template, profile_name)
    target = _open_document(source)
    result = _scan_opened_report(source, template_label, profile, target, self_sample, scope)
    result["output"] = str(output_path)
    result["dry_run"] = bool(dry_run)
    log(f"[格式] 格式来源：{template_label}")
    log(f"[格式] 待调整：{source.name}（处理 {result['totals']['selected_paragraphs']} 段、{result['totals']['tables']} 张表）")
    if dry_run:
        log(f"[格式] 预演完成，未写入文件；计划输出：{output_path.name}")
        return result

    before = _document_snapshot(target)
    counts = _apply_profile(target, profile, scope)
    fd, temp_name = tempfile.mkstemp(prefix=f".{output_path.stem}.", suffix=output_path.suffix, dir=output_path.parent)
    os.close(fd)
    staged = Path(temp_name)
    try:
        target.save(str(staged))
        if source.suffix.lower() == ".docm":
            _rewrite_main_type(staged, macro_enabled=True)
        _validate_output(source, staged, before)
        os.replace(staged, output_path)
    except Exception as exc:
        staged.unlink(missing_ok=True)
        if isinstance(exc, SyncError):
            raise
        raise SyncError(f"生成格式副本失败，原报告未改动：{exc}") from exc
    result["counts"] = counts
    result["applied"] = True
    log(f"[格式] 已生成：{output_path}")
    log("[格式] 已核验文字、书签、表格、分节、图片/嵌入对象和 VBA 与原文一致。")
    return result
