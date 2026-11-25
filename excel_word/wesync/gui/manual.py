"""把随包说明文档转成可导航、可搜索的 HTML。"""

from __future__ import annotations

import html
import re

from .resources import manual_path

HEADING_RE = re.compile(r"^(?:[一二三四五六七八九十]+、)\s*(.+)$")


def load_manual_text() -> str:
    return manual_path().read_text(encoding="utf-8-sig")


def parse_manual(text: str) -> list[tuple[str, list[str]]]:
    sections: list[tuple[str, list[str]]] = []
    title = "软件说明"
    lines: list[str] = []
    for raw in text.splitlines():
        line = raw.rstrip()
        stripped = line.strip()
        if HEADING_RE.match(stripped):
            if any(item.strip() for item in lines):
                sections.append((title, lines))
            title = stripped
            lines = []
        elif stripped and set(stripped) <= {"=", "-"}:
            continue
        else:
            lines.append(line)
    if any(item.strip() for item in lines):
        sections.append((title, lines))
    return sections or [("软件说明", text.splitlines())]


def _lines_html(lines: list[str]) -> str:
    result: list[str] = []
    for raw in lines:
        line = raw.rstrip()
        stripped = line.strip()
        if not stripped:
            result.append("<div class='gap'>&nbsp;</div>")
        elif stripped.startswith("·"):
            result.append(f"<div class='bullet'>{html.escape(stripped)}</div>")
        elif re.match(r"^\d+\)", stripped):
            result.append(f"<div class='step'>{html.escape(stripped)}</div>")
        else:
            result.append(f"<div class='line'>{html.escape(stripped)}</div>")
    return "".join(result)


def manual_html(sections: list[tuple[str, list[str]]]) -> str:
    parts = [
        "<!doctype html><html><head><meta charset='utf-8'><style>",
        "body{font-family:'Microsoft YaHei UI','Microsoft YaHei',sans-serif;color:#1F2937;font-size:14px;line-height:1.75;margin:18px 24px;}",
        "h1{font-size:24px;color:#1F4E79;margin:0 0 6px;}h2{font-size:18px;color:#1F4E79;margin:26px 0 10px;border-bottom:1px solid #E3E6EA;padding-bottom:6px;}",
        ".intro,.line,.step,.bullet{margin:4px 0;}.intro{color:#6B7280;}.gap{height:6px;}",
        ".bullet{padding-left:16px;text-indent:-12px;}.step{padding-left:20px;text-indent:-16px;}",
        "section{scroll-margin-top:12px;}</style></head><body>",
        "<h1>审计工具箱使用说明</h1><p class='intro'>本说明书与软件一同分发，无需联网。左侧选择章节，右上角可搜索关键词。</p>",
    ]
    for index, (title, lines) in enumerate(sections):
        anchor = f"chapter-{index}"
        parts.append(f"<section id='{anchor}'><h2>{html.escape(title)}</h2>{_lines_html(lines)}</section>")
    parts.append("</body></html>")
    return "".join(parts)
