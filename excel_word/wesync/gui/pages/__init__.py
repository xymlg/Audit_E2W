"""页面注册表：新增页面只需继承 BasePage 并调用 register_page()。"""

from __future__ import annotations

from ._page import BasePage
from .batch_page import BatchPage
from .blank_page import BlankPage
from .bookmark_page import BookmarkPage
from .export_page import ExportPage
from .format_page import FormatPage
from .help_page import HelpPage
from .history_page import HistoryPage
from .marks_page import MarksPage
from .writeback_page import WritebackPage

# 全局工具箱界面：左侧导航直接列出功能。
# 后续新增功能（如多文件合并）在此追加，或调用 register_page()。
_PAGES: list[type[BasePage]] = [ExportPage, WritebackPage, MarksPage, BookmarkPage, BlankPage, BatchPage, FormatPage, HistoryPage, HelpPage]


def register_page(page_class: type[BasePage]) -> type[BasePage]:
    """注册新功能页（自动加入导航）。"""
    _PAGES.append(page_class)
    return page_class


def iter_pages() -> list[type[BasePage]]:
    return [*_PAGES]


__all__ = ["BasePage", "register_page", "iter_pages", "ExportPage", "FormatPage", "WritebackPage", "MarksPage", "BookmarkPage", "BlankPage", "BatchPage", "HistoryPage", "HelpPage"]