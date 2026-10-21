#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOC/DOCX -> cleaned HTML/DOCX for WeChat editor.
Main path: Word filtered HTML -> clean -> output HTML/DOCX.

Dependencies:
- pywin32
- beautifulsoup4

Environment:
- Windows 11
- Microsoft Word installed
"""

from __future__ import annotations

import base64
import copy
import html
from collections import Counter
import json
import mimetypes
import re
import shutil
import tempfile
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
from urllib.parse import unquote, urlparse

try:
    from bs4 import BeautifulSoup, NavigableString, Tag
except Exception:
    BeautifulSoup = None  # type: ignore
    NavigableString = None  # type: ignore
    Tag = None  # type: ignore


SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_DIR = SCRIPT_DIR / "configs"

WD_FORMAT_FILTERED_HTML = 10
WD_FORMAT_XML_DOCUMENT = 12
WD_DO_NOT_SAVE_CHANGES = 0
WD_GO_TO_PAGE = 1
WD_GO_TO_ABSOLUTE = 1
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_UNDEFINED = 9999999
# Word 汇总多字符范围的 Bold/Underline 时不计这些字符的格式（单个字符的范围照常报告）。实测（整个 BMP，
# tests/perf-v117/fixtures/probe_ignored.py）为空格、U+00A0、U+2005、U+3000 与控制字符 U+0009、U+000B、U+000C、
# U+000E、U+0012；其余控制字符多数无法作为文字插入实测，一并按不计处理。U+000B 渲染时去掉，其格式不影响结果，不必读取
FORMAT_BLIND_CHARS = frozenset(" \xa0\u2005\u3000").union(
    chr(code) for code in range(0x20) if chr(code) not in "\x0b\r\x07")

# 公众号《内容排版规范》第 3 条：正文不设自定义字体族。官方检测引擎只放行以平台默认
# 字体栈开头的 font-family；结果页复制正文时会把每个元素的计算字体写进内联样式，
# 所以成品声明的字体必须就是这套默认字体栈（见 docs/wechat-mp-official/checklist.md）。
# 不加引号：样式写在双引号属性里，mp-quote 是合法标识符。
WECHAT_FONT_FAMILY = ("mp-quote, PingFang SC, system-ui, -apple-system, BlinkMacSystemFont, "
                      "Helvetica Neue, Hiragino Sans GB, Microsoft YaHei UI, Microsoft YaHei, Arial, sans-serif")
WECHAT_FONT_FAMILY_PREFIX = "mp-quote, PingFang SC, system-ui, -apple-system"
FONT_FAMILY_CONFIG_KEYS = (("output", "font_family"), ("output", "heading_font_family"),
                           ("styles", "subheading_font_family"))


def normalize_font_family(value: str) -> str:
    """与官方检测同口径：去引号、合并空白后比较前缀。"""
    return re.sub(r"\s+", " ", re.sub(r"[\"']", "", value or "")).strip()


def enforce_wechat_font_family(cfg: dict) -> dict:
    """配置里不以平台默认字体栈开头的字体族一律换成默认字体栈。"""
    for section, key in FONT_FAMILY_CONFIG_KEYS:
        block = cfg.get(section)
        if isinstance(block, dict) and key in block and not normalize_font_family(
                str(block[key])).startswith(WECHAT_FONT_FAMILY_PREFIX):
            block[key] = WECHAT_FONT_FAMILY
    return cfg


DEFAULT_CONFIG = {
    "output": {
        "font_family": WECHAT_FONT_FAMILY,
        "font_size_px": 14,
        "line_height": 1.75,
        "page_margin_css": "24px auto",
        "page_padding_css": "0 20px",
        "page_max_width_px": 860,
        "page_background_color": "#ffffff",
        "page_text_color": "#000000",
        "first_line_indent_em": 2,
        "paragraph_margin_bottom_em": 0.85,
        "paragraph_margin_top_css": "0",
        "paragraph_margin_right_css": "0",
        "paragraph_margin_left_css": "0",
        "paragraph_text_align": "left",
        "inline_bold_font_weight": 700,
        "inline_underline_text_decoration": "underline",
        "inline_underline_offset_em": 0.08,
        "generate_docx": False,
        "generate_html": True,
        "generate_images": True,
        "image_pixels_per_inch": 220,
        "figure_caption_font_size_px": 14,
        "figure_source_font_size_px": 13,
        "figure_caption_text_color": "#17365d",
        "figure_caption_margin_top_em": 0.55,
        "figure_caption_margin_bottom_em": 0.22,
        "figure_caption_margin_right_css": "0",
        "figure_caption_margin_left_css": "0",
        "figure_caption_text_indent_css": "0",
        "figure_caption_text_align": "left",
        "figure_caption_line_height": 1.6,
        "figure_source_margin_top_em": 0.18,
        "figure_source_margin_bottom_em": 1.2,
        "figure_source_margin_right_css": "0",
        "figure_source_margin_left_css": "0",
        "figure_source_text_indent_css": "0",
        "figure_source_text_align": "left",
        "figure_source_line_height": 1.6,
        "figure_source_text_color": "#888888",
        "image_margin_css": "0 auto",
        "image_border_css": "0",
        "figure_image_container_margin_css": "0",
        "figure_image_container_text_indent_css": "0",
        "figure_image_container_text_align": "center",
        "figure_image_container_line_height": 1,
        "summary_section_margin_css": "0 0 0.6em 0",
        "toc_image_margin_css": "0.6em 0 1.0em 0",
        "toc_image_text_indent_css": "0",
        "toc_image_text_align": "center",
        "toc_image_line_height": 1,
        "heading_font_family": WECHAT_FONT_FAMILY,
        "heading_level_1_font_size_px": 17,
        "heading_level_2_font_size_px": 14,
        "heading_level_3_font_size_px": 14,
        "heading_level_1_text_color": "#D97A45",
        "heading_level_1_bracket_color": "#5F8FE8",
        "heading_level_23_background_color": "#76A2E9",
        "heading_level_23_text_color": "#FFFFFF",
        "heading_level_23_line_height": 1.5,
        "heading_margin_top_em": 1.15,
        "heading_margin_bottom_em": 0.55,
        "heading_margin_horizontal_css": "0",
        "heading_level_1_margin_bottom_em": 0.0,
        "heading_inline_padding_vertical_em": 0.22,
        "heading_inline_padding_horizontal_em": 0.55,
        "heading_border_radius_px": 8,
        "heading_level_1_border_px": 6,
        "heading_text_indent_css": "0",
        "heading_level_1_text_align": "center",
        "heading_level_23_text_align": "left",
        "heading_level_1_font_weight": 700,
        "heading_level_23_font_weight": 700,
        "heading_level_1_border_style": "solid",
        "heading_level_1_background_color": "#ffffff",
        "heading_level_1_gap_horizontal_px": 8,
        "heading_level_1_gap_border_extra_px": 2,
        "heading_level_1_line_height": 1.6,
        "heading_spacer_margin_css": "0",
        "heading_spacer_text_indent_css": "0",
        "heading_spacer_text_align": "left",
        "list_item_margin_vertical_em": 0.2,
        "list_item_margin_horizontal_css": "0",
        "list_text_align": "left",
        "list_margin_top_css": "0",
        "list_margin_right_css": "0",
        "list_margin_left_em": 1.4,
        "list_padding_left_em": 1.2,
        "table_fallback_text_indent_css": "0",
        "weekly_heading_level_1_outer_margin_horizontal_css": "auto",
        "weekly_heading_level_1_padding_css": "0",
        "weekly_heading_level_1_text_align": "center",
        "weekly_heading_level_1_background_margin_css": "0 auto",
        "weekly_heading_level_1_background_size": "100% 100%",
        "weekly_heading_level_1_background_repeat": "no-repeat",
        "weekly_heading_level_1_background_position": "center",
        "weekly_heading_level_1_text_margin_right_css": "0",
        "weekly_heading_level_1_text_margin_bottom_css": "0",
    },
    "document": {
        "start_mode": "first_main_heading",
        "use_custom_heading_blocks": True,
        "use_heading_style_names_only": True,
        "stop_at_patterns": [
            "\u514d\u8d23\u58f0\u660e",
            "\u514d\u8d23\u53ca\u8bc4\u7ea7\u8bf4\u660e",
            "\u8bc4\u7ea7\u8bf4\u660e",
            "\u6295\u8d44\u8bc4\u7ea7\u6807\u51c6",
        ],
        "drop_title_patterns": [
            "\u5185\u5bb9\u76ee\u5f55",
            "\u56fe\u8868\u76ee\u5f55",
        ],
        "drop_exact_or_contains_patterns": [
            "\u8bf7\u52a1\u5fc5\u9605\u8bfb\u6b63\u6587\u4e4b\u540e\u7684\u514d\u8d23\u58f0\u660e\u90e8\u5206",
        ],
        "remove_toc": True,
        "remove_images": False,
        "remove_tables": True,
        "table_to_text": False,
        "render_tables": True,
        "render_valuation_table": False,
        "valuation_table_keyword": "估值表",
        "heading_as_bold_paragraph": True,
        # 正文导出前把本地图表改写为外链到不存在的工作簿（chart_relink.py）：Word 导出筛选 HTML 时
        # 不再逐张等待 Excel/excelcnv。关闭即完全旧行为（tests/perf-v117/design.md 第二节）
        "relink_local_charts": True,
    },
    "matching": {
        "main_heading_regex": r"^\s*\d+\.\s*\S+",
        "heading_like_regex": r"^\s*(?:\d+(?:\.\d+)*[\.,]?|[ivxlcdmIVXLCDM]+\.)\s*\S+",
        "heading_level_1_regex": r"^\s*\d+\.\s+\S+",
        "heading_level_2_regex": r"^\s*\d+\.\d+\.\s+\S+",
        "heading_level_3_regex": r"^\s*\d+\.\d+\.\d+\.\s+\S+",
        "heading_level_1_style_names": ["\u3010\u6b63\u6587\u3011\u4e00\u7ea7\u6807\u9898"],
        "heading_level_2_style_names": ["\u3010\u6b63\u6587\u3011\u4e8c\u7ea7\u6807\u9898"],
        "heading_level_3_style_names": ["\u3010\u6b63\u6587\u3011\u4e09\u7ea7\u6807\u9898"],
        "caption_regex": r"^\s*(?:\u56fe|\u8868)\s*\d+\s*[\uff1a:]",
        "source_regex": r"^\s*(?:\u6570\u636e\u6765\u6e90|\u8d44\u6599\u6765\u6e90|\u8d44\u6599\u53ca\u6570\u636e\u6765\u6e90|[Ss]ources?)\s*[\uff1a:]",
        "toc_entry_regex": r"\.{3,}\s*\d+\s*$",
    },
    # \u524d\u7f6e\u5185\u5bb9: \u5728\u6b63\u6587\u524d\u81ea\u52a8\u62fc\u300c\u6458\u8981\u300d, \u5e76\u5728\u6458\u8981\u4e0e\u6b63\u6587\u4e4b\u95f4\u63d2\u5165\u300c\u5185\u5bb9\u76ee\u5f55\u300d\u56fe\u7247\u3002
    # enable=False \u65f6\u672c\u5de5\u5177\u884c\u4e3a\u4e0e\u65e7\u7248\u5b8c\u5168\u4e00\u81f4(\u53ea\u8f93\u51fa\u6b63\u6587)\u3002config.json \u91cc enable=True \u5373\u5f00\u542f\u3002
    "front_matter": {
        "enable": False,
        "include_summary": True,        # \u6458\u8981(\u590d\u7528\u300c\u62a5\u544a\u6458\u8981\u8f6c\u516c\u4f17\u53f7\u300d\u903b\u8f91)
        "include_toc_image": True,      # \u5185\u5bb9\u76ee\u5f55\u56fe(\u7ae0\u8282\u76ee\u5f55, \u4e0d\u542b\u56fe\u8868\u76ee\u5f55)
        "toc_image": {
            "title": "\u76ee\u5f55",     # \u76ee\u5f55\u56fe\u9876\u90e8\u6807\u9898\u680f\u6587\u5b57
            "show_title_bar": True,
            "show_page_numbers": True,   # \u4ece Word \u76ee\u5f55\u57df\u8bfb\u9875\u7801; \u65e0\u76ee\u5f55\u57df\u5219\u9000\u56de\u5927\u7eb2(\u65e0\u9875\u7801)
            "max_levels": 3,            # \u76ee\u5f55\u6700\u591a\u6e32\u67d3\u5230\u7b2c\u51e0\u7ea7
        },
    },
}


class ProcessingError(Exception):
    pass


SUMMARY_DEFAULT_CONFIG = {
    "output": {
        "font_family": WECHAT_FONT_FAMILY,
        "font_size_px": 14,
        "line_height": 2.0,
        "page_margin_css": "24px auto",
        "page_padding_css": "0 20px",
        "page_max_width_px": 860,
        "page_background_color": "#ffffff",
        "page_text_color": "#000000",
        "first_line_indent_em": 2,
        "paragraph_margin_bottom_em": 0.85,
        "paragraph_margin_top_css": "0",
        "paragraph_margin_right_css": "0",
        "paragraph_margin_left_css": "0",
        "inline_bold_font_weight": 700,
        "inline_underline_text_decoration": "underline",
        "inline_underline_offset_em": 0.08,
        "generate_html": True,
        "output_suffix": "_summary_extract.html"
    },
    "summary": {
        "control_name_keywords": [
            "\u62a5\u544a\u6458\u8981",
            "AUTO_SAVE_Summary"
        ],
        "start_heading_keywords": [
            "\u6295\u8d44\u8981\u70b9",
            "\u6838\u5fc3\u89c2\u70b9",
            "\u6458\u8981",
            "\u4e3b\u8981\u89c2\u70b9"
        ],
        "require_left_align_for_heading": True,
        "fallback_to_first_page": True,
        "strip_leading_markers": True,
        "detect_heading_by_neighbors": True,
        "fast_format_runs": True
    },
    "styles": {
        "subheading_bg_color": "#76A2E9",
        "subheading_text_color": "#FFFFFF",
        "subheading_font_family": WECHAT_FONT_FAMILY,
        "subheading_font_size_px": 14,
        "subheading_line_height": 1.5,
        "subheading_border_radius_px": 8,
        "subheading_padding_vertical_em": 0.22,
        "subheading_padding_horizontal_em": 0.55,
        "subheading_margin_top_em": 1.15,
        "subheading_margin_right_css": "0",
        "subheading_margin_bottom_em": 0.55,
        "subheading_margin_left_css": "0",
        "subheading_text_indent_css": "0",
        "subheading_text_align": "left",
        "subheading_font_weight": 700,
        "bullet_line_height": 2.0,
        "bullet_paragraph_margin_css": "0",
        "bullet_paragraph_text_indent_css": "0",
        "bullet_paragraph_text_align": "left",
        "bullet_list_style_type": "square",
        "bullet_list_margin_top_css": "0",
        "bullet_list_margin_right_css": "0",
        "bullet_list_margin_left_css": "0",
        "bullet_list_padding_left_em": 1.55,
        "bullet_list_margin_bottom_em": 0.85,
        "bullet_item_margin_top_px": 8,
        "heading_blank_line_before": True,
        "heading_spacer_margin_css": "0",
        "heading_spacer_text_indent_css": "0",
        "heading_spacer_text_align": "left"
    },
    "matching": {
        "leading_marker_regex": r"^\[[A-Za-z0-9_]+\]",
        "bullet_chars_regex": r"^[\u25A0\u25AA\u25CF\u2022\u2027\u00B7\-\*]+",
        "space_regex": r"\s+"
    }
}


def deep_merge(base: dict, override: dict) -> dict:
    result = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def load_config(mode: str, config_dir: Optional[Path] = None) -> dict:
    if mode not in {"summary", "body", "weekly"}:
        raise ValueError(f"Unknown conversion mode: {mode}")
    config_root = config_dir or CONFIG_DIR
    base_path = config_root / "base.json"
    path = config_root / f"{mode}.json"
    defaults = SUMMARY_DEFAULT_CONFIG if mode == "summary" else DEFAULT_CONFIG
    # 深拷贝：返回配置的嵌套 dict 不得与模块级默认共享对象（调用方会就地 update）
    cfg = copy.deepcopy(defaults)
    if base_path.exists():
        with base_path.open("r", encoding="utf-8") as stream:
            cfg = deep_merge(cfg, json.load(stream))
    with path.open("r", encoding="utf-8") as stream:
        cfg = deep_merge(cfg, json.load(stream))
    cfg["_mode"] = mode
    return enforce_wechat_font_family(cfg)


def resolve_image_ppi(cfg: dict) -> int:
    """校验并返回 HTML 导出时的图片 DPI（WebOptions.PixelsPerInch）。

    返回 0 表示「保持 Word 默认（96）」。其余取值必须是整数且落在 Word 文档化的
    PixelsPerInch 区间 19..480。非法取值直接抛 ProcessingError，让配置错误尽早暴露，
    而不是悄悄按 96 DPI 导出模糊图表。
    """
    out = cfg.get("output")
    if not isinstance(out, dict):
        raise ProcessingError("config 'output' section is missing or not a JSON object.")
    raw = out.get("image_pixels_per_inch", 220)
    if isinstance(raw, bool) or not isinstance(raw, (int, float)):
        raise ProcessingError(
            f"output.image_pixels_per_inch must be an integer, got {raw!r}."
        )
    if isinstance(raw, float) and not raw.is_integer():
        raise ProcessingError(
            f"output.image_pixels_per_inch must be a whole number, got {raw!r}."
        )
    val = int(raw)
    if val == 0:
        return 0
    if val < 19 or val > 480:
        raise ProcessingError(
            f"output.image_pixels_per_inch must be 0 (disable) or 19..480, got {val}."
        )
    return val


def ensure_dependencies() -> None:
    missing = []
    if BeautifulSoup is None:
        missing.append("beautifulsoup4")
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
    except Exception:
        missing.append("pywin32")
    if missing:
        raise ProcessingError(
            "缺少依赖包：" + ", ".join(missing) + "。请在工作台侧栏点击「安装缺失依赖」，"
            "或手动执行：py -3 -m pip install pywin32 beautifulsoup4"
        )


# 确定意味着缺图的 render_problems 记录前缀（出口兜底据此补告警）
LOSS_PROBLEM_PREFIXES = ("图片地址为空或无法解析", "图片未能内嵌", "表格单元格里的图片无法读取", "表图渲染失败", "图框内贴入的表格")
# 网页成品里没能内嵌的本机图片标签（data: 与 http(s) 以外的地址）
LOCAL_IMG_RE = re.compile(r"<img\b[^>]*?\bsrc=\"(?!data:|https?:)[^\"]*\"[^>]*>", re.I)


def close_document(doc) -> None:
    """不保存关闭文档。关闭失败只报警：SaveAs 写出的文件、已读出的内容都不依赖关闭；
    Word 若因此不可用，下一次调用会照常报错或由调用方重建实例。"""
    if doc is None:
        return
    try:
        doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
    except Exception as exc:
        print(f"  [warn] Word document close failed: {type(exc).__name__}: {exc}")


def _end_marked_word(caption: str, wait_seconds: float = 10.0) -> bool:
    """结束标题含 caption 的那个 WINWORD（本程序实例的唯一窗口标记）：按 OpusApp 窗口标题定位，
    确认进程映像是 WINWORD.EXE 才结束。Word 打开文档后标题是「文档名 - 标记」，所以做包含匹配。
    返回是否结束了进程。"""
    import ctypes
    import time
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    kernel32.OpenProcess.restype = ctypes.c_void_p
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.QueryFullProcessImageNameW.argtypes = [ctypes.c_void_p, wintypes.DWORD, ctypes.c_wchar_p,
                                                    ctypes.POINTER(wintypes.DWORD)]
    kernel32.TerminateProcess.argtypes = [ctypes.c_void_p, ctypes.c_uint]
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]

    def find_pids():
        found = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def visit(hwnd, _):
            name = ctypes.create_unicode_buffer(64)
            user32.GetClassNameW(hwnd, name, 64)
            if name.value == "OpusApp":
                title = ctypes.create_unicode_buffer(512)
                user32.GetWindowTextW(hwnd, title, 512)
                if caption in title.value:
                    pid = wintypes.DWORD()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                    if pid.value:
                        found.append(pid.value)
            return True

        user32.EnumWindows(visit, 0)
        return sorted(set(found))

    ended = False
    deadline = time.monotonic() + wait_seconds
    while True:
        for pid in find_pids():
            handle = kernel32.OpenProcess(0x0001 | 0x1000, False, pid)  # TERMINATE | QUERY_LIMITED_INFORMATION
            if not handle:
                continue
            try:
                size = wintypes.DWORD(1024)
                image = ctypes.create_unicode_buffer(1024)
                if kernel32.QueryFullProcessImageNameW(handle, 0, image, ctypes.byref(size)) and \
                        image.value.lower().endswith("\\winword.exe"):
                    ended = bool(kernel32.TerminateProcess(handle, 1)) or ended
            finally:
                kernel32.CloseHandle(handle)
        if not find_pids() or time.monotonic() >= deadline:
            return ended
        time.sleep(0.5)


class WordApp:
    def __init__(self, image_ppi: int = 220) -> None:
        self.word = None
        self.pythoncom = None
        self.image_ppi = image_ppi
        self._ppi_warning_shown = False
        self.report_title = ""
        self.main_title = ""
        self.subtitle = ""
        self._title_checked = False
        self._caption = ""

    def __enter__(self):
        ensure_dependencies()
        import pythoncom

        self.pythoncom = pythoncom
        pythoncom.CoInitialize()
        try:
            self._start()
            return self
        except Exception as exc:
            self.__exit__(type(exc), exc, exc.__traceback__)
            raise

    def _start(self) -> None:
        import win32com.client

        self.word = win32com.client.DispatchEx("Word.Application")
        self.word.AutomationSecurity = 3
        self.word.Visible = False
        self.word.DisplayAlerts = 0
        self._mark_word_window()

    def restart(self) -> None:
        """重建 Word 实例（正文导出失败后 Word 已退出、COM 连接失效或文档关不掉时用）。

        先 Quit；Quit 失败则按本实例的窗口标记标题定位并结束这一个 WINWORD（与工作台
        reap_word_instance 的识别方式相同，绝不按进程名）。再 DispatchEx 并重设安全级别、
        隐藏、不弹框与窗口标记。已读的标题等状态保留，调用方持有的本对象引用不变。"""
        old, self.word = self.word, None
        quit_failed = old is None
        if old is not None:
            try:
                old.Quit()
            except Exception:
                quit_failed = True
        del old
        if quit_failed and self._caption:
            _end_marked_word(self._caption)
        self._start()

    def _mark_word_window(self):
        """给本实例 Word 主窗口打上唯一标记标题：调用方取消或超时后按标题实时定位
        这一个 WINWORD。标记优先取工作台经环境变量下发的随机值（进程号不可靠：
        venv 的 python.exe 是启动器，父进程记录的 PID 不是 worker 自己的）。"""
        try:
            import os

            marker = os.environ.get("WECHAT_WORKBENCH_MARKER") or str(os.getpid())
            self._caption = f"wechat-workbench-{marker}"
            self.word.Caption = self._caption
        except Exception:
            pass

    def open_document(self, src: Path):
        document = self.word.Documents.Open(
            str(src), ReadOnly=True, AddToRecentFiles=False,
            ConfirmConversions=False, Visible=False,
            PasswordDocument="", WritePasswordDocument="",
        )
        # 批注不属于正文：导出的 HTML 会把批注引用写成「[李1]」这类锚点混进句子。
        # 只删只读打开的内存副本里的批注，关闭时不保存，源文件不变
        try:
            if int(document.Comments.Count):
                document.DeleteAllComments()
        except Exception as comment_error:
            print(f"  [warn] 批注未能移除，成品中可能残留批注引用标记：{comment_error}")
        if not self._title_checked:
            self._title_checked = True
            try:
                self.main_title, self.subtitle = self._read_title_fields(document)
                self.report_title = self.subtitle
            except Exception as title_error:
                print(f"  [warn] report title unavailable: {title_error}")
        return document

    @staticmethod
    def _plain_title(value) -> str:
        return re.sub(r"\s+", " ", re.sub(r"[\x00-\x1f]", " ", str(value or ""))).strip()

    def _read_title_fields(self, document) -> Tuple[str, str]:
        main_title = ""
        subtitle = ""
        fallback_title = ""
        fallback_main = ""
        legacy_title = ""
        controls = document.ContentControls
        for index in range(1, int(controls.Count) + 1):
            control = controls(index)
            if control.ShowingPlaceholderText:
                continue
            tag = self._plain_title(control.Tag)
            alias = self._plain_title(control.Title)
            title = self._plain_title(control.Range.Text)
            if tag == "Indicator_SecondTitle":
                main_title = main_title or title
            elif tag == "Indicator_ReportTitle":
                subtitle = subtitle or title
            elif alias == "主标题":
                fallback_main = fallback_main or title
            elif alias in {"报告标题", "报告名称", "标题"}:
                fallback_title = fallback_title or title
            elif tag == "BDCONTENTCONTROL_AUTO_SAVE_SubTitle":
                legacy_title = legacy_title or title
        main_title = main_title or fallback_main
        subtitle = subtitle or fallback_title or legacy_title
        if subtitle:
            return main_title, subtitle
        title_styles = {"Title", "标题", "【首页】标题", "【首页】报告标题"}
        first_page = document.GoTo(What=WD_GO_TO_PAGE, Which=WD_GO_TO_ABSOLUTE, Count=1)
        paragraphs = first_page.Bookmarks(r"\Page").Range.Paragraphs
        for index in range(1, int(paragraphs.Count) + 1):
            paragraph = paragraphs(index)
            style = paragraph.Style
            names = {self._plain_title(getattr(style, name, "")) for name in ("NameLocal", "Name")}
            if names & title_styles:
                title = self._plain_title(paragraph.Range.Text)
                if title and title != main_title:
                    return main_title, title
        return main_title, ""

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.word is not None:
                try:
                    self.word.Quit()
                except Exception as close_error:
                    # with 块已正常结束时 Word 的工作都已完成，实例却已退出或断开（常见于内存加速、
                    # 清理类软件结束了后台 Word）：成品不受影响，按窗口标记收尾后继续写出
                    if exc_type is None:
                        print(f"  [warn] Word 已提前退出（常见于内存加速、清理类软件结束了后台 Word），本次内容已全部取得：{close_error}")
                    else:
                        print(f"  [warn] Word close failed: {close_error}")
                    if self._caption:
                        try:
                            _end_marked_word(self._caption, wait_seconds=5.0)
                        except Exception as reap_error:
                            print(f"  [warn] Word cleanup failed: {reap_error}")
        finally:
            self.word = None
            if self.pythoncom is not None:
                try:
                    self.pythoncom.CoUninitialize()
                except Exception:
                    pass
                self.pythoncom = None

    def export_filtered_html(self, src: Path, html_path: Path) -> None:
        doc = None
        try:
            doc = self.open_document(src)
            # 提高 Web 保存时的图片导出分辨率。Word 默认按 96 DPI 计算导出图片的像素尺寸，
            # 会把矢量的 Office Chart 栅格化得很糊。PixelsPerInch 越大，导出 PNG 越清晰。
            # image_ppi 已在 resolve_image_ppi() 中校验；这里只兜底「老版本 Word 不支持该属性」。
            if self.image_ppi:
                try:
                    doc.WebOptions.PixelsPerInch = self.image_ppi
                except Exception as exc:
                    if not self._ppi_warning_shown:
                        self._ppi_warning_shown = True
                        print(
                            f"  [warn] Could not set image DPI to {self.image_ppi}; "
                            f"charts may export at Word's default 96 DPI. ({exc})"
                        )
            doc.SaveAs(str(html_path), FileFormat=WD_FORMAT_FILTERED_HTML)
        finally:
            close_document(doc)

    def export_docx(self, src: Path, docx_path: Path) -> None:
        doc = None
        try:
            doc = self.open_document(src)
            doc.SaveAs(str(docx_path), FileFormat=WD_FORMAT_XML_DOCUMENT)
        finally:
            close_document(doc)

    def convert_html_to_docx(self, html_path: Path, docx_path: Path) -> None:
        doc = None
        try:
            doc = self.open_document(html_path)
            doc.SaveAs(str(docx_path), FileFormat=WD_FORMAT_XML_DOCUMENT)
        finally:
            close_document(doc)

    def extract_front_matter(self, src: Path, summary_extractor, want_summary: bool, want_toc: bool):
        """打开源文档一次, 返回 (摘要正文 HTML, 目录域条目列表)。
        - 摘要: 复用 summary_extractor(未命中返回 "");
        - 目录: 返回**所有**目录域(TablesOfContents)解析结果, 由调用方挑出「内容目录」并剔除图表目录。
        任一步失败都只影响该步(前置内容优雅降级), 不影响正文导出。"""
        doc = None
        summary_html = ""
        toc_list: List[List[dict]] = []
        try:
            doc = self.open_document(src)
            if want_summary and summary_extractor is not None:
                try:
                    summary_html = summary_extractor.extract_body_html(doc) or ""
                except Exception as exc:
                    print(f"  [warn] summary extract failed: {exc}")
                    summary_html = ""
            if want_toc:
                try:
                    toc_list = self._read_toc_entries(doc)
                except Exception as exc:
                    print(f"  [warn] TOC field read failed: {exc}")
                    toc_list = []
        finally:
            close_document(doc)
        return summary_html, toc_list

    def _read_toc_entries(self, doc) -> List[List[dict]]:
        """读出文档里每个目录域(TablesOfContents)的条目, 每个域 -> [{level,title,page}, ...]。

        标题(含「1. / 1.1.」编号)直接取自目录域文本; 页码取自目录条目超链接指向的书签
        (_Toc...) 所在页(Range.Information)。无头模式下 Word 默认不分页, 故先 Repaginate,
        否则页码取不到。任一步失败仅令该条目页码为 None, 不影响标题/层级。"""
        result: List[List[dict]] = []
        try:
            coll = doc.TablesOfContents
            count = int(coll.Count)
        except Exception:
            return result
        if count <= 0:
            return result
        # 无头模式必须先分页, 否则 Information(页码) 取不到真实值。
        try:
            doc.Repaginate()
        except Exception:
            pass
        for i in range(1, count + 1):
            try:
                toc = coll.Item(i)
                rng = toc.Range
                pcount = int(rng.Paragraphs.Count)
            except Exception:
                continue
            entries: List[dict] = []
            for j in range(1, pcount + 1):
                try:
                    para = rng.Paragraphs(j)
                    prange = para.Range
                    raw = str(prange.Text)
                except Exception:
                    continue
                entry = self._parse_toc_paragraph(para, raw)
                if not entry:
                    continue
                # 书签页码最可靠; 取不到时才退回从文本里解析出的页码(若有)。
                page = self._toc_entry_page(doc, prange)
                if page is not None:
                    entry["page"] = page
                entries.append(entry)
            if entries:
                result.append(entries)
        return result

    @staticmethod
    def _toc_entry_page(doc, prange) -> Optional[int]:
        """目录条目通常是指向书签(_Toc...)的超链接; 返回该书签所在页(已分页)。取不到返回 None。"""
        WD_ACTIVE_END_ADJUSTED_PAGE = 1
        try:
            if int(prange.Hyperlinks.Count) < 1:
                return None
            sub = str(prange.Hyperlinks(1).SubAddress or "")
            if not sub:
                return None
            bm = doc.Bookmarks(sub)
            return int(bm.Range.Information(WD_ACTIVE_END_ADJUSTED_PAGE))
        except Exception:
            return None

    def _parse_toc_paragraph(self, para, raw: str) -> Optional[dict]:
        """把一行目录域文本解析成 {level, title, page}。无法识别返回 None。"""
        for ch in ("\r", "\x07", "\x13", "\x14", "\x15", "\x1e"):
            raw = raw.replace(ch, "")
        raw = raw.replace("\x0b", " ").replace("\xa0", " ")
        if not raw.strip():
            return None

        page = None
        title = raw
        # 1) 多数 Word 目录用右制表位放页码: 末段是纯数字即页码
        if "\t" in raw:
            parts = raw.split("\t")
            last = parts[-1].strip()
            if last.isdigit():
                page = int(last)
                title = "\t".join(parts[:-1])
        # 2) 点线引导符(......)/空白 + 末尾页码
        if page is None:
            m = re.search(r"[.…\s\t]{2,}(\d{1,4})\s*$", raw)
            if m:
                page = int(m.group(1))
                title = raw[:m.start()]

        title = title.replace("\t", " ")
        title = re.sub(r"[.…]{2,}\s*$", "", title)
        title = normalize_text(title)
        if not title:
            return None
        # 目录/图表目录的标题行本身(若混进域范围)直接跳过
        if title in ("目录", "内容目录", "图表目录", "Contents", "Table of Contents"):
            return None

        return {"level": self._toc_paragraph_level(para), "title": title, "page": page}

    @staticmethod
    def _toc_paragraph_level(para) -> int:
        """从目录段落样式名(TOC 1 / 目录 1 / ...)末尾的数字推断层级; 取不到则按 1。"""
        name = ""
        try:
            st = para.Style
            name = str(getattr(st, "NameLocal", "") or getattr(st, "Name", "") or "")
        except Exception:
            name = ""
        if not name:
            try:
                name = str(para.Style)
            except Exception:
                name = ""
        m = re.search(r"(\d+)\s*$", name)
        if m:
            try:
                return max(1, min(9, int(m.group(1))))
            except Exception:
                return 1
        return 1


def _looks_figure_caption(title: str) -> bool:
    """是否像「图表目录」里的条目(图1/表2/Figure 3/Exhibit 4...), 用于把图表目录从内容目录里剔掉。"""
    return bool(re.match(r"^\s*(?:图表|图|表|figure|fig|table|exhibit|chart)\s*\d", title or "", re.I))


def _toc_title_key(title: str) -> str:
    """目录条目/大纲标题去掉开头的章节编号(1. / 1.1. / 1.1 / 一、)后再做松散匹配键, 让两侧对齐。

    数字编号必须后接「点号/全角点/空白」才剥离, 避免把以年份/季度开头的真实标题误剥
    (如 `2025年市场回顾`→不剥、`3Q收入`→不剥; 而 `1. xxx`/`1.1 xxx` 正常剥)。"""
    t = re.sub(r"^\s*(?:\d+(?:[.．]\d+)*(?:[.．]|\s+)|[一二三四五六七八九十]+[、.]\s*)", "", title or "")
    return loose_text_key(t)


def select_content_toc(toc_list: List[List[dict]], outline: List[dict]) -> List[dict]:
    """多个目录域时, 用大纲(章节标题)重合度 + 图表条目惩罚, 挑出「内容目录」那一个。"""
    toc_list = [t for t in (toc_list or []) if t]
    if not toc_list:
        return []
    if len(toc_list) == 1:
        return toc_list[0]
    # 大纲侧也用 _toc_title_key 去编号, 与目录条目键对齐(正文标题带编号写法时也能命中)。
    okeys = {_toc_title_key(o.get("text", "")) for o in (outline or []) if o.get("text")}
    okeys.discard("")
    best, best_score = None, None
    for toc in toc_list:
        score = 0
        for e in toc:
            k = _toc_title_key(e.get("title", ""))
            if k and k in okeys:
                score += 2
            if _looks_figure_caption(e.get("title", "")):
                score -= 2
        if best_score is None or score > best_score:
            best_score, best = score, toc
    return best or toc_list[0]


def build_toc_entries(toc_list: List[List[dict]], outline: List[dict], fm_cfg: dict) -> List[dict]:
    """把目录域解析结果整理成可渲染的内容目录条目:
    - 多个目录域里挑出内容目录(剔除图表目录);
    - 用大纲的可靠层级校正各条目的 level, 并剔除残留的图表条目;
    - 无目录域时退回大纲(无页码);
    - 按 max_levels 截断。"""
    max_levels = int(((fm_cfg or {}).get("toc_image") or {}).get("max_levels", 3) or 3)
    entries = select_content_toc(toc_list, outline)

    if entries:
        okey_level: Dict[str, int] = {}
        for o in (outline or []):
            k = _toc_title_key(o.get("text", ""))   # 与目录条目键一致(同样去编号), 否则带编号正文标题对不上
            if k and k not in okey_level:
                okey_level[k] = int(o.get("level", 1) or 1)
        normed: List[dict] = []
        for e in entries:
            title = (e.get("title") or "").strip()
            if not title or _looks_figure_caption(title):
                continue
            level = int(e.get("level", 1) or 1)
            k = _toc_title_key(title)
            if k in okey_level:
                level = okey_level[k]
            normed.append({"level": level, "title": title, "page": e.get("page")})
        entries = normed

    if not entries and outline:
        entries = [
            {"level": int(o.get("level", 1) or 1), "title": (o.get("text") or "").strip(), "page": None}
            for o in outline if (o.get("text") or "").strip()
        ]

    return [e for e in entries if e["title"] and e["level"] <= max_levels
            and not _looks_figure_caption(e["title"])]


class SummaryExtractor:
    def __init__(self, cfg: dict):
        self.out_cfg = cfg["output"]
        self.sum_cfg = cfg["summary"]
        self.style_cfg = cfg["styles"]
        self.match_cfg = cfg["matching"]

        self.re_leading_marker = re.compile(self.match_cfg["leading_marker_regex"])
        self.re_bullet_chars = re.compile(self.match_cfg["bullet_chars_regex"])
        self.re_space = re.compile(self.match_cfg["space_regex"])
        self.last_inline_route: Tuple[str, str, int] = ("per_char", "", 0)

    def extract_html(self, doc) -> str:
        rng = self._find_summary_range(doc)
        if rng is None:
            raise ProcessingError("Could not locate summary content on page 1.")

        items = self._extract_paragraph_items(rng)
        if not items:
            raise ProcessingError("Summary range found, but no usable paragraphs were extracted.")

        body_html = "\n".join(items)
        return self._wrap_full_html(body_html)

    def extract_body_html(self, doc) -> str:
        document = BeautifulSoup(self.extract_html(doc), "html.parser")
        return document.body.decode_contents()

    def _find_summary_range(self, doc):
        rng = self._find_content_control_range(doc)
        if rng is not None:
            return rng

        if not self.sum_cfg.get("fallback_to_first_page", True):
            return None

        page_rng = self._get_page_range(doc, 1)
        if page_rng is None:
            return None

        paras = self._range_paragraphs(page_rng)
        if not paras:
            return None

        start_idx = self._find_start_paragraph_index(paras)
        if start_idx is None:
            return None

        end_idx = self._find_end_paragraph_index(paras, start_idx)
        start_pos = paras[start_idx].Range.Start
        end_pos = paras[end_idx].Range.End
        return doc.Range(Start=start_pos, End=end_pos)

    def _find_content_control_range(self, doc):
        keywords = self.sum_cfg.get("control_name_keywords", [])
        try:
            controls = doc.ContentControls
            count = int(controls.Count)
        except Exception:
            return None

        for idx in range(1, count + 1):
            try:
                cc = controls(idx)
            except Exception:
                continue
            try:
                title = self._normalize_text(str(getattr(cc, "Title", "") or ""))
                tag = self._normalize_text(str(getattr(cc, "Tag", "") or ""))
            except Exception:
                continue
            hay = title + " " + tag
            if any(k in hay for k in keywords):
                try:
                    is_placeholder = bool(getattr(cc, "ShowingPlaceholderText", False))
                except Exception:
                    is_placeholder = False
                if is_placeholder:
                    continue
                try:
                    return cc.Range.Duplicate
                except Exception:
                    continue
        return None

    def _get_page_range(self, doc, page_number: int):
        try:
            anchor = doc.GoTo(What=WD_GO_TO_PAGE, Which=WD_GO_TO_ABSOLUTE, Count=page_number)
            return anchor.Bookmarks(r"\Page").Range
        except Exception:
            return None

    def _range_paragraphs(self, rng) -> List:
        paragraphs = []
        try:
            count = int(rng.Paragraphs.Count)
        except Exception:
            return paragraphs
        for idx in range(1, count + 1):
            try:
                paragraphs.append(rng.Paragraphs(idx))
            except Exception:
                continue
        return paragraphs

    def _find_start_paragraph_index(self, paras: List) -> Optional[int]:
        keywords = self.sum_cfg.get("start_heading_keywords", [])
        for idx, para in enumerate(paras):
            text = self._paragraph_plain_text(para)
            if not text:
                continue
            if any((text == k) or text.startswith(k) for k in keywords):
                return idx
        # fallback: first short non-bullet heading-like line followed by bullets
        for idx, para in enumerate(paras):
            text = self._paragraph_plain_text(para)
            if not text:
                continue
            is_bullet = self._paragraph_is_bullet(para)
            next_bullet = self._next_meaningful_is_bullet(paras, idx)
            if self._paragraph_is_summary_heading(para, text, is_bullet, False, next_bullet):
                return idx
        return None

    def _find_end_paragraph_index(self, paras: List, start_idx: int) -> int:
        last_non_empty = start_idx
        for idx in range(start_idx, len(paras)):
            text = self._paragraph_plain_text(paras[idx])
            if text:
                last_non_empty = idx
        return last_non_empty

    def _extract_paragraph_items(self, rng) -> List[str]:
        items: List[str] = []
        paras = self._range_paragraphs(rng)
        # 每段的文字与是否列表项只向 Word 取一次（原函数、原调用），前后有效邻居用缓存值判断；提取期间文档不修改，结果不变
        texts = [self._paragraph_plain_text(para) for para in paras]
        bullets = [bool(text) and self._paragraph_is_bullet(para) for para, text in zip(paras, texts)]
        prev_bullets: List[bool] = []
        last = False
        for text, bullet in zip(texts, bullets):
            prev_bullets.append(last)
            if text:
                last = bullet
        next_bullets: List[bool] = []
        last = False
        for text, bullet in zip(reversed(texts), reversed(bullets)):
            next_bullets.append(last)
            if text:
                last = bullet
        next_bullets.reverse()
        meta = []
        for idx, para in enumerate(paras):
            raw_text = texts[idx]
            if not raw_text:
                meta.append({"para": para, "text": "", "is_bullet": False, "is_heading": False})
                continue
            is_bullet = bullets[idx]
            prev_bullet = prev_bullets[idx]
            next_bullet = next_bullets[idx]
            is_heading = self._paragraph_is_summary_heading(para, raw_text, is_bullet, prev_bullet, next_bullet)
            meta.append({
                "para": para,
                "text": raw_text,
                "is_bullet": is_bullet,
                "is_heading": is_heading,
            })

        # 步内进度（只供工作台 ETA，工作台不把它放进日志）：第二遍逐段取格式，快速路径与逐字回退共用一个计数。
        # 字数取第一遍已读到的各段文字，不增加 Word 调用；约每 5% 字数、至少隔 0.5 秒打印一次，最后一段必打。
        import time
        sub_total = sum(len(m["text"]) for m in meta)
        sub_state = {"done": 0, "shown": 0, "at": time.monotonic()}
        print(f"[SUB] summary 0/{sub_total}", flush=True)

        def sub_progress(entry):
            sub_state["done"] += len(entry["text"])
            now = time.monotonic()
            last = sub_state["done"] >= sub_total
            if last or (sub_state["done"] - sub_state["shown"] >= 0.05 * sub_total and now - sub_state["at"] >= 0.5):
                sub_state["shown"], sub_state["at"] = sub_state["done"], now
                print(f"[SUB] summary {sub_state['done']}/{sub_total}", flush=True)

        i = 0
        while i < len(meta):
            entry = meta[i]
            text = entry["text"]
            para = entry["para"]
            if not text:
                i += 1
                continue

            if entry["is_heading"]:
                html_inner = self._render_paragraph_inline(para.Range, is_bullet=False)
                html_inner = self._cleanup_inline_html(html_inner)
                if html_inner:
                    if self.style_cfg.get("heading_blank_line_before", True):
                        spacer_margin = self.style_cfg.get("heading_spacer_margin_css", "0")
                        spacer_indent = self.style_cfg.get("heading_spacer_text_indent_css", "0")
                        spacer_align = self.style_cfg.get("heading_spacer_text_align", "left")
                        items.append(
                            f'<p style="margin:{spacer_margin}; text-indent:{spacer_indent}; '
                            f'text-align:{spacer_align};"><br/></p>'
                        )
                    items.append(self._render_subheading(html_inner))
                sub_progress(entry)
                i += 1
                continue

            if entry["is_bullet"]:
                li_blocks: List[str] = []
                while i < len(meta) and meta[i]["text"] and meta[i]["is_bullet"]:
                    li_inner = self._render_paragraph_inline(meta[i]["para"].Range, is_bullet=True)
                    li_inner = self._cleanup_inline_html(li_inner)
                    if li_inner:
                        li_blocks.append(self._render_bullet_list_item(li_inner))
                    sub_progress(meta[i])
                    i += 1
                if li_blocks:
                    items.append(self._render_bullet_list(li_blocks))
                continue

            html_inner = self._render_paragraph_inline(para.Range, is_bullet=False)
            html_inner = self._cleanup_inline_html(html_inner)
            if html_inner:
                items.append(f"<p>{html_inner}</p>")
            sub_progress(entry)
            i += 1
        return items

    def _paragraph_plain_text(self, para) -> str:
        try:
            text = str(para.Range.Text)
        except Exception:
            return ""
        text = text.replace("\r", " ").replace("\x07", " ").replace("\x0b", " ")
        text = text.replace("\xa0", " ").replace("\t", " ")
        text = self._normalize_text(text)
        if self.sum_cfg.get("strip_leading_markers", True):
            text = self.re_leading_marker.sub("", text).strip()
        if not self._ascii_minus_is_content_prefix(text):
            text = self.re_bullet_chars.sub("", text).strip()
        return text

    @staticmethod
    def _ascii_minus_is_content_prefix(text: str) -> bool:
        match = re.match(r"^\s*-", text)
        return bool(match and (match.end() == len(text) or not text[match.end()].isspace()))

    def _paragraph_is_bullet(self, para) -> bool:
        try:
            list_type = int(para.Range.ListFormat.ListType)
            if list_type != 0:
                return True
        except Exception:
            pass
        try:
            list_string = str(para.Range.ListFormat.ListString or "")
            if list_string.strip():
                return True
        except Exception:
            pass
        try:
            raw = self._normalize_text(str(para.Range.Text).replace("\r", " ").replace("\x07", " "))
        except Exception:
            return False
        if self._ascii_minus_is_content_prefix(raw):
            return False
        return bool(self.re_bullet_chars.search(raw))

    def _next_meaningful_is_bullet(self, paras: List, idx: int) -> bool:
        j = idx + 1
        while j < len(paras):
            txt = self._paragraph_plain_text(paras[j])
            if txt:
                return self._paragraph_is_bullet(paras[j])
            j += 1
        return False

    def _paragraph_is_summary_heading(self, para, plain_text: str, is_bullet: bool, prev_bullet: bool, next_bullet: bool) -> bool:
        if is_bullet:
            return False
        if not plain_text:
            return False

        normalized = self._normalize_text(plain_text)
        keywords = self.sum_cfg.get("start_heading_keywords", [])
        if any((normalized == k) or normalized.startswith(k) for k in keywords):
            return True

        if self.sum_cfg.get("require_left_align_for_heading", True):
            try:
                if int(para.Alignment) != WD_ALIGN_PARAGRAPH_LEFT:
                    return False
            except Exception:
                pass
        try:
            left_indent = float(para.LeftIndent)
            first_line = float(para.FirstLineIndent)
            if abs(left_indent) > 1 or abs(first_line) > 1:
                return False
        except Exception:
            pass

        if self.sum_cfg.get("detect_heading_by_neighbors", True):
            if next_bullet:
                return True
            if prev_bullet and not next_bullet:
                return False
        return False

    def _render_paragraph_inline(self, rng, is_bullet: bool = False) -> str:
        # 先按格式二分成段读取（tests/perf-v117/design.md 第一节）；任一前提不成立或出错即丢弃中间结果，整段走逐字路径
        if self.sum_cfg.get("fast_format_runs", True):
            try:
                rendered = self._render_paragraph_inline_fast(rng, is_bullet)
            except Exception as exc:
                self.last_inline_route = ("per_char", f"error:{type(exc).__name__}", self.last_inline_route[2])
                rendered = None
            if rendered is not None:
                return rendered
        else:
            self.last_inline_route = ("per_char", "disabled", 0)
        return self._render_paragraph_inline_per_char(rng, is_bullet)

    def _inline_skip_count(self, raw_text: str, is_bullet: bool) -> int:
        # 与 _render_paragraph_inline_per_char 开头的跳过字数算法相同
        stripped_raw = raw_text.replace("\r", "").replace("\x07", "").replace("\x0b", "")
        skip_count = 0
        if self.sum_cfg.get("strip_leading_markers", True):
            marker_match = self.re_leading_marker.match(stripped_raw)
            if marker_match:
                skip_count = len(marker_match.group(0))
        if is_bullet:
            tmp = stripped_raw[skip_count:]
            if not self._ascii_minus_is_content_prefix(tmp):
                bullet_match = re.match(r"^[\s\u25A0\u25AA\u25CF\u2022\u2027\u00B7\-\*]+", tmp)
                if bullet_match:
                    skip_count += len(bullet_match.group(0))
        return skip_count

    def _render_paragraph_inline_fast(self, rng, is_bullet: bool) -> Optional[str]:
        """格式相同的连续文字一次读取：按 Bold/Underline 是否为 wdUndefined 二分段落范围，叶子内每个字共享叶子的键。

        返回 None 表示前提不成立，由调用方整段走逐字路径；last_inline_route 记录（路径, 前提或原因, 节点数）。
        """
        self.last_inline_route = ("per_char", "error", 0)
        raw_text = str(rng.Text)
        start = int(rng.Start)
        end = int(rng.End)
        # 段落标记 \r 与单元格结束标记 \r\x07 在位置上都只占 1 个位置
        if raw_text.endswith("\r\x07"):
            body, end = raw_text[:-2], end - 1
        elif raw_text.endswith("\r"):
            body, end = raw_text[:-1], end - 1
        else:
            body = raw_text
        # 位置与字符一一对应才能按位置二分；代理对、隐藏文字、域代码会让两者不等
        if len(body) != end - start:
            self.last_inline_route = ("per_char", "length", 0)
            return None
        # 文字读取含域代码时，逐字路径读到的与按位置分段读到的不同（合成夹具实测）
        if rng.TextRetrievalMode.IncludeFieldCodes:
            self.last_inline_route = ("per_char", "field_codes", 0)
            return None
        if int(rng.Fields.Count) != 0:
            self.last_inline_route = ("per_char", "fields", 0)
            return None
        if int(rng.InlineShapes.Count) != 0:
            self.last_inline_route = ("per_char", "inline_shapes", 0)
            return None

        # Duplicate 保持原范围所在的 story 与文字读取设置；位置由 Python 整数传递，不再读 Start/End
        probe = rng.Duplicate
        leaves: List[Tuple[str, Tuple[bool, bool]]] = []
        nodes = 0
        stack = [(start, end)] if end > start else []
        while stack:
            a, b = stack.pop()
            nodes += 1
            self.last_inline_route = ("per_char", "error", nodes)
            probe.SetRange(a, b)
            bold = int(probe.Bold)
            underline = int(probe.Underline) if bold != WD_UNDEFINED else WD_UNDEFINED
            if bold != WD_UNDEFINED and underline != WD_UNDEFINED:
                text = str(probe.Text)
                if len(text) != b - a:
                    self.last_inline_route = ("per_char", "text_mismatch", nodes)
                    return None
                key = (bold != 0, underline != 0)
                # 汇总值不反映叶子里空白字符自己的格式：这些字符逐个读取，按各自的键拆开叶子
                cut = 0
                for k, ch in enumerate(text if b - a > 1 else ""):
                    if ch not in FORMAT_BLIND_CHARS:
                        continue
                    nodes += 1
                    probe.SetRange(a + k, a + k + 1)
                    blind_bold = int(probe.Bold)
                    blind_underline = int(probe.Underline)
                    if WD_UNDEFINED in (blind_bold, blind_underline):
                        self.last_inline_route = ("per_char", "mixed_single", nodes)
                        return None
                    if k > cut:
                        leaves.append((text[cut:k], key))
                    leaves.append((ch, (blind_bold != 0, blind_underline != 0)))
                    cut = k + 1
                if cut < len(text):
                    leaves.append((text[cut:], key))
                continue
            if b - a <= 1:
                self.last_inline_route = ("per_char", "mixed_single", nodes)
                return None
            mid = (a + b) // 2
            stack.append((mid, b))
            stack.append((a, mid))
        if "".join(text for text, _ in leaves) != body:
            self.last_inline_route = ("per_char", "text_mismatch", nodes)
            return None

        # 与逐字路径逐字的处理顺序相同：去标记 → 为空跳过（不读键、不打断合并段）→ 换空白 → 消耗跳过字数 → 按键合并
        parts: List[str] = []
        current_key: Optional[Tuple[bool, bool]] = None
        current_buf: List[str] = []
        remaining_skip = self._inline_skip_count(raw_text, is_bullet)
        for txt, key in leaves:
            txt = txt.replace("\r", "").replace("\x07", "").replace("\x0b", "")
            txt = txt.replace("\t", " ").replace("\xa0", " ")
            if not txt:
                continue

            if remaining_skip > 0:
                if len(txt) <= remaining_skip:
                    remaining_skip -= len(txt)
                    continue
                txt = txt[remaining_skip:]
                remaining_skip = 0

            if current_key is None:
                current_key = key
            if key != current_key:
                parts.append(self._wrap_run("".join(current_buf), current_key[0], current_key[1]))
                current_buf = []
                current_key = key
            current_buf.append(html.escape(txt))

        if current_buf and current_key is not None:
            parts.append(self._wrap_run("".join(current_buf), current_key[0], current_key[1]))

        self.last_inline_route = ("fast", "", nodes)
        return self._cleanup_inline_html("".join(parts))

    def _render_paragraph_inline_per_char(self, rng, is_bullet: bool = False) -> str:
        raw_text = str(rng.Text)
        stripped_raw = raw_text.replace("\r", "").replace("\x07", "").replace("\x0b", "")
        skip_count = 0
        if self.sum_cfg.get("strip_leading_markers", True):
            marker_match = self.re_leading_marker.match(stripped_raw)
            if marker_match:
                skip_count = len(marker_match.group(0))
        if is_bullet:
            tmp = stripped_raw[skip_count:]
            if not self._ascii_minus_is_content_prefix(tmp):
                bullet_match = re.match(r"^[\s\u25A0\u25AA\u25CF\u2022\u2027\u00B7\-\*]+", tmp)
                if bullet_match:
                    skip_count += len(bullet_match.group(0))

        parts: List[str] = []
        current_key: Optional[Tuple[bool, bool]] = None
        current_buf: List[str] = []
        remaining_skip = skip_count

        try:
            char_count = int(rng.Characters.Count)
        except Exception:
            return ""

        for idx in range(1, char_count + 1):
            try:
                ch = rng.Characters(idx)
                txt = str(ch.Text)
            except Exception:
                continue

            txt = txt.replace("\r", "").replace("\x07", "").replace("\x0b", "")
            txt = txt.replace("\t", " ").replace("\xa0", " ")
            if not txt:
                continue

            if remaining_skip > 0:
                if len(txt) <= remaining_skip:
                    remaining_skip -= len(txt)
                    continue
                txt = txt[remaining_skip:]
                remaining_skip = 0
                if not txt:
                    continue

            try:
                bold = bool(int(ch.Bold) != 0)
            except Exception:
                bold = False
            try:
                underline = bool(int(ch.Underline) != 0)
            except Exception:
                underline = False

            key = (bold, underline)
            if current_key is None:
                current_key = key
            if key != current_key:
                parts.append(self._wrap_run("".join(current_buf), current_key[0], current_key[1]))
                current_buf = []
                current_key = key
            current_buf.append(html.escape(txt))

        if current_buf and current_key is not None:
            parts.append(self._wrap_run("".join(current_buf), current_key[0], current_key[1]))

        rendered = "".join(parts)
        return self._cleanup_inline_html(rendered)

    def _wrap_run(self, text: str, bold: bool, underline: bool) -> str:
        if not text:
            return ""
        if bold and underline:
            return f"<strong><u>{text}</u></strong>"
        if bold:
            return f"<strong>{text}</strong>"
        if underline:
            return f"<u>{text}</u>"
        return text

    def _cleanup_inline_html(self, html_text: str) -> str:
        html_text = self.re_space.sub(" ", html_text)
        html_text = re.sub(r"(?:<strong>\s*</strong>|<u>\s*</u>)", "", html_text)
        html_text = re.sub(r"<strong><strong>(.*?)</strong></strong>", r"<strong>\1</strong>", html_text)
        html_text = re.sub(r"<u><u>(.*?)</u></u>", r"<u>\1</u>", html_text)
        return html_text.strip()

    def _normalize_text(self, text: str) -> str:
        return self.re_space.sub(" ", text).strip()

    def _render_subheading(self, inner_html: str) -> str:
        s = self.style_cfg
        return (
            '<p style="margin:{mt}em {mr} {mb}em {ml}; text-indent:{ti}; text-align:{ta}; line-height:{lh};">'
            '<span style="display:inline-block; max-width:100%; vertical-align:top; '
            'padding:{pv}em {ph}em; font-family:{ff}; font-size:{fs}px; font-weight:{fw}; '
            'line-height:{lh}; color:{tc}; background:{bg}; border-radius:{br}px; '
            'white-space:normal; text-align:{ta}; word-break:break-word; overflow-wrap:break-word; '
            'box-sizing:border-box;">{inner}</span></p>'
        ).format(
            mt=s.get("subheading_margin_top_em", 1.15),
            mr=s.get("subheading_margin_right_css", "0"),
            mb=s.get("subheading_margin_bottom_em", 0.55),
            ml=s.get("subheading_margin_left_css", "0"),
            ti=s.get("subheading_text_indent_css", "0"),
            ta=s.get("subheading_text_align", "left"),
            lh=s["subheading_line_height"],
            pv=s["subheading_padding_vertical_em"],
            ph=s["subheading_padding_horizontal_em"],
            ff=s["subheading_font_family"],
            fs=s["subheading_font_size_px"],
            fw=s.get("subheading_font_weight", 700),
            tc=s["subheading_text_color"],
            bg=s["subheading_bg_color"],
            br=s["subheading_border_radius_px"],
            inner=inner_html,
        )

    def _render_bullet_list_item(self, inner_html: str) -> str:
        s = self.style_cfg
        return (
            '<li style="margin-top:{mt}px;">'
            '<p style="margin:{pm}; text-indent:{pti}; text-align:{pta}; line-height:{lh};">{inner}</p>'
            '</li>'
        ).format(
            mt=s["bullet_item_margin_top_px"],
            pm=s.get("bullet_paragraph_margin_css", "0"),
            pti=s.get("bullet_paragraph_text_indent_css", "0"),
            pta=s.get("bullet_paragraph_text_align", "left"),
            lh=s["bullet_line_height"],
            inner=inner_html,
        )

    def _render_bullet_list(self, li_blocks: List[str]) -> str:
        s = self.style_cfg
        return (
            '<ul class="list-paddingleft-2" '
            'style="list-style-type:{lst}; margin:{mt} {mr} {mb}em {ml}; padding-left:{pl}em;">'
            '{items}'
            '</ul>'
        ).format(
            lst=s.get("bullet_list_style_type", "square"),
            mt=s.get("bullet_list_margin_top_css", "0"),
            mr=s.get("bullet_list_margin_right_css", "0"),
            mb=s["bullet_list_margin_bottom_em"],
            ml=s.get("bullet_list_margin_left_css", "0"),
            pl=s["bullet_list_padding_left_em"],
            items="".join(li_blocks),
        )

    def _wrap_full_html(self, body_html: str) -> str:
        font_family = self.out_cfg["font_family"]
        font_size = self.out_cfg["font_size_px"]
        line_height = self.out_cfg["line_height"]
        indent = self.out_cfg["first_line_indent_em"]
        margin_bottom = self.out_cfg["paragraph_margin_bottom_em"]
        page_margin = self.out_cfg.get("page_margin_css", "24px auto")
        page_padding = self.out_cfg.get("page_padding_css", "0 20px")
        page_max_width = self.out_cfg.get("page_max_width_px", 860)
        page_background = self.out_cfg.get("page_background_color", "#ffffff")
        page_text_color = self.out_cfg.get("page_text_color", "#000000")
        paragraph_margin_top = self.out_cfg.get("paragraph_margin_top_css", "0")
        paragraph_margin_right = self.out_cfg.get("paragraph_margin_right_css", "0")
        paragraph_margin_left = self.out_cfg.get("paragraph_margin_left_css", "0")
        bold_weight = self.out_cfg.get("inline_bold_font_weight", 700)
        underline_decoration = self.out_cfg.get("inline_underline_text_decoration", "underline")
        underline_offset = self.out_cfg.get("inline_underline_offset_em", 0.08)
        # body 显式 text-align:left：复制正文时结果页把计算对齐写进每个元素，未声明对齐的
        # 节点会写成 start，公众号《内容排版规范》1.6 判为非标准值。
        return f"""<!DOCTYPE html>
<html lang=\"zh-CN\">
<head>
<meta charset=\"utf-8\" />
<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
<title>summary_extract_output</title>
<style>
body {{
    margin: {page_margin};
    padding: {page_padding};
    max-width: {page_max_width}px;
    background: {page_background};
    color: {page_text_color};
    font-family: {font_family};
    font-size: {font_size}px;
    line-height: {line_height};
    text-align: left;
    word-break: break-word;
}}
p {{
    margin: {paragraph_margin_top} {paragraph_margin_right} {margin_bottom}em {paragraph_margin_left};
    text-indent: {indent}em;
}}
strong {{
    font-weight: {bold_weight};
}}
u {{
    text-decoration: {underline_decoration};
    text-underline-offset: {underline_offset}em;
}}
</style>
</head>
<body>
{body_html}
</body>
</html>
"""


class Block:
    def __init__(self, tag: Tag, idx: int):
        self.tag = tag
        self.idx = idx
        self.text = normalize_text(tag.get_text(" ", strip=True))
        self.has_image = bool(
            tag.find(
                lambda t: isinstance(t, Tag)
                and (t.name or "").lower() in {"img", "v:imagedata"}
            )
        )
        cls = " ".join(tag.get("class", [])) if tag.get("class") else ""
        style = tag.get("style", "") or ""
        self.class_text = cls
        self.style_text = style
        self.is_toc = False
        self.drop = False
        self.render_failure: Optional[Exception] = None   # 预处理期这一块出错：渲染时只留空行并定位告警
        self.role = "normal"
        self.figure_segments = []
        self.heading_level = 0
        # 数据表格 -> 图片 渲染所需
        self.table_build = None        # _build_table_spec() 结果(spec + 题注/来源单元格)
        self.table_png_path = None     # 渲染好的 PNG 路径缓存(None=未渲染, ""=渲染失败)


class CleanResult:
    def __init__(self, html_text: str, docx_html_text: str, asset_map: Dict[Path, str]):
        self.html_text = html_text
        self.docx_html_text = docx_html_text
        self.asset_map = asset_map


def normalize_text(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def loose_text_key(text: str) -> str:
    text = normalize_text(text).lower()
    text = re.sub(r"[\s\u3000]+", "", text)
    text = re.sub(r"[()（）\[\]【】{}<>《》:：;；,，.。!！?？、\"'‘’“”\-—–_/\\]+", "", text)
    return text


def heading_text_match_score(block_text: str, outline_text: str) -> int:
    bt = normalize_text(block_text)
    ot = normalize_text(outline_text)
    if not bt or not ot:
        return 0
    if bt == ot:
        return 100
    bkey = loose_text_key(bt)
    okey = loose_text_key(ot)
    if not bkey or not okey:
        return 0
    if bkey == okey:
        return 95
    if bt.startswith(ot) or ot.startswith(bt):
        return 85
    if bkey.startswith(okey) or okey.startswith(bkey):
        return 80
    if ot in bt or bt in ot:
        return 70
    if okey in bkey or bkey in okey:
        return 65
    return 0


def extract_docx_heading_outline(docx_path: Path, matching_cfg: dict) -> List[dict]:
    if docx_path.suffix.lower() != ".docx":
        return []
    try:
        import zipfile
        import xml.etree.ElementTree as ET
    except Exception:
        return []

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    def wanted_style_map() -> Dict[str, int]:
        result: Dict[str, int] = {}
        for level in (1, 2, 3):
            for name in matching_cfg.get(f"heading_level_{level}_style_names", []) or []:
                norm = normalize_text(str(name))
                if norm:
                    result[norm.lower()] = level
        return result

    target_names = wanted_style_map()
    if not target_names:
        return []

    try:
        with zipfile.ZipFile(str(docx_path), "r") as zf:
            styles_root = ET.fromstring(zf.read("word/styles.xml"))
            doc_root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        return []

    style_id_to_name: Dict[str, str] = {}
    for st in styles_root.findall("w:style", ns):
        if st.get("{%s}type" % ns["w"]) != "paragraph":
            continue
        style_id = st.get("{%s}styleId" % ns["w"])
        name_el = st.find("w:name", ns)
        if not style_id or name_el is None:
            continue
        style_name = normalize_text(name_el.get("{%s}val" % ns["w"]) or "")
        if style_name:
            style_id_to_name[style_id] = style_name

    outline: List[dict] = []
    for p in doc_root.findall(".//w:body/w:p", ns):
        ppr = p.find("w:pPr", ns)
        style_id = None
        if ppr is not None:
            pstyle = ppr.find("w:pStyle", ns)
            if pstyle is not None:
                style_id = pstyle.get("{%s}val" % ns["w"])
        style_name = normalize_text(style_id_to_name.get(style_id or "", ""))
        level = target_names.get(style_name.lower())
        if not level:
            continue
        text = normalize_text("".join((t.text or "") for t in p.findall(".//w:t", ns)))
        if not text:
            continue
        outline.append({"level": level, "text": text, "style_name": style_name})
    return outline


def heading_occurrence_key(text: str) -> str:
    """标题出现序列的对位键：先剥掉行首的圆圈数字类自动编号再取 loose_text_key。

    Word 自动编号（①②③…）不写进 docx 的 w:t 文本，但导出 HTML 时会拼进块文本，
    两侧必须用同一口径才能对上。只剥行首编号，文本内部的 ①② 保留。"""
    text = normalize_text(text)
    # 只剥数字类编号（①⑴⒈❶➀㈠㊀ 等），不剥圆圈拉丁字母(Ⓐ)与罗马数字(Ⅱ型糖尿病是内容)
    text = re.sub(r"^[①-⒛㉑-㉟㊱-㊿❶-➓㈠-㈩㊀-㊉]+[\s.、．]*", "", text)
    return loose_text_key(text)


def extract_docx_heading_text_occurrences(docx_path: Path, matching_cfg: dict) -> Dict[str, List[int]]:
    """源 docx 顶层段落里「同一文本既作正文段又作标题段」的出现序列。

    返回 {loose_text_key: [每次出现的标题级别]}，0 表示该次出现是正文段；只保留
    混合出现（既有 0 又有 >0）的键。Word 导出的 HTML 丢失自定义标题样式名后，
    _annotate_heading_levels 只能靠大纲文本模糊匹配识别标题——正文里逐字重复标题
    文本的段落（如问卷报告开头的题目清单）文本上无法区分，出现次序可以。"""
    if docx_path.suffix.lower() != ".docx":
        return {}
    try:
        import zipfile
        import xml.etree.ElementTree as ET
    except Exception:
        return {}

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    target_names: Dict[str, int] = {}
    for level in (1, 2, 3):
        for name in matching_cfg.get(f"heading_level_{level}_style_names", []) or []:
            norm = normalize_text(str(name))
            if norm:
                target_names[norm.lower()] = level
    if not target_names:
        return {}

    try:
        with zipfile.ZipFile(str(docx_path), "r") as zf:
            styles_root = ET.fromstring(zf.read("word/styles.xml"))
            doc_root = ET.fromstring(zf.read("word/document.xml"))
    except Exception:
        return {}

    style_id_to_name: Dict[str, str] = {}
    for st in styles_root.findall("w:style", ns):
        if st.get("{%s}type" % ns["w"]) != "paragraph":
            continue
        style_id = st.get("{%s}styleId" % ns["w"])
        name_el = st.find("w:name", ns)
        if not style_id or name_el is None:
            continue
        style_name = normalize_text(name_el.get("{%s}val" % ns["w"]) or "")
        if style_name:
            style_id_to_name[style_id] = style_name

    occurrences: Dict[str, List[int]] = {}
    for p in doc_root.findall(".//w:body/w:p", ns):
        ppr = p.find("w:pPr", ns)
        style_id = None
        if ppr is not None:
            pstyle = ppr.find("w:pStyle", ns)
            if pstyle is not None:
                style_id = pstyle.get("{%s}val" % ns["w"])
        style_name = normalize_text(style_id_to_name.get(style_id or "", ""))
        level = target_names.get(style_name.lower(), 0)
        key = heading_occurrence_key("".join((t.text or "") for t in p.findall(".//w:t", ns)))
        if not key:
            continue
        occurrences.setdefault(key, []).append(level)
    return {key: levels for key, levels in occurrences.items()
            if any(lvl > 0 for lvl in levels) and any(lvl == 0 for lvl in levels)}


def detect_encoding(raw: bytes) -> str:
    head = raw[:2000].decode("ascii", errors="ignore")
    m = re.search(r"charset=([A-Za-z0-9_\-]+)", head, re.I)
    candidates = []
    if m:
        candidates.append(m.group(1))
    candidates += ["utf-8", "gbk", "cp936", "gb18030"]
    for enc in candidates:
        try:
            raw.decode(enc)
            return enc
        except Exception:
            continue
    return "utf-8"


def read_text_auto(path: Path) -> str:
    raw = path.read_bytes()
    enc = detect_encoding(raw)
    return raw.decode(enc, errors="replace")


def extract_color_style(style: str) -> str:
    if not style:
        return ""
    color_value = ""
    for item in style.split(";"):
        raw_item = item.strip()
        if not raw_item:
            continue
        pair = raw_item.split(":", 1)
        if len(pair) != 2:
            continue
        key = pair[0].strip().lower()
        value = pair[1].strip()
        if key == "color" and value:
            color_value = value
    return f"color: {color_value};" if color_value else ""


def iter_blocks(root: Tag) -> Iterable[Tag]:
    for child in root.children:
        if isinstance(child, NavigableString):
            continue
        if not isinstance(child, Tag):
            continue
        if child.name in {"script", "style", "meta", "link", "title", "head"}:
            continue
        if child.name in {"p", "ul", "ol", "table"}:
            yield child
        elif child.name in {"body", "div", "section", "article", "main", "center"}:
            yield from iter_blocks(child)


class HtmlCleaner:
    def __init__(self, config: dict):
        self.config = config
        self.matching = config["matching"]
        self.doc_cfg = config["document"]
        self.out_cfg = config["output"]

        self.re_main_heading = re.compile(self.matching["main_heading_regex"])
        self.re_heading_like = re.compile(self.matching["heading_like_regex"])
        self.re_heading_level_1 = re.compile(self.matching["heading_level_1_regex"])
        self.re_heading_level_2 = re.compile(self.matching["heading_level_2_regex"])
        self.re_heading_level_3 = re.compile(self.matching["heading_level_3_regex"])
        self.re_caption = re.compile(self.matching["caption_regex"])
        self.re_source = re.compile(self.matching["source_regex"])
        self.re_toc_entry = re.compile(self.matching["toc_entry_regex"])

        self.current_html_path: Optional[Path] = None
        self.asset_subdir_name: str = "wechat_clean_files"
        self.asset_map: Dict[Path, str] = {}
        self.used_asset_names: Dict[str, int] = {}
        self.heading_outline: List[dict] = []
        self.heading_text_occurrences: Dict[str, List[int]] = {}
        # 数据表格 -> 图片：渲染好的 PNG 放这里(每个文件处理前由 process_one_file 设置)
        self.generated_dir: Optional[Path] = None
        self.source_docx_path: Optional[Path] = None
        self._valuation_cache: Optional[dict] = None
        # 本次清洗中「源里有图、成品里没有」的结构化记录，每条带源块序号 idx。
        # 渲染器一贯的策略是失败即跳过并继续，成品因此可能少一张图而外观完好；
        # 点评附三表那一遍据此判断摘出的区块是否完整（见 convert_worker）。序号用来
        # 把判断限定在三表区块自身，报告别处的问题不牵连它。只记录，不改渲染结果。
        self.render_problems: List[dict] = []
        self.failed_blocks: List[dict] = []
        self._current_heading = ""
        self._previous_caption = ""
        self._heading_at: Dict[int, str] = {}
        self._current_block: Optional[Block] = None
        # 每个进入成品的源块的 (序号, 归一化文字)，供调用方按文字定位区块起点。
        self.rendered_block_texts: List[Tuple[int, str]] = []
        self._current_block_idx = -1
        # 本次清洗中「真正被成品用掉」的图片，按 (源块序号, src) 计数。计数而不是集合：
        # 同一地址在一块里出现两次、只用掉一次时，集合会把另一次也当成用过了。
        # 记在「用掉」而不是「读到」的时点：渲染器可能读了图又把它所在的行剪掉
        # （表格的题注/来源行），也可能整支子树都不走（wrapper 表选中内层网格转图），
        # 两种丢法在取址出口上都看不见，只能靠「源块里的图有没有被用掉」兜住。
        self._visited_image_counts: Counter = Counter()

    def _caption_label(self, text: str) -> str:
        match = re.match(r"\s*([图表])\s*(\d+)", text or "")
        return f"{match.group(1)}{match.group(2)}" if match and self.re_caption.search(text) else ""

    def _note_failed(self, block: Block, detail: str, error: str) -> None:
        # 同一块同一类问题只记一次（单元格图片在建表与核对时会被读到不止一遍）
        if any(item["idx"] == block.idx and item["detail"] == detail for item in self.failed_blocks):
            return
        kind ="表格" if (block.tag.name or "").lower() == "table" else ("图片" if block.has_image else "段落")
        # 图表序号只用于定位失败的图表；文字段落失败按页码与章节定位
        caption = ""
        if kind != "段落":
            built = getattr(block, "table_build", None) or {}
            caption_text = BeautifulSoup(built.get("caption_html") or "", "html.parser").get_text(" ", strip=True)
            caption = self._caption_label(caption_text) or self._caption_label(block.text) or self._previous_caption
        self.failed_blocks.append({"idx": block.idx, "kind": kind, "opening": block.text[:24], "caption": caption,
                                   "detail": detail, "error": error})
        self._note_render_problem(f"{detail}：{error}")

    def _render_failed_block(self, block: Block, exc: Exception) -> str:
        self._note_failed(block, "未能转换，原位置只留一个空行", f"{type(exc).__name__}: {exc}")
        # 只留空行、不放占位文字：误发出去时空行几乎无害，占位句会直接出现在读者面前
        return self._render_body_paragraph_html("<br/>")

    def report_failed_blocks(self, pages: Optional[Dict[int, List[int]]] = None) -> None:
        """失败块逐条告警（结果页红色信息条）：Word 页码、章节、图表序号、块序号与开头文字。
        章节与开头文字放进「」：脱敏错误报告整行剔除含「」的行，不把正文带进报告。"""
        for item in self.failed_blocks:
            where = []
            hits = (pages or {}).get(item["idx"]) or []
            if hits:
                where.append("Word 第 " + "、".join(str(page) for page in hits) + " 页")
            chapter = self._heading_at.get(item["idx"], "")
            if chapter:
                where.append(f"「{chapter[:30]}」一节")
            if item["caption"]:
                where.append(item["caption"])
            where.append(f"正文第 {item['idx']} 块（{item['kind']}，开头「{item['opening']}」）")
            print(f"  [warn] {'、'.join(where)}{item['detail']}，请对照原文人工补充：{item['error']}", flush=True)

    def _note_render_problem(self, message: str) -> None:
        self.render_problems.append({"idx": self._current_block_idx, "message": message})

    def _note_visited_image(self, src: str) -> None:
        self._visited_image_counts[(self._current_block_idx, (src or "").strip())] += 1

    @staticmethod
    def _image_srcs(tag: Tag) -> List[str]:
        srcs = []
        for node in tag.find_all(True):
            if (node.name or "").lower() in {"img", "v:imagedata"}:
                src = node.get("src", "") or node.get("o:href", "") or node.get("href", "")
                srcs.append((src or "").strip())
        return srcs

    def set_heading_outline(self, outline: List[dict]) -> None:
        self.heading_outline = list(outline or [])

    def set_heading_text_occurrences(self, occurrences: Dict[str, List[int]]) -> None:
        self.heading_text_occurrences = dict(occurrences or {})

    def clean_html_file(self, html_path: Path, asset_subdir_name: str,
                        summary_html: str = "", toc_png_path: Optional[str] = None) -> CleanResult:
        self.current_html_path = html_path
        self.asset_subdir_name = asset_subdir_name
        self.asset_map = {}
        self.used_asset_names = {}
        # 单元格图片(base64)落盘缓存按文件重置：folder 模式下 HtmlCleaner 复用同一实例，
        # 上一文件的 temp 目录已在 finally 删除，不能让下一文件命中已删除的旧路径。
        self._img_data_cache = {}
        self._img_tmp_counter = 0
        self.render_problems = []
        self.rendered_block_texts = []
        self._current_block_idx = -1
        self._visited_image_counts = Counter()

        html_text = read_text_auto(html_path)
        soup = BeautifulSoup(html_text, "html.parser")
        body = soup.body or soup
        blocks = [Block(tag, idx) for idx, tag in enumerate(iter_blocks(body))]

        if not blocks:
            raise ProcessingError(f"No usable content found in HTML: {html_path}")

        self._current_block_idx = -1
        self._current_block = None
        self.failed_blocks = []
        self._heading_at = {}
        self._mark_toc_blocks(blocks)
        start_idx = self._find_start_index(blocks)
        stop_idx = self._find_stop_index(blocks, start_idx)
        self._mark_out_of_range(blocks, start_idx, stop_idx)
        self._mark_drop_patterns(blocks)
        self._mark_media_related_blocks(blocks)
        self._annotate_heading_levels(blocks)

        # 只有需要 docx 输出时才跑 for_docx 渲染。它会把图片/表图注册成**外链文件**资产
        # (_register_asset inline=False)，从而触发 copy_used_assets 复制到输出目录；
        # 纯 HTML(base64 内联)用不到这些外链资产，跳过可避免生成多余的 *_files 资产文件夹。
        need_docx_html = bool(self.out_cfg.get("generate_docx", False))
        rendered_blocks = []
        rendered_blocks_docx = []
        self._current_heading = ""
        self._previous_caption = ""
        # 导出 HTML 会在中文与数字间多出空格、标题前多出自动编号：去掉空白再比较，章节名用大纲原文
        heading_texts = {re.sub(r"\s+", "", entry.get("text", "")): entry.get("text", "").strip()
                         for entry in self.heading_outline if (entry.get("text") or "").strip()}
        for block in blocks:
            if block.drop:
                continue
            self._current_block_idx = block.idx
            # 不在这里清空 _visited_image_counts：表格规格是在渲染循环之前的
            # _mark_media_related_blocks 里建好的，清空会把那批用图记录抹掉
            compact = re.sub(r"\s+", "", block.text)
            for key, heading in heading_texts.items():
                if compact == key or (compact.endswith(key) and len(compact) - len(key) <= 6):
                    self._current_heading = heading
                    break
            self._heading_at[block.idx] = self._current_heading
            self._current_block = block
            try:
                if getattr(block, "render_failure", None) is not None:
                    raise block.render_failure
                rendered = self._render_block(block, for_docx=False)
            except Exception as exc:
                # 单块出错不让整份作废：原位置只留空行，告警（结果页红色信息条）点明第几块，
                # 并记进 render_problems，使三表附加等要求完整的调用方照常放弃
                rendered = self._render_failed_block(block, exc)
            rendered = rendered.strip()
            self._previous_caption = self._caption_label(block.text)
            source_srcs = self._image_srcs(block.tag)
            # 源块里有图、渲染结果却一张也没有：这一块整个消失了（如 remove_tables 下
            # 的单格图片表），既不报错也不打日志，只有这里能记下来
            if source_srcs and "<img" not in rendered:
                self._note_render_problem("源块中的图片没有进入成品")
            # 源块里有几张图，成品就该用掉几张。逐地址比对出现次数与用掉次数：
            # 分类器只走一支子树、题注/来源行连图一起被剪掉，都会在这里露出来
            for src, needed in Counter(source_srcs).items():
                used = self._visited_image_counts.get((block.idx, src), 0)
                if used < needed:
                    self._note_render_problem(
                        f"源块中的图片没有被成品用掉（{needed} 张里只用了 {used} 张）：{src[:120]!r}")
            if rendered:
                rendered_blocks.append(rendered)
                self.rendered_block_texts.append((block.idx, normalize_text(block.tag.get_text(" ", strip=True))))
            if need_docx_html:
                rendered_docx = self._render_block(block, for_docx=True)
                rendered_docx = rendered_docx.strip()
                if rendered_docx:
                    rendered_blocks_docx.append(rendered_docx)

        if not rendered_blocks:
            raise ProcessingError("Nothing left after cleaning. Check config or document structure.")

        # 出口兜底：确定意味着缺图的内部记录（图片解析不到、未能内嵌、单元格图片取不到、表图或贴入表渲染失败）
        # 却还没告警的块，补一条定位告警。「源块中的图片没有进入成品 / 没被用掉」是给三表判定用的从严记账，
        # 在较早模板上噪声多（55 份批量中 18 份各有 8~35 条），不升级为告警
        by_idx = {block.idx: block for block in blocks}
        reported = {item["idx"] for item in self.failed_blocks}
        for problem in list(self.render_problems):
            if not problem["message"].startswith(LOSS_PROBLEM_PREFIXES):
                continue
            block = by_idx.get(problem["idx"])
            if block is not None and block.idx not in reported:
                reported.add(block.idx)
                self._note_failed(block, "中有图片或表格未能完整进入成品（缺失处留空）", problem["message"])

        # 前置内容(摘要 + 内容目录图)拼到正文之前。资产注册在本方法的资产生命周期内完成
        # (current_html_path / asset_map 已就绪), 故必须在这里拼, 不能在 clean_html_file 之外。
        front_html = self._build_front_matter(summary_html, toc_png_path, for_docx=False)
        rendered_blocks = front_html + rendered_blocks
        # 网页成品只认内嵌图：没能换成 data: 的本机图片地址粘贴后是破图、还带出本机路径，删掉只留空白（已在上面告警）
        rendered_blocks = [LOCAL_IMG_RE.sub("", part) for part in rendered_blocks]
        if need_docx_html:
            front_docx = self._build_front_matter(summary_html, toc_png_path, for_docx=True)
            rendered_blocks_docx = front_docx + rendered_blocks_docx

        return CleanResult(
            self._wrap_full_html("\n".join(rendered_blocks)),
            self._wrap_full_html("\n".join(rendered_blocks_docx or rendered_blocks)),
            dict(self.asset_map),
        )

    @staticmethod
    def _block_has_image(tag: Tag) -> bool:
        return tag.find(
            lambda t: isinstance(t, Tag) and (t.name or "").lower() in {"img", "v:imagedata"}
        ) is not None

    def _mark_toc_blocks(self, blocks: List[Block]) -> None:
        if not self.doc_cfg.get("remove_toc", True):
            return
        for block in blocks:
            class_text = (block.class_text or "") + " " + (block.style_text or "")
            class_text = class_text.lower()
            text = block.text
            if "toc" in class_text or "mso_toc" in class_text:
                block.is_toc = True
            if self.re_toc_entry.search(text):
                block.is_toc = True
            if text in self.doc_cfg.get("drop_title_patterns", []):
                block.is_toc = True
            compact_text = re.sub(r"\s+", "", text)
            if compact_text in {"\u76ee\u5f55", "\u5185\u5bb9\u76ee\u5f55", "\u56fe\u8868\u76ee\u5f55"}:
                block.is_toc = True

    @staticmethod
    def _looks_like_toc_entry(block: Block) -> bool:
        """图表目录条目不一定被识别成 is_toc：指向 `_Toc` 书签的链接、图表目录样式也算目录块。"""
        return ("tableoffigures" in (block.class_text or "").lower()
                or block.tag.find("a", href=re.compile(r"^#_Toc")) is not None)

    def _find_start_index(self, blocks: List[Block]) -> int:
        mode = self.doc_cfg.get("start_mode", "first_main_heading")
        if mode == "first_main_heading":
            for i, block in enumerate(blocks):
                if block.is_toc:
                    continue
                text = block.text
                if not text:
                    continue
                if self.re_main_heading.search(text):
                    return self._start_after_toc(blocks, i)
        for i, block in enumerate(blocks):
            if not block.is_toc and block.text:
                return i
        return 0

    def _start_after_toc(self, blocks: List[Block], heading_idx: int) -> int:
        """第一个主标题之前有目录时，正文从最后一个目录块之后的第一个有内容的块起算：
        目录与第一章之间的导语和图表属于正文。首页摘要在目录之前，不受影响；没有目录时仍从主标题起。"""
        toc_like = [j for j in range(heading_idx)
                    if blocks[j].is_toc or self._looks_like_toc_entry(blocks[j])]
        if not toc_like:
            return heading_idx
        for j in range(toc_like[-1] + 1, heading_idx):
            if blocks[j].text or self._block_has_image(blocks[j].tag):
                return j
        return heading_idx

    def _find_stop_index(self, blocks: List[Block], start_idx: int) -> int:
        stop_patterns = [p.strip() for p in self.doc_cfg.get("stop_at_patterns", []) if p.strip()]
        if not stop_patterns:
            return len(blocks)
        number_token = r"(?:\d+(?:\.\d+)*|[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e]+)"
        numbered_prefix = (
            r"^\s*(?:"
            r"\u7b2c\s*" + number_token + r"\s*(?:\u7ae0|\u8282|\u90e8\u5206)\s*"
            r"|[\uff08(]\s*" + number_token + r"\s*[\uff09)]\s*"
            r"|" + number_token + r"(?:\s*[\u3001.\uff0e:\uff1a\uff09)]\s*|\s+)"
            r")"
        )
        for i in range(start_idx, len(blocks)):
            text = blocks[i].text
            if not text:
                continue
            compact_length = len(re.sub(r"\s+", "", text))
            for pat in stop_patterns:
                if pat not in text:
                    continue
                if (compact_length <= 30
                        or re.match(r"^\s*" + re.escape(pat), text)
                        or re.match(numbered_prefix + re.escape(pat), text)):
                    return i
        return len(blocks)

    def _mark_out_of_range(self, blocks: List[Block], start_idx: int, stop_idx: int) -> None:
        for i, block in enumerate(blocks):
            if i < start_idx or i >= stop_idx:
                block.drop = True
        for block in blocks:
            if block.is_toc:
                block.drop = True

    def _mark_drop_patterns(self, blocks: List[Block]) -> None:
        patterns = self.doc_cfg.get("drop_exact_or_contains_patterns", [])
        for block in blocks:
            if block.drop:
                continue
            text = block.text
            for pat in patterns:
                if pat and pat in text:
                    block.drop = True
                    break

    def _cell_text(self, cell: Optional[Tag]) -> str:
        if cell is None:
            return ""
        return normalize_text(cell.get_text(" ", strip=True))

    def _cell_images_html(self, cell: Optional[Tag], for_docx: bool = False) -> List[str]:
        if cell is None:
            return []
        images = []
        for node in cell.find_all(True):
            name = (node.name or "").lower()
            if name not in {"img", "v:imagedata"}:
                continue
            src = node.get("src", "") or node.get("o:href", "") or node.get("href", "")
            resolved = self._register_asset(src, inline=not for_docx)
            if not resolved:
                continue
            alt = html.escape(node.get("alt", "") or "")
            image_margin = self.out_cfg.get("image_margin_css", "0 auto")
            image_border = self.out_cfg.get("image_border_css", "0")
            images.append(
                f'<img src="{html.escape(resolved)}" alt="{alt}" '
                f'style="display:block; max-width:100%; height:auto; margin:{image_margin}; border:{image_border};" />'
            )
        return images

    def _table_rows_expanded(self, table: Tag, direct_only: bool = False) -> List[List[Tag]]:
        rows: List[List[Tag]] = []
        for tr in (self._table_direct_rows(table) if direct_only else table.find_all("tr")):
            row_cells: List[Tag] = []
            for cell in tr.find_all(["td", "th"], recursive=False):
                try:
                    colspan = int((cell.get("colspan", "1") or "1").strip())
                except Exception:
                    colspan = 1
                colspan = max(1, colspan)
                for _ in range(colspan):
                    row_cells.append(cell)
            if row_cells:
                rows.append(row_cells)
        return rows

    def _cell_has_image(self, cell: Optional[Tag]) -> bool:
        if cell is None:
            return False
        return cell.find(
            lambda t: isinstance(t, Tag) and (t.name or "").lower() in {"img", "v:imagedata"}
        ) is not None

    def _cell_pasted_grid_table(self, cell: Optional[Tag]) -> Optional[Tag]:
        """图框单元格里「以表代图」贴进去的嵌套表格。

        单元格无图片、嵌套表 >=2 行且至少一行 >=2 格才算——通用渲染器要求 >=2 列，
        单列结构认作图会在渲染时静默失败，不如留在原路径。"""
        if cell is None or self._cell_has_image(cell):
            return None
        nested = cell.find("table")
        if nested is None:
            return None
        nested_rows = self._table_direct_rows(nested)
        if len(nested_rows) < 2:
            return None
        if not any(len(tr.find_all(["td", "th"], recursive=False)) >= 2 for tr in nested_rows):
            return None
        return nested

    def _ensure_pasted_table_png(self, nested: Tag) -> Optional[str]:
        """把贴进图框的嵌套表格用通用表格渲染器转成 PNG(按节点缓存，只渲染一次)。"""
        cache = getattr(self, "_pasted_table_pngs", None)
        if cache is None:
            cache = self._pasted_table_pngs = {}
        key = id(nested)
        if key in cache:
            return cache[key] or None
        png_path = ""
        if self.generated_dir:
            try:
                spec = self._grid_rows_to_spec(self._table_direct_rows(nested), set())
                if spec is None:
                    print("  -> figure-frame pasted table skipped: unrenderable grid")
                    self._note_render_problem("图框内贴入的表格无法渲染（网格不可解析）")
                if spec is not None:
                    import sys
                    if str(SCRIPT_DIR) not in sys.path:
                        sys.path.insert(0, str(SCRIPT_DIR))
                    import table_image as ti

                    out_png = Path(self.generated_dir) / f"figure_pasted_table_{len(cache)}.png"
                    missing: List[str] = []
                    ti.render_table(spec, str(out_png), missing=missing)
                    png_path = str(out_png)
                    print(f"  -> figure-frame pasted table image: {spec['nrows']} rows x {spec['ncols']} cols")
                    if missing and self._current_block is not None:
                        self._note_failed(self._current_block, f"图框内贴入的表格中有 {len(missing)} 张单元格图片未能读取（表图中该格留空）",
                                          "image unreadable")
            except Exception as e:
                print(f"  -> figure-frame pasted table skipped: {e}")
                if self._current_block is not None:
                    self._note_failed(self._current_block, "图框内贴入的表格未能转成图片（该处留空）", f"{type(e).__name__}: {e}")
                else:
                    self._note_render_problem(f"图框内贴入的表格渲染失败：{e}")
        cache[key] = png_path
        return png_path or None

    def _pasted_table_img_html(self, png: str, inline: bool) -> str:
        resolved = self._register_asset(Path(png).as_uri(), inline=inline)
        if not resolved:
            return ""
        image_margin = self.out_cfg.get("image_margin_css", "0 auto")
        image_border = self.out_cfg.get("image_border_css", "0")
        return (
            f'<img src="{html.escape(resolved)}" alt="" '
            f'style="display:block; max-width:100%; height:auto; margin:{image_margin}; border:{image_border};" />'
        )

    def _extract_figure_segments_from_table(self, table: Tag) -> List[dict]:
        # 整张图框被套在「单行单格」wrapper 里时先逐层解包，否则直属行口径只见一行、
        # 图框被静默丢弃（题注/图片/来源全丢）。
        while True:
            direct = self._table_direct_rows(table)
            if len(direct) == 1:
                cells = direct[0].find_all(["td", "th"], recursive=False)
                if len(cells) == 1:
                    inner = cells[0].find("table")
                    if inner is not None:
                        table = inner
                        continue
            break

        # 只取图框自己的直属行：嵌套表格（「以表代图」贴进图框的表）的内部行
        # 不得混进题注/注/来源行的定位，否则内表文字会被当注释挂到图片下。
        rows = self._table_rows_expanded(table, direct_only=True)
        if len(rows) < 2:
            return []

        # 「以表代图」兼容：仅当图框里同时存在真实图片时，才把嵌套表格单元格当作
        # 图片格参与分组(如 图15 为照片、图16 被作者贴成表格)。整个图框都没有真实
        # 图片时保持原行为(交给 data_table 的嵌套网格路径处理)。
        has_real_image = any(self._cell_has_image(c) for row in rows for c in row)

        def cell_figure_like(c: Optional[Tag]) -> bool:
            if self._cell_has_image(c):
                return True
            return has_real_image and self._cell_pasted_grid_table(c) is not None

        # 按内容定位"题注行/图片行/数据来源行"，而不是写死 rows[0]/[1]/[2]：
        # Word 导出的 HTML 可能在表格里插入空白行，固定行号会错位导致
        # 题注或数据来源(图X下方的注)丢失。
        # 同一张表里可能竖向叠放多组图(图3 在上、图4 在下)，故遍历每个题注行，
        # 把它与"到下一个题注行之前"的图片行/数据来源行配成一组。
        cap_indices = [i for i, r in enumerate(rows)
                       if any(self.re_caption.search(self._cell_text(c)) for c in r)]
        if not cap_indices:
            return []

        segments = []
        for gi, cap_idx in enumerate(cap_indices):
            next_cap = cap_indices[gi + 1] if gi + 1 < len(cap_indices) else len(rows)
            img_indices = [i for i in range(cap_idx + 1, next_cap)
                           if any(cell_figure_like(c) for c in rows[i])]
            if not img_indices:
                continue
            src_indices = [i for i in range(cap_idx + 1, next_cap)
                           if any(self.re_source.search(self._cell_text(c)) for c in rows[i])]
            note_indices = [
                i for i in range(cap_idx + 1, next_cap)
                if i not in img_indices
                and i not in src_indices
                and any(self._cell_text(c) for c in rows[i])
            ]

            cap_row = rows[cap_idx]
            img_rows = [rows[i] for i in img_indices]
            note_rows = [rows[i] for i in note_indices]
            src_rows = [rows[i] for i in src_indices]

            max_cols = max(len(row) for row in [cap_row, *img_rows, *note_rows, *src_rows])

            def unique_ranges(row: List[Tag]):
                ranges = []
                seen = set()
                for col, cell in enumerate(row):
                    key = id(cell)
                    if key in seen:
                        continue
                    seen.add(key)
                    end = col + 1
                    while end < len(row) and row[end] is cell:
                        end += 1
                    ranges.append((col, end, cell))
                return ranges

            cap_ranges = [(a, b, c) for a, b, c in unique_ranges(cap_row)
                          if self.re_caption.search(self._cell_text(c))]
            img_ranges = [(a, b, c) for row in img_rows for a, b, c in unique_ranges(row)
                          if cell_figure_like(c)]
            note_ranges = [(a, b, c) for row in note_rows for a, b, c in unique_ranges(row)
                           if self._cell_text(c)]
            src_ranges = [(a, b, c) for row in src_rows for a, b, c in unique_ranges(row)
                          if self.re_source.search(self._cell_text(c))]

            parent = list(range(max_cols))

            def find(col: int) -> int:
                while parent[col] != col:
                    parent[col] = parent[parent[col]]
                    col = parent[col]
                return col

            def union(left: int, right: int) -> None:
                left_root, right_root = find(left), find(right)
                if left_root != right_root:
                    parent[right_root] = left_root

            for start, end, _ in cap_ranges + img_ranges + note_ranges + src_ranges:
                for col in range(start + 1, end):
                    union(start, col)

            groups: Dict[int, List[int]] = {}
            for col in range(max_cols):
                if any(start <= col < end for start, end, _ in img_ranges):
                    groups.setdefault(find(col), []).append(col)

            def overlapping(ranges, cols):
                left, right = min(cols), max(cols) + 1
                return [(start, end, cell) for start, end, cell in ranges
                        if start < right and end > left]

            def cell_inline(cell: Tag, preserve_color: bool) -> str:
                return self._cleanup_inline_html(
                    self._render_inline_children(cell, preserve_color=preserve_color))

            for cols in sorted(groups.values(), key=min):
                caps = overlapping(cap_ranges, cols)
                imgs = overlapping(img_ranges, cols)
                notes = overlapping(note_ranges, cols)
                sources = overlapping(src_ranges, cols)
                images_html = []
                images_docx = []
                for _, _, cell in imgs:
                    cell_web = self._cell_images_html(cell, for_docx=False)
                    if cell_web:
                        images_html.extend(cell_web)
                        images_docx.extend(self._cell_images_html(cell, for_docx=True))
                        continue
                    nested = self._cell_pasted_grid_table(cell)
                    png = self._ensure_pasted_table_png(nested) if nested is not None else None
                    if png:
                        web = self._pasted_table_img_html(png, inline=True)
                        docx = self._pasted_table_img_html(png, inline=False)
                        if web and docx:
                            images_html.append(web)
                            images_docx.append(docx)
                image_count = max(len(images_html), len(images_docx))
                if not image_count or not caps:
                    continue

                cap_html = [(start, end, cell_inline(cell, True)) for start, end, cell in caps]
                note_html = [(start, end, cell_inline(cell, False)) for start, end, cell in notes]
                src_html = [(start, end, cell_inline(cell, False)) for start, end, cell in sources]
                pair_items = (
                    image_count > 1
                    and len(cap_html) == image_count
                    and len(src_html) in {0, 1, image_count}
                    and len(images_html) == len(images_docx) == image_count
                )
                if pair_items:
                    note_for = [""] * image_count
                    if len(note_html) == image_count:
                        note_for = [x[2] for x in note_html]
                    elif note_html:
                        note_for[-1] = self._cleanup_inline_html(
                            "<br/>".join(x[2] for x in note_html if x[2]))
                    src_for = [""] * image_count
                    if len(src_html) == image_count:
                        src_for = [x[2] for x in src_html]
                    elif len(src_html) == 1:
                        # 唯一来源格：按列跨度归属。只压住一个题注列时挂那张图；
                        # 横跨整组(合并来源行)时挂最后一张图，即整组底部。
                        s_start, s_end, s_inline = src_html[0]
                        covered = [i for i, (c_start, c_end, _) in enumerate(cap_html)
                                   if c_start < s_end and c_end > s_start]
                        src_for[covered[0] if len(covered) == 1 else image_count - 1] = s_inline
                    for item in range(image_count):
                        meta_html = self._cleanup_inline_html("<br/>".join(
                            value for value in (note_for[item], src_for[item]) if value
                        ))
                        segments.append({
                            "caption_html": cap_html[item][2],
                            "images_html": [images_html[item]],
                            "images_docx": [images_docx[item]],
                            "source_html": meta_html,
                        })
                    continue

                meta_html = self._cleanup_inline_html("<br/>".join(
                    x[2] for x in [*note_html, *src_html] if x[2]
                ))
                segments.append({
                    "caption_html": self._cleanup_inline_html("<br/>".join(x[2] for x in cap_html if x[2])),
                    "images_html": images_html,
                    "images_docx": images_docx,
                    "source_html": meta_html,
                })
        return segments

    def _table_first_row_has_table_caption(self, table: Tag) -> bool:
        rows = self._table_rows_expanded(table)
        if not rows:
            return False
        for cell in rows[0]:
            text = self._cell_text(cell)
            stripped = text.lstrip()
            if stripped.startswith("\u8868") and re.search(r"\d+[\uff1a:]", stripped):
                return True
        return False

    # ---------------- \u901a\u7528\u6570\u636e\u8868\u683c -> \u56fe\u7247 ----------------
    # \u628a\u6b63\u6587\u91cc\u7684\u771f\u5b9e\u6570\u636e\u8868\u683c(\u653f\u7b56\u8868/\u76c8\u5229\u9884\u6d4b\u8868/\u53ef\u6bd4\u516c\u53f8\u4f30\u503c\u8868\u2026)\u6574\u8868\u6e32\u67d3\u6210 PNG\uff0c
    # \u601d\u8def\u4e0e\u5468\u62a5\u7248\u300c\u4f30\u503c\u8868\u8f6c\u56fe\u7247\u300d\u4e00\u81f4\uff1a\u9898\u6ce8(\u8868X\uff1a)\u62bd\u5230\u56fe\u4e0a\u65b9\u3001\u6570\u636e\u6765\u6e90\u62bd\u5230\u56fe\u4e0b\u65b9\u3001
    # \u8868\u4f53\u8f6c\u56fe\u3002\u533a\u522b\u5728\u4e8e\u8fd9\u91cc\u7528**\u901a\u7528**\u7f51\u683c\u6e32\u67d3\u5668(table_image.render_table)\uff0c
    # \u5217\u6570/\u6a2a\u5411(colspan)\u4e0e\u7eb5\u5411(rowspan)\u5408\u5e76/\u8868\u5934\u884c\u6570\u90fd\u81ea\u9002\u5e94\uff0c\u517c\u5bb9\u6027\u8fdc\u5927\u4e8e\u4f30\u503c\u4e13\u8868\u3002
    # \u76f4\u63a5\u5403 Word \u5bfc\u51fa\u7684 HTML \u8868\u683c(filtered HTML \u4f1a\u4fdd\u7559 colspan/rowspan)\uff0c
    # \u56e0\u6b64\u65e0\u9700\u56de\u5230\u6e90 docx \u505a\u5339\u914d\u3002

    def _table_direct_rows(self, table: Tag) -> List[Tag]:
        """\u53ea\u53d6\u5c5e\u4e8e\u672c\u8868\u7684 <tr>(\u6392\u9664\u5355\u5143\u683c\u91cc\u5d4c\u5957\u8868\u683c\u7684\u884c)\u3002"""
        return [tr for tr in table.find_all("tr") if tr.find_parent("table") is table]

    def _plain_with_breaks(self, node) -> str:
        """\u5355\u5143\u683c\u7eaf\u6587\u672c\uff1a<br> -> \u6362\u884c\uff0c\u4e22\u5f03\u56fe\u7247/\u811a\u672c\u7b49\u3002"""
        if isinstance(node, NavigableString):
            return str(node).replace("\xa0", " ")
        if not isinstance(node, Tag):
            return ""
        name = (node.name or "").lower()
        if name in {"script", "style", "img", "object", "svg", "v:shape", "v:imagedata"}:
            return ""
        if name == "br":
            return "\n"
        return "".join(self._plain_with_breaks(ch) for ch in node.children)

    def _cell_plain_lines(self, cell: Tag) -> List[str]:
        """\u5355\u5143\u683c\u6309\u663e\u5f0f\u6362\u884c(\u6bcf\u4e2a <p> \u4e00\u884c\uff0c<br> \u518d\u5206\u884c)\u62c6\u6210\u591a\u884c\u7eaf\u6587\u672c\u3002
        \u4e22\u5f03 Word \u4e66\u7b7e\u4ea7\u751f\u7684 [Table_XXX] \u5360\u4f4d\u6807\u8bb0(\u5982 [Table_Finance]/[Table_EPS])\u3002"""
        paras = cell.find_all("p")
        chunks = paras if paras else [cell]
        lines: List[str] = []
        for ch in chunks:
            for ln in self._plain_with_breaks(ch).split("\n"):
                ln = normalize_text(ln)
                if ln and not re.fullmatch(r"\[Table_\w+\]", ln):
                    lines.append(ln)
        return lines

    def _cell_align(self, cell: Tag) -> Optional[str]:
        """\u4ece HTML \u8bfb\u53d6\u5355\u5143\u683c\u6c34\u5e73\u5bf9\u9f50(text-align / align)\u3002\u8bfb\u4e0d\u5230\u8fd4\u56de None\u3002"""
        for el in cell.find_all("p") + [cell]:
            style = (el.get("style", "") or "").lower()
            m = re.search(r"text-align\s*:\s*(left|right|center)", style)
            if m:
                return {"left": "l", "right": "r", "center": "c"}[m.group(1)]
            a = (el.get("align", "") or "").lower()
            if a in ("left", "right", "center"):
                return {"left": "l", "right": "r", "center": "c"}[a]
        return None

    def _cell_is_bold(self, cell: Tag) -> bool:
        """\u5355\u5143\u683c\u662f\u5426\u6574\u4f53\u52a0\u7c97\uff1a\u6240\u6709\u300c\u6709\u5b57\u300d\u7684\u6587\u672c\u8282\u70b9\u90fd\u5904\u4e8e\u52a0\u7c97\u4e0a\u4e0b\u6587\u4e2d\u3002
        \u53ea\u6709\u4e00\u4e2a\u8bcd\u52a0\u7c97\u4e0d\u7b97(\u907f\u514d\u8bef\u628a\u666e\u901a\u6570\u636e\u5355\u5143\u683c\u6574\u4f53\u52a0\u7c97)\u3002"""
        flags: List[bool] = []

        def walk(node, bold: bool) -> None:
            if isinstance(node, NavigableString):
                if str(node).strip():
                    flags.append(bold)
                return
            if not isinstance(node, Tag):
                return
            b = bold or self._tag_is_bold(node)
            for ch in node.children:
                walk(ch, b)

        walk(cell, False)
        return bool(flags) and all(flags)

    def _resolve_local_image(self, src: str) -> Optional[str]:
        """把单元格图片 src 解析成本地文件路径(供 Pillow 读取)。
        base64 data URI 落盘到 generated_dir；远程图(http/https)不下载，返回 None。"""
        src = (src or "").strip()
        if not src:
            return None
        if src.startswith("data:"):
            try:
                header, b64 = src.split(",", 1)
                ext = ".png"
                low = header.lower()
                if "jpeg" in low or "jpg" in low:
                    ext = ".jpg"
                elif "gif" in low:
                    ext = ".gif"
                if not self.generated_dir:
                    return None
                data = base64.b64decode(b64)
                import hashlib
                key = hashlib.md5(data).hexdigest()
                cache = getattr(self, "_img_data_cache", None)
                if cache is None:
                    cache = self._img_data_cache = {}
                if key in cache:           # 同一 base64 多格引用 -> 复用同一临时文件，不重复落盘
                    return cache[key]
                self._img_tmp_counter = getattr(self, "_img_tmp_counter", 0) + 1
                fp = Path(self.generated_dir) / f"cellimg_{self._img_tmp_counter}{ext}"
                fp.write_bytes(data)
                cache[key] = str(fp)
                return str(fp)
            except Exception:
                return None
        parsed = urlparse(src)
        if parsed.scheme in {"http", "https"}:
            return None
        if parsed.scheme == "file":
            raw_path = unquote(parsed.path)
            if parsed.netloc and parsed.netloc.lower() != "localhost":
                # UNC 地址(file://server/share/...)：netloc 是服务器名，不能丢；
                # file://localhost/C:/... 是本机地址，按无 netloc 处理
                raw_path = "\\\\" + unquote(parsed.netloc) + raw_path.replace("/", "\\")
            elif raw_path.startswith("/") and re.match(r"^/[A-Za-z]:", raw_path):
                raw_path = raw_path[1:]
            abs_path = Path(raw_path)
        else:
            if self.current_html_path is None:
                return None
            cleaned = src.split("#", 1)[0].split("?", 1)[0]
            abs_path = (self.current_html_path.parent / cleaned).resolve()
        if abs_path.exists() and abs_path.is_file():
            return str(abs_path)
        return None

    def _cell_image_files(self, cell: Tag) -> List[str]:
        """单元格内可读取的本地图片路径列表(空列表=无图)。"""
        files: List[str] = []
        for node in cell.find_all(True):
            if (node.name or "").lower() not in {"img", "v:imagedata"}:
                continue
            src = node.get("src", "") or node.get("o:href", "") or node.get("href", "")
            resolved = self._resolve_local_image(src)
            if resolved:
                files.append(resolved)
            elif self._current_block is not None:
                # 表格转图时单元格里的图取不到：整表照样渲染成一张完好的 PNG，该格留空，
                # 少掉的那张图在成品里看不出来，只有这里能记下并定位告警
                self._note_failed(self._current_block, "中有单元格图片无法读取（表图中该格留空）", "image not found")
            else:
                self._note_render_problem(f"表格单元格里的图片无法读取：{(src or '').strip()[:120]!r}")
        return files

    def _cell_own_text(self, cell: Tag) -> str:
        """单元格自身文本，**排除内层嵌套表**的文字(避免 wrapper 把内表文字算进题注/来源判定)。"""
        if cell.find("table") is None:
            return cell.get_text(" ", strip=True)
        import copy
        clone = copy.copy(cell)
        for nt in clone.find_all("table"):
            nt.decompose()
        return clone.get_text(" ", strip=True)

    _META_MAX_SKIP = 2
    # 「明确辅助行」白名单：单位:/单位（百万元）/注:/注1:/备注:/说明:/英文 Note(s):/Unit: 等。
    # 只跳过/合并这类行，不把任意单格行(板块行、小标题)都当辅助行，避免误吞表体内容。
    _RE_META_AUX = re.compile(
        r"^\s*(?:"
        r"单位\s*[:：（(/]"                                    # 单位:/单位（百万元）/单位(..)/单位/..
        r"|(?:备注|備註|注释|註釋|说明|說明|注|註)\s*\d*\s*[:：]"
        r"|(?:notes?|unit)\s*[:：]"                            # 英文 Note:/Notes:/Unit:
        r")", re.I)

    def _is_meta_aux_row(self, text: str) -> bool:
        return bool(self._RE_META_AUX.match(text or ""))

    def _meta_row(self, rows: List[Tag], regex, from_top: bool = True):
        """返回 (题注/数据来源行号, [途中跳过的辅助行号])；找不到返回 (None, [])。
        从顶(或底)起第一个**有自身内容**的行，若该行**第一个有字单元格**就以 regex 开头
        (题注/来源都以「表X：」「数据来源：」起头)即命中。
        - 用「领头单元格匹配」而非「整行只有 1 格」：兼容题注被 Word 拆成「表1：| 标题」多格；
        - 向里扫时只跳过「明确辅助行」(单格且匹配 _is_meta_aux_row，如 单位:/注:)，最多跳 2 行后再找一次；
          遇到「数据行 / 单格小标题 / 板块行」等任何**非辅助**行就停止，绝不把它们误吞进题注/来源；
        - 判定用 _cell_own_text(排除内层表)：嵌套 wrapper 的网格行(内容是内层表)被视为空、跳过。"""
        order = list(range(len(rows))) if from_top else list(range(len(rows) - 1, -1, -1))
        skipped: List[int] = []
        for i in order:
            cells = [c for c in rows[i].find_all(["td", "th"], recursive=False)
                     if normalize_text(self._cell_own_text(c))]
            if not cells:
                continue
            lead = normalize_text(self._cell_own_text(cells[0]))
            if regex.search(lead):
                return i, skipped
            if len(cells) == 1 and self._is_meta_aux_row(lead):   # 明确辅助行：跳过继续找
                skipped.append(i)
                if len(skipped) > self._META_MAX_SKIP:
                    return None, []
                continue
            return None, []                                       # 其余(数据行/小标题/板块行)：停止
        return None, []

    def _adjacent_aux(self, rows: List[Tag], idx: int, step: int) -> List[int]:
        """从 idx 沿 step 方向收集**紧邻的明确辅助行**(单位:/注:)。用于把题注下方的单位行、
        来源上方的注释行也并进题注/来源(命中后再向「表体一侧」额外吸收)。
        遇到空行即停止——空行往往是段/表的分界，跨空行吸收会把下一段/下一张表的注释误并进来。"""
        out: List[int] = []
        j = idx + step
        while 0 <= j < len(rows):
            cells = [c for c in rows[j].find_all(["td", "th"], recursive=False)
                     if normalize_text(self._cell_own_text(c))]
            if not cells:
                break
            if len(cells) == 1 and self._is_meta_aux_row(normalize_text(self._cell_own_text(cells[0]))):
                out.append(j)
                j += step
                continue
            break
        return out

    def _img_display_dim(self, node: Tag, which: str) -> float:
        """图片显示尺寸(px)：先读 width/height 属性，再读 style 的 width/height(px/pt)。读不到返回 0。"""
        v = str(node.get(which, "") or "").strip()
        m = re.match(r"([\d.]+)", v)
        if m:
            try:
                return float(m.group(1))
            except Exception:
                pass
        m = re.search(rf"{which}\s*:\s*([\d.]+)\s*(px|pt)?", str(node.get("style", "") or ""), re.I)
        if m:
            try:
                val = float(m.group(1))
                if (m.group(2) or "").lower() == "pt":
                    val *= 4.0 / 3.0
                return val
            except Exception:
                pass
        return 0.0

    def _img_is_big(self, node: Tag) -> bool:
        """图表本体一般是大图：宽 >=260px 或 高 >=180px。读不到尺寸算非大图。"""
        return self._img_display_dim(node, "width") >= 260 or self._img_display_dim(node, "height") >= 180

    def _has_chart_like_image(self, table: Tag) -> bool:
        """是否存在「纯图片行(该行无文字)里的大图」——图表本体的典型形态。
        数据表里的图片在**带文字的数据行**里(如产品图示列)，不算图表本体，故不会因含图就一票判成图表。"""
        grid = self._select_grid_table(table) or table
        for tr in self._table_direct_rows(grid):
            cells = tr.find_all(["td", "th"], recursive=False)
            if any(normalize_text(self._cell_own_text(c)) for c in cells):
                continue                    # 该行有文字 -> 不是纯图片行
            for c in cells:
                for n in c.find_all(["img", "v:imagedata"]):
                    if self._img_is_big(n):
                        return True
        return False

    def _meta_inline_from_row(self, row: Tag, preserve_color: bool) -> str:
        """把题注/来源行的**所有有字 direct 单元格**拼成内联 HTML(排除内层表)，
        兼容题注被拆成「表1：| 标题文字」多格的情况。"""
        parts: List[str] = []
        for cell in row.find_all(["td", "th"], recursive=False):
            if not normalize_text(self._cell_own_text(cell)):
                continue
            seg = self._table_meta_inline(cell, preserve_color=preserve_color)
            if seg:
                parts.append(seg)
        return self._cleanup_inline_html(" ".join(parts))

    def _select_grid_table(self, table: Tag) -> Optional[Tag]:
        """\u5728 table \u81ea\u8eab\u4e0e\u5176\u5185\u90e8\u5d4c\u5957\u8868\u683c\u4e2d\uff0c\u6311\u51fa\u771f\u6b63\u627f\u8f7d\u6570\u636e\u7f51\u683c\u7684\u90a3\u5f20\u8868
        (\u884c\u6570>=2 \u4e14 \u5355\u884c\u6700\u591a\u5217\u6570>=2 \u4e14\u5355\u5143\u683c\u603b\u6570\u6700\u591a\u8005)\u3002\u8fd4\u56de None=\u6ca1\u6709\u53ef\u6e32\u67d3\u7f51\u683c\u3002

        \u6df1\u5ea6\u62a5\u544a\u91cc\u300c\u53ef\u6bd4\u516c\u53f8\u4f30\u503c\u8868\u300d\u5e38\u88ab Word \u5305\u6210\uff1a\u5916\u5c42\u8868(\u9898\u6ce8\u884c/\u7f51\u683c\u884c/\u6570\u636e\u6765\u6e90\u884c)\uff0c
        \u5176\u4e2d\u7f51\u683c\u884c\u7684\u5355\u5143\u683c\u91cc\u518d\u5d4c\u4e00\u5f20\u771f\u6b63\u7684\u6570\u636e\u8868\uff1b\u6b64\u65f6\u5e94\u4e0b\u94bb\u5230\u5185\u5c42\u8868\u5f53\u7f51\u683c\u3002

        \u53ea\u6709\u5f53\u672c\u8868\u81ea\u8eab**\u4e0d**\u50cf\u6570\u636e\u7f51\u683c(\u5373\u4e0d\u6ee1\u8db3 >=2 \u884c x >=2 \u5217)\u65f6\u624d\u4e0b\u94bb\u5230\u5d4c\u5957\u8868\uff0c
        \u907f\u514d\u300c\u5916\u5c42\u662f\u771f\u8868\u3001\u67d0\u683c\u91cc\u53c8\u5d4c\u4e86\u5c0f\u8868\u300d\u65f6\u8bef\u9009\u5c0f\u8868\u3002
        \u5224\u5b9a\u300c>=2 \u5217\u300d\u65f6\u6309**\u5355\u5143\u683c\u81ea\u8eab\u6587\u5b57**(_cell_own_text\uff0c\u6392\u9664\u5185\u5c42\u8868)\u6570\uff1a\u7a7a\u767d\u5360\u4f4d\u683c\u548c\u300c\u7eaf\u5305\u5185\u5c42\u8868\u3001
        \u81ea\u8eab\u65e0\u5b57\u300d\u7684\u683c\u90fd\u4e0d\u7b97\u6570\u636e\u5217(\u4e8e\u662f wrapper \u884c\u300c\u7a7a\u767d\u683c + \u5305\u5185\u8868\u7684\u683c\u300d\u4e0d\u88ab\u8bef\u5f53\u4e24\u5217)\uff1b\u4f46\u300c\u81ea\u8eab\u6709\u5b57\u3001
        \u540c\u65f6\u53c8\u5d4c\u4e86\u5c0f\u8868\u300d\u7684\u683c\u4ecd\u7b97\u4e00\u5217(\u4e0d\u4f1a\u628a\u771f\u5916\u5c42\u8868\u6f0f\u5224)\u3002"""
        def qualifies(t: Tag) -> bool:
            rows = self._table_direct_rows(t)
            if len(rows) < 2:
                return False
            return max((sum(1 for c in tr.find_all(["td", "th"], recursive=False)
                            if normalize_text(self._cell_own_text(c)))
                        for tr in rows), default=0) >= 2

        if qualifies(table):
            return table

        best = None
        best_cells = -1
        for cand in [x for x in table.find_all("table") if x is not table]:
            if not qualifies(cand):
                continue
            ncells = sum(len(tr.find_all(["td", "th"], recursive=False))
                         for tr in self._table_direct_rows(cand))
            if ncells > best_cells:
                best_cells = ncells
                best = cand
        return best

    def _table_meta_inline(self, cell: Tag, preserve_color: bool) -> str:
        """\u6e32\u67d3\u9898\u6ce8/\u6570\u636e\u6765\u6e90\u5355\u5143\u683c\u7684\u5185\u8054 HTML\uff0c\u5254\u9664\u5176\u4e2d\u53ef\u80fd\u5d4c\u5957\u7684\u8868\u683c(\u907f\u514d\u628a\u6574\u8868\u6587\u5b57\u6df7\u8fdb\u9898\u6ce8)\u3002"""
        import copy
        clone = copy.copy(cell)
        for nt in clone.find_all("table"):
            nt.decompose()
        return self._cleanup_inline_html(
            self._render_inline_children(clone, preserve_color=preserve_color))

    def _build_table_spec(self, table: Tag) -> Optional[dict]:
        """\u628a HTML \u8868\u683c\u89e3\u6790\u6210 table_image \u7684\u7f51\u683c spec\uff0c\u5e76\u62bd\u51fa\u9898\u6ce8/\u6570\u636e\u6765\u6e90(\u5df2\u6e32\u67d3\u6210\u5185\u8054 HTML)\u3002

        \u8fd4\u56de {"spec": dict|None, "caption_html": str, "source_html": str}\u3002
        spec \u4e3a None \u8868\u793a\u6ca1\u6709\u53ef\u6e32\u67d3\u7684\u8868\u4f53\u3002\u6574\u4f53\u8fd4\u56de None \u8868\u793a\u6839\u672c\u4e0d\u662f\u53ef\u6e32\u67d3\u7684\u8868\u3002
        """
        grid_table = self._select_grid_table(table)
        if grid_table is None:
            return None
        nested_mode = grid_table is not table
        grid_rows = self._table_direct_rows(grid_table)

        # \u9898\u6ce8/\u6570\u636e\u6765\u6e90\uff1a\u5d4c\u5957\u65f6\u53d6\u81ea\u5916\u5c42\u8868(\u7f51\u683c\u4e4b\u5916)\uff1b\u975e\u5d4c\u5957\u65f6\u53d6\u81ea\u7f51\u683c\u81ea\u8eab\u7684\u9898\u6ce8\u884c/\u6765\u6e90\u884c\u3002
        # \u4e24\u6761\u8def\u5f84\u90fd\u7528\u6536\u7d27\u540e\u7684 _meta_row(\u53ea\u8ba4\u6700\u9760\u9876/\u5e95\u3001\u9886\u5934\u5355\u5143\u683c\u5339\u914d\u7684\u884c\uff1b\u6392\u9664\u5185\u5c42\u8868\u6587\u5b57)\u3002
        exclude_rows: set = set()
        meta_rows = self._table_direct_rows(table) if nested_mode else grid_rows
        cap_i, cap_skip = self._meta_row(meta_rows, self.re_caption, from_top=True)
        src_i, src_skip = self._meta_row(meta_rows, self.re_source, from_top=False)
        # 命中的题注/来源行 + 沿途跳过的辅助行(单位:/注:) + 命中后向表体一侧紧邻的辅助行 一起：
        # 按阅读顺序合并文字、从网格排除。于是「表X 下面的 单位:」「数据来源 上/下面的 注:」都会
        # 随题注/来源走到图上/图下，而不是留在表图里。
        cap_idx = (sorted(set([cap_i] + cap_skip + self._adjacent_aux(meta_rows, cap_i, +1)))
                   if cap_i is not None else [])
        if src_i is not None:
            src_set = set([src_i] + src_skip + self._adjacent_aux(meta_rows, src_i, -1))
            src_set -= set(cap_idx)          # 别和题注侧重复占用同一行(极小表里可能交叠)
            src_idx = sorted(src_set)
        else:
            src_idx = []
        if not nested_mode:
            exclude_rows.update(cap_idx)
            exclude_rows.update(src_idx)

        def join_rows_inline(idxs, preserve_color):
            parts = []
            for j in idxs:
                seg = self._meta_inline_from_row(meta_rows[j], preserve_color=preserve_color)
                if seg:
                    parts.append(seg)
            return self._cleanup_inline_html("<br/>".join(parts))

        caption_html = ""
        source_html = ""
        if cap_idx:
            caption_html = self._strip_leading_inline_breaks(
                join_rows_inline(cap_idx, preserve_color=True))
        if src_idx:
            source_html = join_rows_inline(src_idx, preserve_color=False)

        spec = self._grid_rows_to_spec(grid_rows, exclude_rows)
        return {"spec": spec, "caption_html": caption_html, "source_html": source_html}

    def _grid_rows_to_spec(self, trs: List[Tag], exclude_rows: set) -> Optional[dict]:
        """\u628a\u7f51\u683c\u884c(\u5df2\u53bb\u6389\u9898\u6ce8/\u6765\u6e90\u884c)\u89e3\u6790\u6210\u7d27\u51d1 spec\uff1a\u5904\u7406\u6a2a\u5411/\u7eb5\u5411\u5408\u5e76\u3001\u4e22\u7a7a\u5217\u3002"""
        if not trs:
            return None

        # 1) \u653e\u7f6e\u6bcf\u4e2a\u5355\u5143\u683c\u5230\u7f51\u683c(\u5904\u7406 colspan \u6a2a\u5411 + rowspan \u7eb5\u5411\u5408\u5e76)
        placed: List[dict] = []
        occupied = set()
        ncols = 0
        for r, tr in enumerate(trs):
            c = 0
            for cell in tr.find_all(["td", "th"], recursive=False):
                while (r, c) in occupied:
                    c += 1
                try:
                    cs = max(1, int((cell.get("colspan", "1") or "1").strip()))
                except Exception:
                    cs = 1
                try:
                    rs = max(1, int((cell.get("rowspan", "1") or "1").strip()))
                except Exception:
                    rs = 1
                placed.append({"r": r, "c": c, "rs": rs, "cs": cs, "cell": cell,
                               "lines": self._cell_plain_lines(cell),
                               "images": self._cell_image_files(cell)})
                for dr in range(rs):
                    for dc in range(cs):
                        occupied.add((r + dr, c + dc))
                c += cs
                ncols = max(ncols, c)
        nrows = len(trs)
        if ncols == 0:
            return None

        # 2) \u8868\u4f53\u884c = \u975e\u9898\u6ce8/\u6765\u6e90\u884c \u4e14 \u8be5\u884c(\u542b\u88ab\u7eb5\u5411\u5408\u5e76\u8986\u76d6)\u6709\u5185\u5bb9(\u6587\u5b57\u6216\u56fe\u7247)
        def has_content(p: dict) -> bool:
            return bool(p["lines"] or p["images"])

        def row_has_content(r: int) -> bool:
            return any(has_content(p) and p["r"] <= r < p["r"] + p["rs"] for p in placed)

        kept_rows = [r for r in range(nrows) if r not in exclude_rows and row_has_content(r)]
        if not kept_rows:
            return None
        rowpos = {old: i for i, old in enumerate(kept_rows)}
        kept_set_rows = set(kept_rows)

        # 3) \u4e22\u5f03\u7a7a\u5217\u3002col_has=\u4efb\u610f\u884c\u6709\u300c\u5355\u5217(cs==1)\u4e14\u6709\u6587\u5b57\u300d\u7684\u5185\u5bb9\uff1b
        #    col_has_body=\u8868\u5934\u884c(\u9996\u4e2a\u4fdd\u7559\u884c)\u4e4b\u5916\u8fd8\u6709\u5355\u5217\u5185\u5bb9\u3002
        #    \u884c\u6570>=3 \u65f6\uff0c\u4e22\u6389\u300c\u53ea\u6709\u8868\u5934\u3001\u6570\u636e\u5168\u7a7a\u300d\u7684\u5217(\u5e38\u89c1\u4e8e\u88ab\u5265\u6389\u56fe\u7247\u7684\u56fe\u793a\u5217)\uff1b
        #    \u884c\u6570\u5f88\u5c11\u65f6\u4fdd\u5b88\u8d77\u89c1\u53ea\u4e22\u5168\u7a7a\u5217\u3002Word \u63d2\u5165\u7684\u6781\u7a84\u5360\u4f4d\u7a7a\u5217\u4e5f\u5728\u6b64\u4e00\u5e76\u4e22\u6389\u3002
        # \u53ea\u4e22\u300c\u6574\u5217\u65e0\u4efb\u4f55\u5185\u5bb9\u300d\u7684\u7a7a\u5217(\u5e38\u89c1\u4e8e Word \u5728\u8868\u95f4\u63d2\u5165\u7684\u6781\u7a84\u5360\u4f4d\u5217)\u3002
        # \u6709\u4efb\u4f55\u5185\u5bb9(\u6587\u5b57\u6216\u56fe\u7247)\u7684\u5217\u4e00\u5f8b\u4fdd\u7559\u2014\u2014\u4e0d\u518d\u6309\u300c\u53ea\u6709\u8868\u5934\u6709\u5185\u5bb9\u300d\u4e22\u5217\uff0c
        # \u90a3\u4f1a\u8bef\u5220\u65e0\u8868\u5934\u8868\u91cc\u300c\u53ea\u5728\u9996\u884c\u6709\u6570\u636e\u300d\u7684\u5408\u6cd5\u5217(\u56fe\u7247\u5217\u5df2\u8ba1\u4e3a\u5185\u5bb9\uff0c\u65e0\u9700\u518d\u9760\u4e22\u5217\u5904\u7406)\u3002
        col_has = [False] * ncols
        for p in placed:
            if p["r"] in kept_set_rows and p["cs"] == 1 and has_content(p):
                col_has[p["c"]] = True
        kept_set_cols = {c for c in range(ncols) if col_has[c]} or set(range(ncols))
        # \u4fdd\u62a4\uff1a\u8de8\u5217\u4e14\u6709\u5185\u5bb9\u7684\u5355\u5143\u683c\u81f3\u5c11\u4fdd\u7559\u5176\u8de8\u5ea6\u5185\u7684\u4e00\u5217
        for p in placed:
            if p["r"] in kept_set_rows and p["cs"] > 1 and has_content(p):
                if not any(c in kept_set_cols for c in range(p["c"], p["c"] + p["cs"])):
                    kept_set_cols.add(p["c"])
        kept_cols = sorted(kept_set_cols)
        if not kept_cols:
            return None
        colpos = {old: i for i, old in enumerate(kept_cols)}

        # 4) \u91cd\u6620\u5c04\u5230\u7d27\u51d1\u7f51\u683c
        spec_cells: List[dict] = []
        for p in placed:
            if p["r"] not in rowpos:
                continue
            rspan = [r for r in range(p["r"], p["r"] + p["rs"]) if r in rowpos]
            cspan = [c for c in range(p["c"], p["c"] + p["cs"]) if c in colpos]
            if not rspan or not cspan:
                continue
            spec_cells.append({
                "r": rowpos[rspan[0]], "c": colpos[cspan[0]],
                "rs": len(rspan), "cs": len(cspan),
                "lines": p["lines"], "images": p["images"],
                "align": self._cell_align(p["cell"]),
                "bold": self._cell_is_bold(p["cell"]),
            })
            # 只有真正进了 spec 的单元格才算「这张图被用掉了」。题注/来源行在上面被剪掉，
            # 它们里面的图不会进表图——若在读取时就记成已访问，这类丢图就查不出来了
            for src in self._image_srcs(p["cell"]):
                self._note_visited_image(src)
        nb = len(kept_rows)
        nc = len(kept_cols)
        if nb < 1 or nc < 2 or not spec_cells:
            return None

        # 5) \u8868\u5934\u884c\u6570\uff1a\u9ed8\u8ba4\u9996\u884c\uff1b\u82e5\u9996\u884c\u6709\u6a2a\u5411\u5408\u5e76(\u5206\u7ec4\u8868\u5934)\u4e14\u603b\u884c\u6570>=3\uff0c\u5219\u524d\u4e24\u884c\u90fd\u7b97\u8868\u5934
        header_rows = 1
        if nb >= 3 and any(cell["r"] == 0 and cell["cs"] > 1 for cell in spec_cells):
            header_rows = 2

        return {"ncols": nc, "nrows": nb, "header_rows": header_rows, "cells": spec_cells}

    def _strip_leading_inline_breaks(self, inline_html: str) -> str:
        """\u53bb\u6389\u9898\u6ce8\u5f00\u5934\u7a7a\u884c(<br/>\uff0c\u53ef\u80fd\u88ab\u7a7a span \u5305\u88f9)\u9020\u6210\u7684\u591a\u4f59\u7a7a\u767d\u3002"""
        while True:
            stripped = re.sub(r"^(?:<span[^>]*>)*\s*<br/>\s*(?:</span>)*\s*", "", inline_html).strip()
            if stripped == inline_html:
                return inline_html
            inline_html = stripped

    def _ensure_table_png(self, block: Block, spec: Optional[dict]) -> Optional[str]:
        """\u6e32\u67d3\u8be5\u8868\u4e3a PNG(\u6bcf\u5757\u53ea\u6e32\u67d3\u4e00\u6b21\u5e76\u7f13\u5b58)\u3002\u8fd4\u56de\u8def\u5f84\u6216 None\u3002"""
        if block.table_png_path is not None:
            return block.table_png_path or None
        if spec is None or not self.generated_dir:
            block.table_png_path = ""
            return None
        try:
            import sys
            if str(SCRIPT_DIR) not in sys.path:
                sys.path.insert(0, str(SCRIPT_DIR))
            import table_image as ti

            out_png = Path(self.generated_dir) / f"data_table_{block.idx}.png"
            missing: List[str] = []
            ti.render_table(spec, str(out_png), missing=missing)
            block.table_png_path = str(out_png)
            print(f"  -> data table image: {spec['nrows']} rows x {spec['ncols']} cols (block {block.idx})")
            if missing:
                self._note_failed(block, f"中有 {len(missing)} 张单元格图片未能读取（表图中该格留空）", "image unreadable")
        except Exception as e:
            print(f"  -> data table skipped (block {block.idx}): {e}")
            self._note_render_problem(f"表图渲染失败（block {block.idx}）：{e}")
            block.table_png_path = ""
        return block.table_png_path or None

    def _render_data_table_block(self, block: Block, for_docx: bool = False) -> str:
        """\u6570\u636e\u8868\u683c\u4e0e\u56fe1/\u56fe2 \u540c\u6784\uff1a\u9898\u6ce8\u5728\u4e0a\u3001\u8868\u4f53\u56fe\u5728\u4e2d\u3001\u6570\u636e\u6765\u6e90/\u6ce8\u5728\u4e0b\u3002
        \u590d\u7528\u56fe\u8868\u7684\u9898\u6ce8/\u6765\u6e90/\u56fe\u7247\u6e32\u67d3\uff0c\u53ea\u628a\u300c\u672c\u4f53\u300d\u6362\u6210\u751f\u6210\u7684\u8868\u683c\u56fe\u7247\u3002"""
        if not self.out_cfg.get("generate_images", True):
            return ""
        built = block.table_build or {}
        spec = built.get("spec")

        parts: List[str] = []
        # \u9898\u6ce8(\u8868X\uff1a)\uff1a\u590d\u7528\u56fe\u8868\u6807\u9898\u5904\u7406\u3002\u7f16\u53f7\u7531 Word \u89e3\u6790\u9898\u6ce8\u57df\u5f97\u5230\uff0c\u81ea\u52a8\u9002\u914d\u88681/\u88682\u2026
        cap_html = built.get("caption_html") or ""
        if cap_html:
            parts.append(self._render_figure_caption_html(cap_html))
        # \u8868\u4f53\uff1a\u751f\u6210\u7684\u8868\u683c\u56fe\u7247\uff1b\u6e32\u67d3\u5931\u8d25\u5219\u515c\u5e95\u6210\u6587\u5b57\u6bb5\u843d\uff0c\u907f\u514d\u300c\u770b\u4f3c\u6210\u529f\u4f46\u8868\u6ca1\u4e86\u300d
        png = self._ensure_table_png(block, spec)
        if png:
            try:
                resolved = self._register_asset(Path(png).as_uri(), inline=not for_docx)
            except PermissionError as exc:
                # 表图一直被其他程序独占：纯文字表改走文字兜底（内容都在 spec 里）；含图片的表没有完整退路，照常失败
                if spec is None or any(cell.get("images") for cell in spec.get("cells", [])):
                    raise
                print(f"  [warn] table image unreadable, kept as text: {exc}")
                resolved = None
                fallback = self._render_spec_as_text(spec)
                if fallback:
                    parts.append(fallback)
            if resolved:
                image_margin = self.out_cfg.get("image_margin_css", "0 auto")
                image_border = self.out_cfg.get("image_border_css", "0")
                img_html = (
                    f'<img src="{html.escape(resolved)}" alt="" '
                    f'style="display:block; max-width:100%; height:auto; margin:{image_margin}; border:{image_border};" />'
                )
                parts.append(self._render_figure_images_html([img_html]))
        elif spec is not None:
            # 含图片的表没有完整的文字退路：交给逐块失败处理（只留空行、定位告警），不写「[图片]」占位
            if any(cell.get("images") for cell in spec.get("cells", [])):
                raise ProcessingError("含图片的表格未能转成图片")
            fallback = self._render_spec_as_text(spec)
            if fallback:
                parts.append(fallback)
        # \u6570\u636e\u6765\u6e90/\u6ce8\uff1a\u590d\u7528\u56fe\u8868\u6765\u6e90\u5904\u7406(\u6977\u4f53\u3001\u4fdd\u7559 <br/> \u591a\u884c)
        src_html = built.get("source_html") or ""
        if src_html:
            parts.append(self._render_figure_source_html(src_html))
        return "\n".join(p for p in parts if p)

    def _render_spec_as_text(self, spec: Optional[dict]) -> str:
        """表图渲染失败时的兜底：把表体按行拼成「a | b | c」文字段落，至少不丢内容。"""
        if not spec:
            return ""
        byrow: Dict[int, List[dict]] = {}
        for c in spec.get("cells", []):
            byrow.setdefault(c["r"], []).append(c)
        paras: List[str] = []
        for r in sorted(byrow):
            cells = sorted(byrow[r], key=lambda x: x["c"])
            pieces: List[str] = []
            for c in cells:
                lines = c.get("lines") or []
                if lines:
                    pieces.append(" ".join(lines))
            txt = " | ".join(pieces)
            if txt.strip(" |"):
                paras.append(self._render_body_paragraph_html(html.escape(txt)))
        return "\n".join(paras)

    def _render_leftover_table_text(self, table: Tag) -> str:
        """已不是 data_table/figure_table 的遗留表，若它实际上只有**一个有字单元格**
        (典型如单格标题表，例如「豪悦护理三大财务预测表」)，则兜底成普通段落，避免标题被静默丢弃。
        多格表保持旧行为(返回空串，交由调用方处理)。"""
        import copy
        content_cells = [c for tr in self._table_direct_rows(table)
                         for c in tr.find_all(["td", "th"], recursive=False)
                         if normalize_text(c.get_text(" ", strip=True))]
        if len(content_cells) != 1:
            return ""
        clone = copy.copy(content_cells[0])
        for nt in clone.find_all("table"):
            nt.decompose()
        inner = self._cleanup_inline_html(self._render_inline_children(clone))
        if not inner:
            return ""
        plain = normalize_text(BeautifulSoup(inner, "html.parser").get_text(" ", strip=True))
        if not plain or re.fullmatch(r"\[Table_\w+\]", plain):
            return ""
        font_family = self.out_cfg["font_family"]
        font_size = self.out_cfg["font_size_px"]
        line_height = self.out_cfg["line_height"]
        margin_bottom = self.out_cfg["paragraph_margin_bottom_em"]
        margin_top = self.out_cfg.get("paragraph_margin_top_css", "0")
        margin_right = self.out_cfg.get("paragraph_margin_right_css", "0")
        margin_left = self.out_cfg.get("paragraph_margin_left_css", "0")
        text_indent = self.out_cfg.get("table_fallback_text_indent_css", "0")
        text_align = self.out_cfg.get("paragraph_text_align", "left")
        return (
            f'<p style="margin:{margin_top} {margin_right} {margin_bottom}em {margin_left}; '
            f'text-indent:{text_indent}; text-align:{text_align}; '
            f'font-family:{font_family}; font-size:{font_size}px; line-height:{line_height};">{inner}</p>'
        )

    def _is_valuation_table(self, table: Tag) -> bool:
        keyword = self.doc_cfg.get("valuation_table_keyword", "估值表")
        rows = self._table_rows_expanded(table)
        for row in rows[:3]:
            for cell in row:
                if keyword in self._cell_text(cell):
                    return True
        return False

    def _valuation_html_signature(self, table: Tag, signature_builder) -> tuple:
        rows = []
        for tr in self._table_direct_rows(table):
            rows.append([
                self._cell_own_text(cell)
                for cell in tr.find_all(["td", "th"], recursive=False)
            ])
        return signature_builder(rows)

    def _ensure_valuation_png(self, block: Block) -> Optional[dict]:
        """仅在源 docx 中唯一命中估值表时渲染图片。返回 {png,title,notes} 或 None。"""
        if self._valuation_cache is not None:
            return self._valuation_cache or None
        self._valuation_cache = {}
        if not self.source_docx_path or not self.generated_dir:
            return None
        try:
            import sys
            if str(SCRIPT_DIR) not in sys.path:
                sys.path.insert(0, str(SCRIPT_DIR))
            import valuation_table_image as vti

            keyword = self.doc_cfg.get("valuation_table_keyword", "估值表")
            doc = vti.docx.Document(str(self.source_docx_path))
            signature = self._valuation_html_signature(block.tag, vti.content_signature)
            table = vti.find_valuation_table(
                doc, keyword=keyword, require_keyword=True, signature=signature)
            if table is None:
                return None
            data = vti.parse_table(table, notes_re=self.re_source)
            out_png = Path(self.generated_dir) / "valuation_table.png"
            vti.render(data, str(out_png))  # 默认不画标题/备注，全表居中
            self._valuation_cache = {
                "png": str(out_png),
                "title": data.get("title", ""),
                "notes": data.get("notes", []),
            }
            n = sum(len(rows) for _, rows in data.get("sections", []))
            print(f"  -> valuation table image: {len(data.get('sections', []))} sectors / {n} names")
        except Exception as e:
            print(f"  -> valuation table skipped: {e}")
            self._valuation_cache = {}
        return self._valuation_cache or None

    def _render_valuation_block(self, block: Block, for_docx: bool = False) -> str:
        """估值表与图1/图2 同构：首行题注(表X：)、中间本体、末行数据来源/注。
        复用图表的题注/来源/图片渲染，只把"本体"换成生成的表格图片。"""
        if not self.out_cfg.get("generate_images", True):
            return ""
        info = self._ensure_valuation_png(block)
        if not info:
            # 估值表无法在源 docx 命中关键词/解析/渲染时，退回通用数据表渲染，
            # 避免整张表被静默丢弃(此前返回 "" 会直接丢表)。
            # 标记 valuation_table 时(见 _mark_media_related_blocks)未建通用 spec，
            # block.table_build 仍是 None，这里惰性补建，否则 _render_data_table_block 无 spec 可画、仍返回空。
            if block.table_build is None:
                block.table_build = self._build_table_spec(block.tag)
            return self._render_data_table_block(block, for_docx=for_docx)
        resolved = self._register_asset(Path(info["png"]).as_uri(), inline=not for_docx)
        if not resolved:
            return ""

        rows = self._table_rows_expanded(block.tag)
        keyword = self.doc_cfg.get("valuation_table_keyword", "估值表")
        # 题注行：自上而下找第一个含"表X："或关键词的单元格(扫全行所有列，
        # 容忍 Word 插入的前导空白行/题注被挪到非首列)
        cap_row = None
        for row in rows:
            for cell in row:
                t = self._cell_text(cell)
                if self.re_caption.search(t) or keyword in t:
                    cap_row = cell.find_parent("tr")
                    break
            if cap_row is not None:
                break
        # 数据来源/注 行：按源顺序收集所有"数据来源…"行(同样扫全行所有列)
        src_rows = []
        for row in rows:
            for cell in row:
                if self.re_source.search(self._cell_text(cell)):
                    src_rows.append(cell.find_parent("tr"))
                    break

        parts = []
        # 标题(题注)：复用图表标题处理。编号"表X："由 Word 解析题注域得到，自动适配表1/表2…
        if cap_row is not None:
            cap_html = self._meta_inline_from_row(cap_row, preserve_color=True)
            cap_html = self._strip_leading_inline_breaks(cap_html)
            if cap_html:
                parts.append(self._render_figure_caption_html(cap_html))
        # 本体：生成的估值表图片
        image_margin = self.out_cfg.get("image_margin_css", "0 auto")
        image_border = self.out_cfg.get("image_border_css", "0")
        img_html = (
            f'<img src="{html.escape(resolved)}" alt="" '
            f'style="display:block; max-width:100%; height:auto; margin:{image_margin}; border:{image_border};" />'
        )
        parts.append(self._render_figure_images_html([img_html]))
        # 数据来源/注：复用图表来源处理(正文字体灰色、保留 <br/> 多行)
        src_html = self._cleanup_inline_html("<br/>".join(
            self._meta_inline_from_row(row, preserve_color=False)
            for row in src_rows
        ))
        if src_html:
            parts.append(self._render_figure_source_html(src_html))
        return "\n".join(p for p in parts if p)

    def _looks_dense_grid(self, table: Tag) -> bool:
        """是否像一张「密集数据网格」：有 >=2 个**非题注/来源**行、每行至少 2 个有字 direct 单元格。
        用于在「既匹配图表 wrapper、又像数据表」时优先按数据表处理，避免含图数据表被误拆成图表。
        真正的图表 wrapper(题注行/图片行/来源行)里：题注/来源行被跳过、图片行没有文字，故 cnt=0。

        看的是 _select_grid_table 选出的**网格表**(而非外层 wrapper)：嵌套的「含图数据表」其密集网格
        在内层，只看外层 direct 行会把内层当成空、漏判成图表。

        每行按「内容格(有字或有图)」计数，但要求该行**至少有 1 个有字格**才算数据行：
        - 这样「名称 | 图片」这种两列图片数据表(数据行=文字+图片)能被算成数据表(修掉被误拆图表)；
        - 而图表 wrapper 的图片行是**纯图片无文字**，不计入，纯图表仍走 figure_table。"""
        grid = self._select_grid_table(table) or table
        cnt = 0
        for tr in self._table_direct_rows(grid):
            rowtext = normalize_text(tr.get_text(" ", strip=True))
            if self.re_caption.search(rowtext) or self.re_source.search(rowtext):
                continue
            text_cells = 0
            content_cells = 0
            for c in tr.find_all(["td", "th"], recursive=False):
                has_text = bool(normalize_text(self._cell_own_text(c)))
                if has_text:
                    text_cells += 1
                if has_text or self._cell_has_image(c):
                    content_cells += 1
            if content_cells >= 2 and text_cells >= 1:
                cnt += 1
        return cnt >= 2

    def _mark_media_related_blocks(self, blocks: List[Block]) -> None:
        remove_images = self.doc_cfg.get("remove_images", True) or (not self.out_cfg.get("generate_images", True))
        remove_tables = self.doc_cfg.get("remove_tables", True)
        valuation_taken = False

        for block in blocks:
            if block.drop:
                continue
            # 这一趟就会解析表格规格并读取单元格里的图片，丢图记录必须落在正确的源块上：
            # 少了这一行，预处理期的记录都挂到 idx=-1，按序号取区块内记录时会被漏掉
            self._current_block_idx = block.idx
            self._current_block = block

            if block.tag.name == "table":
                # 估值表(行业公司估值表)：用专门的 11 列渲染器(valuation_table_image)转图，最优先。
                # 必须早于通用 data_table 检测——否则它也会匹配密集网格、走更朴素的通用渲染。
                if (self.doc_cfg.get("render_valuation_table", False)
                        and not remove_images
                        and self.source_docx_path is not None
                        and not valuation_taken
                        and self._is_valuation_table(block.tag)):
                    # 同一输出只允许一张表进入专用渲染；源文档无法唯一匹配时，
                    # _ensure_valuation_png 会失败并让该表回退到通用渲染。
                    valuation_taken = True
                    block.role = "valuation_table"
                    continue

                # 图表 wrapper 优先；但若该表同时像密集数据网格(含图数据表)，则交给 data_table 处理，
                # 避免「题注行 colspan 盖到图片列」时被误当图表 wrapper、丢掉其它数据列/行。
                # 例外：表里有「纯图片行里的大图」(图表本体的典型形态)时仍按图表处理，避免把大图缩成表格缩略图；
                # 但「文字+大图」的数据表(大图在带文字的数据行里)不受此例外影响，仍走 data_table。
                try:
                    figure_segments = self._extract_figure_segments_from_table(block.tag)
                except Exception as exc:
                    # 预处理期就读图（单元格图片被独占等）：只把这一块记为失败，渲染时留空行并定位告警
                    block.render_failure = exc
                    continue
                if figure_segments and not remove_images:
                    if self._has_chart_like_image(block.tag) or not self._looks_dense_grid(block.tag):
                        block.role = "figure_table"
                        block.figure_segments = figure_segments
                        continue

                # 正文里的真实数据表格(政策表/盈利预测表/可比公司估值表…)：整表转图片。
                # 微信编辑器对复杂表格(尤其带合并单元格的)还原很差，转成 PNG 最稳。
                # 仅当确实是一张「>=2 行 x >=2 列」的表才转，避免把布局用的单格表也转图。
                if (self.doc_cfg.get("render_tables", True)
                        and self.out_cfg.get("generate_images", True)):
                    try:
                        built = self._build_table_spec(block.tag)
                    except Exception as exc:
                        block.render_failure = exc
                        continue
                    if built is not None and built.get("spec") is not None:
                        block.role = "data_table"
                        block.table_build = built
                        continue

                if remove_tables and self._table_first_row_has_table_caption(block.tag):
                    block.drop = True
                continue

            if block.has_image and not block.text and not remove_images:
                block.role = "figure_image"

    def _render_figure_caption_html(self, caption_html: str) -> str:
        caption_html = caption_html.strip()
        if not caption_html:
            return ""
        font_size = self.out_cfg.get("figure_caption_font_size_px", max(12, self.out_cfg["font_size_px"] - 1))
        margin_top = self.out_cfg.get("figure_caption_margin_top_em", 0.55)
        margin_bottom = self.out_cfg.get("figure_caption_margin_bottom_em", 0.22)
        font_family = self.out_cfg.get("font_family", WECHAT_FONT_FAMILY)
        default_color = self.out_cfg.get("figure_caption_text_color", "#17365d")
        margin_right = self.out_cfg.get("figure_caption_margin_right_css", "0")
        margin_left = self.out_cfg.get("figure_caption_margin_left_css", "0")
        text_indent = self.out_cfg.get("figure_caption_text_indent_css", "0")
        text_align = self.out_cfg.get("figure_caption_text_align", "left")
        line_height = self.out_cfg.get("figure_caption_line_height", 1.6)
        return (
            f'<p style="margin:{margin_top}em {margin_right} {margin_bottom}em {margin_left}; '
            f'text-indent:{text_indent}; text-align:{text_align}; '
            f'font-family:{font_family}; font-size:{font_size}px; line-height:{line_height}; color:{default_color};">'
            f'<strong>{caption_html}</strong></p>'
        )

    def _render_figure_source_html(self, source_html: str) -> str:
        source_html = source_html.strip()
        if not source_html:
            return ""
        font_size = self.out_cfg.get("figure_source_font_size_px", max(12, self.out_cfg["font_size_px"] - 1))
        margin_top = self.out_cfg.get("figure_source_margin_top_em", 0.18)
        margin_bottom = self.out_cfg.get("figure_source_margin_bottom_em", 1.2)
        font_family = self.out_cfg.get("font_family", WECHAT_FONT_FAMILY)
        color = self.out_cfg.get("figure_source_text_color", "#888888")
        margin_right = self.out_cfg.get("figure_source_margin_right_css", "0")
        margin_left = self.out_cfg.get("figure_source_margin_left_css", "0")
        text_indent = self.out_cfg.get("figure_source_text_indent_css", "0")
        text_align = self.out_cfg.get("figure_source_text_align", "left")
        line_height = self.out_cfg.get("figure_source_line_height", 1.6)
        return (
            f'<p style="margin:{margin_top}em {margin_right} {margin_bottom}em {margin_left}; '
            f'text-indent:{text_indent}; text-align:{text_align}; color:{color}; '
            f'font-family:{font_family}; font-size:{font_size}px; line-height:{line_height};">{source_html}</p>'
        )

    def _render_figure_images_html(self, images: List[str]) -> str:
        parts = []
        margin = self.out_cfg.get("figure_image_container_margin_css", "0")
        text_indent = self.out_cfg.get("figure_image_container_text_indent_css", "0")
        text_align = self.out_cfg.get("figure_image_container_text_align", "center")
        line_height = self.out_cfg.get("figure_image_container_line_height", 1)
        for image_html in images:
            parts.append(
                f'<p style="margin:{margin}; text-indent:{text_indent}; '
                f'text-align:{text_align}; line-height:{line_height};">' +
                image_html +
                '</p>'
            )
        return "\n".join(parts)

    # ----------------------- 前置内容(摘要 + 内容目录图) -----------------------
    def _build_front_matter(self, summary_html: str, toc_png_path: Optional[str], for_docx: bool = False) -> List[str]:
        """返回拼在正文前面的 HTML 片段列表(摘要在上、目录图在下)。两者都缺时返回 []。"""
        parts: List[str] = []
        summary_html = (summary_html or "").strip()
        if summary_html:
            parts.append(self._render_summary_section(summary_html))
        if toc_png_path:
            toc_html = self._render_toc_image(toc_png_path, for_docx=for_docx)
            if toc_html:
                parts.append(toc_html)
        return parts

    def _render_summary_section(self, summary_html: str) -> str:
        """摘要区: 直接复用「报告摘要转公众号」生成的小标题 + 项目符号 HTML, 底部留白与目录/正文分隔。"""
        margin = self.out_cfg.get("summary_section_margin_css", "0 0 0.6em 0")
        return (
            f'<section style="margin:{margin};">'
            f'{summary_html}'
            '</section>'
        )

    def _render_toc_image(self, toc_png_path: str, for_docx: bool = False) -> str:
        """内容目录图: 居中插入(复用图表图片的居中容器)。注册失败/文件缺失时返回 ""。"""
        try:
            p = Path(toc_png_path)
            if not p.exists():
                return ""
            resolved = self._register_asset(p.as_uri(), inline=not for_docx)
            if not resolved:
                return ""
            image_margin = self.out_cfg.get("image_margin_css", "0 auto")
            image_border = self.out_cfg.get("image_border_css", "0")
            img_html = (
                f'<img src="{html.escape(resolved)}" alt="" '
                f'style="display:block; max-width:100%; height:auto; margin:{image_margin}; border:{image_border};" />'
            )
            margin = self.out_cfg.get("toc_image_margin_css", "0.6em 0 1.0em 0")
            text_indent = self.out_cfg.get("toc_image_text_indent_css", "0")
            text_align = self.out_cfg.get("toc_image_text_align", "center")
            line_height = self.out_cfg.get("toc_image_line_height", 1)
            return (
                f'<p style="margin:{margin}; text-indent:{text_indent}; '
                f'text-align:{text_align}; line-height:{line_height};">'
                + img_html +
                '</p>'
            )
        except Exception:
            return ""

    def _render_figure_segment(self, segment: dict, for_docx: bool = False) -> str:
        if not self.out_cfg.get("generate_images", True):
            return ""
        parts = []
        caption_html = (segment.get("caption_html") or "").strip()
        if caption_html:
            parts.append(self._render_figure_caption_html(caption_html))
        images = segment.get("images_docx") if for_docx else segment.get("images_html")
        images = images or []
        if images:
            parts.append(self._render_figure_images_html(images))
        source_html = (segment.get("source_html") or "").strip()
        if source_html:
            parts.append(self._render_figure_source_html(source_html))
        return "\n".join(parts)

    def _render_block(self, block: Block, for_docx: bool = False) -> str:
        tag = block.tag

        if block.role == "valuation_table":
            return self._render_valuation_block(block, for_docx=for_docx)

        if block.role == "data_table":
            return self._render_data_table_block(block, for_docx=for_docx)

        if block.role == "figure_table":
            if not self.out_cfg.get("generate_images", True):
                return ""
            return "\n".join(self._render_figure_segment(seg, for_docx=for_docx) for seg in block.figure_segments)

        if tag.name == "table":
            # 已不是 data_table/figure_table 的遗留表：开启表格渲染时，
            # 把「单格标题表」之类的纯文本兜底成段落，避免标题被静默丢弃。
            if self.doc_cfg.get("render_tables", True):
                leftover = self._render_leftover_table_text(tag)
                if leftover:
                    return leftover
            if self.doc_cfg.get("remove_tables", False):
                return ""
            if self.doc_cfg.get("table_to_text", True):
                return self._table_to_text(tag)
            return ""

        if block.role == "figure_image":
            if not self.out_cfg.get("generate_images", True):
                return ""
            return self._render_inline_image_container(tag, for_docx=for_docx)

        if tag.name in {"ul", "ol"}:
            return self._render_list(tag)

        text = normalize_text(tag.get_text(" ", strip=True))
        if not text:
            return ""

        inner = self._render_inline_children(
            tag, for_docx=for_docx,
            include_images=(block.has_image
                            and not self.doc_cfg.get("remove_images", True)
                            and self.out_cfg.get("generate_images", True)))
        inner = self._cleanup_inline_html(inner)
        if not inner:
            return ""

        heading_level = block.heading_level
        if heading_level in {1, 2, 3}:
            if for_docx:
                return self._render_heading_spacer() + f"<p><strong>{inner}</strong></p>"
            if self.doc_cfg.get("use_custom_heading_blocks", True):
                return self._render_custom_heading(heading_level, inner)

        is_heading_like = self.doc_cfg.get("heading_as_bold_paragraph", True) and self._is_heading_like(tag, text)
        if is_heading_like:
            inner = f"<strong>{inner}</strong>"
        return self._render_body_paragraph_html(inner)

    def _annotate_heading_levels(self, blocks: List[Block]) -> None:
        outline = list(self.heading_outline or [])
        occurrences = self.heading_text_occurrences or {}
        occ_seen: Dict[str, int] = {}

        # 左对齐消费位次。已知限制（裁定接受）：docx 里更早的重复出现（封面/正文
        # 范围外）若没有对应块，真标题会误领正文位次——真实报告中该文本重复只出现在
        # 正文清单场景（清单块在前且在正文内），而右对齐在真标题块带编号前缀
        # （键不匹配、不计数）时反而把清单块推到标题位次，实测回归。
        def occurrence_level(text: str) -> Optional[int]:
            key = heading_occurrence_key(text)
            levels = occurrences.get(key)
            if not levels:
                return None
            idx = occ_seen.get(key, 0)
            occ_seen[key] = idx + 1
            return levels[idx] if idx < len(levels) else levels[-1]

        if outline:
            pos = 0
            total = len(outline)
            for block in blocks:
                if block.drop or not block.text:
                    continue
                occ = occurrence_level(block.text)

                # Prefer explicit HTML-exported style names when available.
                detected = self._detect_heading_level(block.tag, block.text)
                if detected in {1, 2, 3}:
                    block.heading_level = detected
                    best_idx = -1
                    best_score = 0
                    for idx in range(pos, total):
                        if int(outline[idx].get("level", 0) or 0) != detected:
                            continue
                        score = heading_text_match_score(block.text, outline[idx].get("text", ""))
                        if score > best_score:
                            best_score = score
                            best_idx = idx
                        if score >= 95:
                            break
                    if best_idx >= 0 and best_score >= 65:
                        pos = best_idx + 1
                    continue

                if pos >= total:
                    continue

                # 源文档里同一文本既有正文段又有标题段时，按出现次序对位：
                # 本次出现在源文档中是正文段(0)的，不得因大纲相似度被升级为标题
                # (如问卷报告开头逐字重复各问题标题的题目清单)。
                if occ == 0:
                    continue

                best_idx = -1
                best_score = 0
                for idx in range(pos, total):
                    score = heading_text_match_score(block.text, outline[idx].get("text", ""))
                    if score > best_score:
                        best_score = score
                        best_idx = idx
                    if score >= 95:
                        break

                if best_idx >= 0 and best_score >= 65:
                    block.heading_level = int(outline[best_idx].get("level", 0) or 0)
                    pos = best_idx + 1
            return

        for block in blocks:
            if block.drop or not block.text:
                continue
            block.heading_level = self._detect_heading_level(block.tag, block.text)

    def _tag_attribute_text(self, tag: Tag) -> str:
        pieces = []
        for value in tag.attrs.values():
            if isinstance(value, list):
                pieces.append(" ".join(str(x) for x in value))
            else:
                pieces.append(str(value))
        return " ".join(pieces)

    def _extract_style_name_candidates(self, tag: Tag) -> List[str]:
        candidates: List[str] = []
        for value in tag.attrs.values():
            values = value if isinstance(value, list) else [value]
            for raw_value in values:
                raw_text = str(raw_value)
                if raw_text:
                    candidates.append(raw_text)
                for match in re.finditer(r'mso-style-name\s*:\s*"?([^;"\']+)"?', raw_text, re.I):
                    style_name = normalize_text(match.group(1))
                    if style_name:
                        candidates.append(style_name)
        class_values = tag.get("class", []) or []
        for class_name in class_values:
            candidates.append(str(class_name))
        return [normalize_text(x).lower() for x in candidates if normalize_text(x)]

    def _matches_heading_style_name(self, tag: Tag, level: int) -> bool:
        style_names = [normalize_text(x).lower() for x in self.matching.get(f"heading_level_{level}_style_names", []) if normalize_text(x)]
        if not style_names:
            return False
        candidates = self._extract_style_name_candidates(tag)
        for candidate in candidates:
            for style_name in style_names:
                if candidate == style_name or style_name in candidate:
                    return True
        return False

    def _detect_heading_level(self, tag: Tag, text: str) -> int:
        if self._matches_heading_style_name(tag, 3):
            return 3
        if self._matches_heading_style_name(tag, 2):
            return 2
        if self._matches_heading_style_name(tag, 1):
            return 1

        if not self.doc_cfg.get("use_heading_style_names_only", True):
            if self.re_heading_level_3.search(text):
                return 3
            if self.re_heading_level_2.search(text):
                return 2
            if self.re_heading_level_1.search(text):
                return 1
        return 0

    def _render_heading_spacer(self) -> str:
        margin = self.out_cfg.get("heading_spacer_margin_css", "0")
        text_indent = self.out_cfg.get("heading_spacer_text_indent_css", "0")
        text_align = self.out_cfg.get("heading_spacer_text_align", "left")
        return f'<p style="margin:{margin}; text-indent:{text_indent}; text-align:{text_align};"><br/></p>'

    def _render_body_paragraph_html(self, inner_html: str) -> str:
        font_family = self.out_cfg['font_family']
        font_size = self.out_cfg['font_size_px']
        line_height = self.out_cfg['line_height']
        indent = self.out_cfg['first_line_indent_em']
        margin_bottom = self.out_cfg['paragraph_margin_bottom_em']
        margin_top = self.out_cfg.get("paragraph_margin_top_css", "0")
        margin_right = self.out_cfg.get("paragraph_margin_right_css", "0")
        margin_left = self.out_cfg.get("paragraph_margin_left_css", "0")
        text_align = self.out_cfg.get("paragraph_text_align", "left")
        return (
            f'<p style="margin:{margin_top} {margin_right} {margin_bottom}em {margin_left}; '
            f'text-indent:{indent}em; text-align:{text_align}; '
            f'font-family:{font_family}; font-size:{font_size}px; line-height:{line_height};">{inner_html}</p>'
        )


    def _strip_weekly_heading_number(self, text: str) -> str:
        plain = normalize_text(BeautifulSoup(text, 'html.parser').get_text(' ', strip=True)) if '<' in text else normalize_text(text)
        plain = re.sub(
            r'^\s*(?:第\s*)?\d+(?:[.．]\d+)*'
            r'(?:[.．、]\s*|\s*[章节篇部]\s*|\s+(?![年月日季度QqHh号]))',
            '', plain)
        return html.escape(plain)

    def _render_custom_heading(self, level: int, inner: str) -> str:
        font_family = self.out_cfg.get("heading_font_family", self.out_cfg["font_family"])
        margin_top = self.out_cfg.get("heading_margin_top_em", 1.15)
        margin_bottom = self.out_cfg.get("heading_margin_bottom_em", 0.55)
        margin_bottom_l1 = self.out_cfg.get("heading_level_1_margin_bottom_em", margin_bottom)
        padding_v = self.out_cfg.get("heading_inline_padding_vertical_em", 0.22)
        padding_h = self.out_cfg.get("heading_inline_padding_horizontal_em", 0.55)
        radius_px = self.out_cfg.get("heading_border_radius_px", 8)
        text_indent = self.out_cfg.get("heading_text_indent_css", "0")
        margin_horizontal = self.out_cfg.get("heading_margin_horizontal_css", "0")
        spacer = self._render_heading_spacer()

        if level == 1 and self.config.get("_mode") == "weekly":
            heading_text = self._strip_weekly_heading_number(inner)
            bg_url = self.out_cfg.get("weekly_heading_level_1_background_image", "https://mmbiz.qpic.cn/mmbiz_png/fgnkxfGnnkRntkTEhlnCAXsPQXKEx3C0usHib3AgibeKROMf1PosItCxXbbJoAYibulRGR6JfbyE3fonvhN4pGtibw/640?wx_fmt=png")
            width_px = self.out_cfg.get("weekly_heading_level_1_width_px", 270)
            height_px = self.out_cfg.get("weekly_heading_level_1_height_px", 60)
            text_width_px = self.out_cfg.get("weekly_heading_level_1_text_width_px", 200)
            text_height_px = self.out_cfg.get("weekly_heading_level_1_text_height_px", 30)
            offset_left_px = self.out_cfg.get("weekly_heading_level_1_offset_left_px", 30)
            offset_top_px = self.out_cfg.get("weekly_heading_level_1_offset_top_px", 20)
            font_size = self.out_cfg.get("weekly_heading_level_1_font_size_px", 18)
            text_color = self.out_cfg.get("weekly_heading_level_1_text_color", "rgb(77, 12, 12)")
            line_height_px = self.out_cfg.get("weekly_heading_level_1_line_height_px", 30)
            font_weight = self.out_cfg.get("weekly_heading_level_1_font_weight", 400)
            text_align = self.out_cfg.get("weekly_heading_level_1_text_align", "center")
            outer_margin_horizontal = self.out_cfg.get("weekly_heading_level_1_outer_margin_horizontal_css", "auto")
            padding = self.out_cfg.get("weekly_heading_level_1_padding_css", "0")
            background_margin = self.out_cfg.get("weekly_heading_level_1_background_margin_css", "0 auto")
            background_size = self.out_cfg.get("weekly_heading_level_1_background_size", "100% 100%")
            background_repeat = self.out_cfg.get("weekly_heading_level_1_background_repeat", "no-repeat")
            background_position = self.out_cfg.get("weekly_heading_level_1_background_position", "center")
            text_margin_right = self.out_cfg.get("weekly_heading_level_1_text_margin_right_css", "0")
            text_margin_bottom = self.out_cfg.get("weekly_heading_level_1_text_margin_bottom_css", "0")
            # 背景框用 min-height 而非固定 height：标题折成两行时框随内容撑高，不按
            # overflow:hidden 裁掉第二行（《内容排版规范》1.5.2 内容溢出容器）。
            return (
                spacer +
                f'<section style="margin:{margin_top}em {outer_margin_horizontal} '
                f'{margin_bottom_l1}em {outer_margin_horizontal}; padding:{padding}; '
                f'text-indent:{text_indent}; text-align:{text_align}; '
                f'max-width:100%; box-sizing:border-box;">'
                f'<section style="display:inline-block; width:{width_px}px; max-width:100%; '
                f'min-height:{height_px}px; margin:{background_margin}; padding:{padding}; '
                f'overflow:hidden; background-image:url({bg_url}); background-size:{background_size}; '
                f'background-repeat:{background_repeat}; background-position:{background_position}; box-sizing:border-box;">'
                f'<section style="margin:{offset_top_px}px {text_margin_right} {text_margin_bottom} '
                f'{offset_left_px}px; padding:{padding}; '
                f'width:{text_width_px}px; max-width:calc(100% - {offset_left_px}px); '
                f'min-height:{text_height_px}px; color:{text_color}; font-size:{font_size}px; '
                f'font-family:{font_family}; font-weight:{font_weight}; text-align:{text_align}; '
                f'line-height:{line_height_px}px; box-sizing:border-box; overflow-wrap:break-word; '
                f'word-break:break-word; white-space:normal;">{heading_text}</section>'
                f'</section>'
                f'</section>'
            )

        if level == 1:
            font_size = self.out_cfg.get("heading_level_1_font_size_px", 17)
            text_color = self.out_cfg.get("heading_level_1_text_color", "#D97A45")
            bracket_color = self.out_cfg.get("heading_level_1_bracket_color", "#5F8FE8")
            border_px = self.out_cfg.get("heading_level_1_border_px", 6)
            text_align = self.out_cfg.get("heading_level_1_text_align", "center")
            border_style = self.out_cfg.get("heading_level_1_border_style", "solid")
            background = self.out_cfg.get("heading_level_1_background_color", "#ffffff")
            gap_horizontal_px = self.out_cfg.get("heading_level_1_gap_horizontal_px", 8)
            gap_extra_px = self.out_cfg.get("heading_level_1_gap_border_extra_px", 2)
            line_height = self.out_cfg.get("heading_level_1_line_height", 1.6)
            font_weight = self.out_cfg.get("heading_level_1_font_weight", 700)
            # 外层居中容器用 <section> 而不是 <div>：微信编辑器粘贴时会 unwrap <div>
            # 并丢掉 text-align:center，导致大标题左对齐；<section> 在微信里能稳定保留
            # text-align（周报版 banner 标题已用此方式）。内层是块级 <section>，不能用 <p>。
            return (
                spacer +
                f'<section style="margin:{margin_top}em {margin_horizontal} '
                f'{margin_bottom_l1}em {margin_horizontal}; '
                f'text-indent:{text_indent}; text-align:{text_align};">'
                f'<section style="display:inline-block; max-width:100%; box-sizing:border-box; '
                f'overflow-wrap:break-word; word-break:break-word; border-radius:{radius_px}px; '
                f'border-width:{border_px}px; border-style:{border_style}; border-color:{bracket_color}; '
                f'background:{background};">'
                f'<section style="margin:-{border_px}px {gap_horizontal_px}px; '
                f'border-top:{border_px + gap_extra_px}px {border_style} {background}; '
                f'border-bottom:{border_px + gap_extra_px}px {border_style} {background}; box-sizing:border-box;">'
                f'<section style="display:block; padding:{padding_v}em {padding_h}em; '
                f'color:{text_color}; font-size:{font_size}px; font-family:{font_family}; '
                f'font-weight:{font_weight}; line-height:{line_height}; text-align:{text_align}; '
                f'white-space:normal; overflow-wrap:break-word; word-break:break-word; '
                f'box-sizing:border-box;">{inner}</section>'
                f'</section>'
                f'</section>'
                f'</section>'
            )

        font_size = self.out_cfg.get(f"heading_level_{level}_font_size_px", self.out_cfg.get("heading_level_2_font_size_px", 14))
        text_color = self.out_cfg.get("heading_level_23_text_color", "#FFFFFF")
        background = self.out_cfg.get("heading_level_23_background_color", "#76A2E9")
        line_height = self.out_cfg.get("heading_level_23_line_height", 1.5)
        text_align = self.out_cfg.get("heading_level_23_text_align", "left")
        font_weight = self.out_cfg.get("heading_level_23_font_weight", 700)
        return (
            spacer +
            f'<p style="margin:{margin_top}em {margin_horizontal} {margin_bottom}em '
            f'{margin_horizontal}; text-indent:{text_indent}; '
            f'text-align:{text_align}; line-height:{line_height};">'
            f'<span style="display:inline-block; max-width:100%; vertical-align:top; '
            f'padding:{padding_v}em {padding_h}em; font-family:{font_family}; font-size:{font_size}px; '
            f'font-weight:{font_weight}; line-height:{line_height}; color:{text_color}; '
            f'background:{background}; border-radius:{radius_px}px; white-space:normal; '
            f'text-align:{text_align}; word-break:break-word; overflow-wrap:break-word; '
            f'box-sizing:border-box;">{inner}</span>'
            f'</p>'
        )

    def _is_heading_like(self, tag: Tag, text: str) -> bool:
        if self.doc_cfg.get("use_heading_style_names_only", True):
            return self._matches_heading_style_name(tag, 1) or self._matches_heading_style_name(tag, 2) or self._matches_heading_style_name(tag, 3)
        class_text = ((" ".join(tag.get("class", [])) if tag.get("class") else "") + " " + self._tag_attribute_text(tag)).lower()
        if "heading" in class_text:
            return True
        if self.re_heading_like.search(text):
            if len(text) <= 120:
                return True
        return False

    def _table_to_text(self, table: Tag) -> str:
        paragraphs: List[str] = []
        for tr in table.find_all("tr"):
            cells = []
            for cell in tr.find_all(["td", "th"], recursive=False):
                text = normalize_text(cell.get_text(" ", strip=True))
                if text:
                    cells.append(html.escape(text))
            if cells:
                paragraphs.append(self._render_body_paragraph_html(' | '.join(cells)))
        return "\n".join(paragraphs)

    def _render_inline_image_container(self, tag: Tag, for_docx: bool = False) -> str:
        images = []
        for node in tag.find_all(True):
            name = (node.name or "").lower()
            if name not in {"img", "v:imagedata"}:
                continue
            src = node.get("src", "") or node.get("o:href", "") or node.get("href", "")
            resolved = self._register_asset(src, inline=not for_docx)
            if not resolved:
                continue
            alt = html.escape(node.get("alt", "") or "")
            image_margin = self.out_cfg.get("image_margin_css", "0 auto")
            image_border = self.out_cfg.get("image_border_css", "0")
            images.append(
                f'<img src="{html.escape(resolved)}" alt="{alt}" '
                f'style="display:block; max-width:100%; height:auto; margin:{image_margin}; border:{image_border};" />'
            )
        if not images:
            return ""
        return self._render_figure_images_html(images)

    def _render_list(self, list_tag: Tag) -> str:
        items = []
        font_family = self.out_cfg['font_family']
        font_size = self.out_cfg['font_size_px']
        line_height = self.out_cfg['line_height']
        margin_bottom = self.out_cfg['paragraph_margin_bottom_em']
        item_margin_vertical = self.out_cfg.get("list_item_margin_vertical_em", 0.2)
        item_margin_horizontal = self.out_cfg.get("list_item_margin_horizontal_css", "0")
        text_align = self.out_cfg.get("list_text_align", "left")
        margin_top = self.out_cfg.get("list_margin_top_css", "0")
        margin_right = self.out_cfg.get("list_margin_right_css", "0")
        margin_left = self.out_cfg.get("list_margin_left_em", 1.4)
        padding_left = self.out_cfg.get("list_padding_left_em", 1.2)
        for li in list_tag.find_all("li", recursive=False):
            inner = self._render_inline_children(li)
            inner = self._cleanup_inline_html(inner)
            if inner:
                items.append(
                    f'<li style="margin:{item_margin_vertical}em {item_margin_horizontal}; '
                    f'font-family:{font_family}; font-size:{font_size}px; line-height:{line_height}; '
                    f'text-align:{text_align};">{inner}</li>'
                )
        if not items:
            return ""
        tag_name = "ol" if list_tag.name == "ol" else "ul"
        return (
            f'<{tag_name} style="margin:{margin_top} {margin_right} {margin_bottom}em {margin_left}em; '
            f'padding-left:{padding_left}em; font-family:{font_family}; font-size:{font_size}px; '
            f'line-height:{line_height};">'
            + "".join(items) + f"</{tag_name}>"
        )

    def _render_inline_children(self, node: Tag, preserve_color: bool = False,
                                for_docx: bool = False, include_images: bool = False) -> str:
        return "".join(
            self._render_inline(
                child, preserve_color=preserve_color,
                for_docx=for_docx, include_images=include_images)
            for child in node.children)

    def _render_inline(self, node, preserve_color: bool = False,
                       for_docx: bool = False, include_images: bool = False) -> str:
        if isinstance(node, NavigableString):
            text = str(node).replace("\xa0", " ")
            return html.escape(text)

        if not isinstance(node, Tag):
            return ""

        name = node.name.lower()
        if name in {"img", "v:imagedata"} and include_images:
            src = node.get("src", "") or node.get("o:href", "") or node.get("href", "")
            resolved = self._register_asset(src, inline=not for_docx)
            if not resolved:
                return ""
            alt = html.escape(node.get("alt", "") or "")
            image_margin = self.out_cfg.get("image_margin_css", "0 auto")
            image_border = self.out_cfg.get("image_border_css", "0")
            return (
                f'<img src="{html.escape(resolved)}" alt="{alt}" '
                f'style="display:inline-block; max-width:100%; height:auto; vertical-align:middle; '
                f'margin:{image_margin}; border:{image_border};" />'
            )
        if name == "v:shape" and include_images:
            return "".join(
                self._render_inline(
                    child, preserve_color=preserve_color,
                    for_docx=for_docx, include_images=True)
                for child in node.children
            )
        if name in {"script", "style", "img", "object", "svg", "v:shape", "v:imagedata"}:
            return ""
        if name == "br":
            return "<br/>"

        style = node.get("style", "") or ""
        style_parts = []
        if preserve_color:
            extracted = extract_color_style(style)
            if extracted:
                style_parts.append(extracted.strip().rstrip(';'))

        child_html = "".join(
            self._render_inline(
                child, preserve_color=preserve_color,
                for_docx=for_docx, include_images=include_images)
            for child in node.children
        )
        is_bold = self._tag_is_bold(node)
        is_underline = self._tag_is_underline(node)

        if is_bold:
            font_weight = self.out_cfg.get("inline_bold_font_weight", 700)
            style_parts.append(f"font-weight:{font_weight}")
        if is_underline:
            decoration = self.out_cfg.get("inline_underline_text_decoration", "underline")
            offset = self.out_cfg.get("inline_underline_offset_em", 0.08)
            style_parts.append(f"text-decoration:{decoration}")
            style_parts.append(f"text-underline-offset:{offset}em")

        if style_parts:
            merged = "; ".join(style_parts)
            return f'<span style="{html.escape(merged, quote=True)}">{child_html}</span>'
        return child_html


    @staticmethod
    def _tag_is_bold(tag: Tag) -> bool:
        if tag.name.lower() in {"b", "strong"}:
            return True
        style = (tag.get("style", "") or "").lower().replace(" ", "")
        if "font-weight:bold" in style:
            return True
        if re.search(r"font-weight:(700|800|900)\b", style):
            return True
        return False

    @staticmethod
    def _tag_is_underline(tag: Tag) -> bool:
        if tag.name.lower() == "u":
            return True
        style = (tag.get("style", "") or "").lower().replace(" ", "")
        if "text-decoration:underline" in style or "text-decoration-line:underline" in style:
            return True
        return False

    def _cleanup_inline_html(self, html_text: str) -> str:
        html_text = re.sub(r"\s+", " ", html_text)
        html_text = re.sub(r"(?:<strong>\s*</strong>|<u>\s*</u>|<span[^>]*>\s*</span>)", "", html_text)
        html_text = re.sub(r"<strong><strong>(.*?)</strong></strong>", r"<strong>\1</strong>", html_text)
        html_text = re.sub(r"<u><u>(.*?)</u></u>", r"<u>\1</u>", html_text)
        html_text = re.sub(r"\s*(<br/>\s*)+\s*", "<br/>", html_text)
        return html_text.strip()

    def _register_asset(self, src: str, inline: bool = False) -> str:
        """把图片地址换成成品里能用的地址（inline 时是 data: URI）。

        取不到时返回空串或原样地址，调用方一律跳过该图——外观上看不出少了图，
        所以这里同时把「没换成 data:」记进 render_problems 供调用方判断成品是否完整。"""
        self._note_visited_image(src)
        resolved = self._resolve_asset(src, inline=inline)
        if not resolved:
            self._note_render_problem(f"图片地址为空或无法解析：{(src or '').strip()[:120]!r}")
        elif inline and not resolved.startswith(("data:", "http://", "https://")):
            # inline 输出必须内嵌字节；仍是原地址说明文件不在本机，成品里会是一个死链
            self._note_render_problem(f"图片未能内嵌，源文件不可读：{resolved[:120]!r}")
        return resolved

    def _resolve_asset(self, src: str, inline: bool = False) -> str:
        src = (src or "").strip()
        if not src:
            return ""
        if src.startswith("data:"):
            return src

        parsed = urlparse(src)
        if parsed.scheme in {"http", "https"}:
            return src
        if parsed.scheme == "file":
            raw_path = unquote(parsed.path)
            if parsed.netloc and parsed.netloc.lower() != "localhost":
                # UNC 地址(file://server/share/...)：netloc 是服务器名，不能丢；
                # file://localhost/C:/... 是本机地址，按无 netloc 处理
                raw_path = "\\\\" + unquote(parsed.netloc) + raw_path.replace("/", "\\")
            elif raw_path.startswith("/") and re.match(r"^/[A-Za-z]:", raw_path):
                raw_path = raw_path[1:]
            abs_path = Path(raw_path)
        else:
            if self.current_html_path is None:
                return src
            cleaned_src = src.split("#", 1)[0].split("?", 1)[0]
            abs_path = (self.current_html_path.parent / cleaned_src).resolve()

        if not abs_path.exists() or not abs_path.is_file():
            return src

        if inline:
            mime_type = mimetypes.guess_type(str(abs_path))[0] or "application/octet-stream"
            data = read_bytes_retrying(abs_path)
            if mime_type == "image/png":
                data = flatten_transparent_png(data)
            data_b64 = base64.b64encode(data).decode("ascii")
            return f"data:{mime_type};base64,{data_b64}"

        if abs_path in self.asset_map:
            return self.asset_map[abs_path]

        suffix = abs_path.suffix.lower() or ".bin"
        base_name = re.sub(r"[^A-Za-z0-9._-]+", "_", abs_path.stem).strip("._-") or "asset"
        candidate = base_name + suffix
        if candidate in self.used_asset_names:
            self.used_asset_names[candidate] += 1
            candidate = f"{base_name}_{self.used_asset_names[candidate]}{suffix}"
        else:
            self.used_asset_names[candidate] = 1

        rel_path = f"{self.asset_subdir_name}/{candidate}"
        self.asset_map[abs_path] = rel_path
        return rel_path

    def _wrap_full_html(self, body_html: str) -> str:
        font_family = self.out_cfg["font_family"]
        font_size = self.out_cfg["font_size_px"]
        line_height = self.out_cfg["line_height"]
        page_margin = self.out_cfg.get("page_margin_css", "24px auto")
        page_padding = self.out_cfg.get("page_padding_css", "0 20px")
        page_max_width = self.out_cfg.get("page_max_width_px", 860)
        page_background = self.out_cfg.get("page_background_color", "#ffffff")
        page_text_color = self.out_cfg.get("page_text_color", "#000000")
        # body 显式 text-align:left：复制正文时结果页把计算对齐写进每个元素，未声明对齐的
        # 节点会写成 start，公众号《内容排版规范》1.6 判为非标准值。
        return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>cleaned_wechat_output</title>
</head>
<body style="margin:{page_margin}; padding:{page_padding}; max-width:{page_max_width}px; background:{page_background}; color:{page_text_color}; font-family:{font_family}; font-size:{font_size}px; line-height:{line_height}; text-align:left; word-break:break-word;">
{body_html}
</body>
</html>
"""


def flatten_transparent_png(data: bytes, background=(255, 255, 255)) -> bytes:
    """有透明像素的 PNG 铺到白底上重新编码；不透明或无法解码的原样返回。

    公众号深色模式把正文底色换成 #191919，透明底图表里的黑灰文字会看不见
    （《内容排版规范》4.3.2）。页面本身是白底，铺白后浅色模式下外观不变。"""
    try:
        from io import BytesIO
        from PIL import Image
        with Image.open(BytesIO(data)) as image:
            if not (image.mode in ("RGBA", "LA", "PA")
                    or (image.mode == "P" and "transparency" in image.info)
                    or (image.mode in ("RGB", "L") and "transparency" in image.info)):
                return data
            rgba = image.convert("RGBA")
            icc_profile = image.info.get("icc_profile")
        if rgba.getchannel("A").getextrema()[0] == 255:
            return data
        flat = Image.new("RGB", rgba.size, background)
        flat.paste(rgba, mask=rgba.getchannel("A"))
        out = BytesIO()
        flat.save(out, "PNG", **({"icc_profile": icc_profile} if icc_profile else {}))
        return out.getvalue()
    except Exception:
        return data


def safe_stem_ascii(path: Path) -> str:
    stem = path.stem
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem)
    stem = re.sub(r"_+", "_", stem).strip("._-")
    if stem:
        return stem[:80]
    digest = __import__("hashlib").sha1(str(path).encode("utf-8", errors="ignore")).hexdigest()[:10]
    return f"doc_{digest}"




def copy_used_assets(asset_map: Dict[Path, str], out_dir: Path) -> None:
    for src_path, rel_path in asset_map.items():
        dst_path = out_dir / Path(rel_path)
        dst_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_path, dst_path)


def build_front_matter_assets(src: Path, temp_dir: Path, stem: str, cleaner: "HtmlCleaner",
                            word: WordApp, cfg: dict, summary_extractor):
    """计算前置内容资产, 返回 (summary_html, toc_png_path)。
    front_matter.enable=False 或摘要/目录都关闭时直接返回 ("", None), 整段不触碰 Word。
    任一步失败都只影响该步(优雅降级), 不影响正文。"""
    fm_cfg = cfg.get("front_matter", {}) or {}
    if not fm_cfg.get("enable", False):
        return "", None
    want_summary = bool(fm_cfg.get("include_summary", True)) and summary_extractor is not None
    want_toc = bool(fm_cfg.get("include_toc_image", True))
    if not want_summary and not want_toc:
        return "", None

    summary_html: str = ""
    toc_list: List[List[dict]] = []
    try:
        summary_html, toc_list = word.extract_front_matter(src, summary_extractor, want_summary, want_toc)
    except Exception as exc:
        print(f"  [warn] front matter extract failed: {exc}")
        summary_html, toc_list = "", []
    if want_summary and summary_html:
        print("  -> summary extracted")
    elif want_summary:
        print("  -> summary skipped: no summary content matched")

    toc_png_path = None
    if want_toc:
        try:
            entries = build_toc_entries(toc_list, cleaner.heading_outline, fm_cfg)
        except Exception as exc:
            print(f"  -> TOC image skipped: {exc}")
            entries = None
        if entries:
            try:
                import sys
                if str(SCRIPT_DIR) not in sys.path:
                    sys.path.insert(0, str(SCRIPT_DIR))
                import toc_image
                png = temp_dir / f"{stem}_toc.png"
                toc_image.render_toc(entries, str(png), opts=fm_cfg.get("toc_image", {}))
                toc_png_path = str(png)
                print(f"  -> TOC image: {len(entries)} entries")
            except Exception as exc:
                print(f"  -> TOC image skipped: {exc}")
                toc_png_path = None
        elif entries is not None:
            print("  -> TOC image skipped: no content TOC entries found")
    return summary_html, toc_png_path




def build_content_toc_fragment(src: Path, temp_dir: Path, word: WordApp, cfg: dict) -> str:
    try:
        cfg.setdefault("front_matter", {}).update(enable=True, include_summary=False, include_toc_image=True)
        cleaner = HtmlCleaner(cfg)
        probe = src if src.suffix.lower() == ".docx" else temp_dir / "toc_probe.docx"
        if probe != src:
            word.export_docx(src, probe)
        cleaner.set_heading_outline(extract_docx_heading_outline(probe, cfg["matching"]))
        _, png = build_front_matter_assets(src, temp_dir, "summary", cleaner, word, cfg, None)
        return cleaner._render_toc_image(png) if png else ""
    except Exception as exc:
        print(f"  -> TOC image skipped: {exc}")
        return ""


def relinked_export_source(src: Path, temp_dir: Path, cfg: dict) -> Optional[Path]:
    """正文导出用的改写副本：本地图表改为外链到本次 temp_dir 下一个从不创建的工作簿（chart_relink.py）。

    只用于 export_filtered_html；摘要、目录与样式探测仍用原副本。仅 .docx；开关
    document.relink_local_charts 关闭、整份命中守卫、没有可改写的图表或改写出错时返回 None（用原副本）。
    在 stdout 记一行「[relink] …」，不进成品。"""
    if src.suffix.lower() != ".docx" or not cfg.get("document", {}).get("relink_local_charts", False):
        return None
    target = temp_dir / "relink" / src.name
    try:
        import chart_relink

        result = chart_relink.relink_local_charts(src, target, temp_dir / "relink-missing" / "chart-data.xlsx")
    except Exception as exc:
        print(f"  [relink] failed: {type(exc).__name__}: {exc}")
        return None
    if result["document_skip"]:
        print(f"  [relink] skipped document: {result['document_skip']}")
        return None
    print("  [relink] " + " ".join(f"{key}={value}" for key, value in sorted(result["counts"].items())))
    return target if result["changed"] else None


def read_bytes_retrying(path: Path, attempts: int = 6, delay: float = 0.5) -> bytes:
    """读刚写出的文件：安全软件扫描新文件时会短暂独占，稍等重读；一直读不了才抛出。"""
    import time
    for attempt in range(attempts):
        try:
            return path.read_bytes()
        except PermissionError:
            if attempt == attempts - 1:
                raise
            time.sleep(delay)


def locate_pages(word: "WordApp", src: Path, failed_blocks: List[dict]) -> Dict[int, List[int]]:
    """按失败块开头文字在原文档里查找所在页（wdActiveEndPageNumber）。同一句可能也出现在首页摘要里，
    所以列出全部匹配页（最多 3 处），不猜哪一处；只供定位提示，任何失败都只是不给页码。"""
    pages: Dict[int, List[int]] = {}
    doc = None
    try:
        doc = word.open_document(src)
        # 无头模式要先分页，Information 才给出真实页码（同目录页码的读取）；取物理页序，与 Word 状态栏一致
        try:
            doc.Repaginate()
        except Exception:
            pass
        for item in failed_blocks:
            # 导出 HTML 会在中文与数字间多出空格：取开头里最长的一段不含空白的连续文字去找
            pieces = re.split(r"\s+", (item.get("opening") or "").replace("^", " "))
            needle = max(pieces, key=len, default="")[:20]
            if len(needle) < 4:
                continue
            found = doc.Content
            finder = found.Find
            finder.ClearFormatting()
            hits: List[int] = []
            while len(hits) < 3 and finder.Execute(FindText=needle, MatchCase=False, MatchWildcards=False,
                                                    Forward=True, Wrap=0):
                page = int(found.Information(3))
                if page not in hits:
                    hits.append(page)
                found.Collapse(0)  # wdCollapseEnd：从这一处之后继续找
            if hits:
                pages[item["idx"]] = hits
    except Exception as exc:
        print(f"  [warn] could not locate Word page numbers: {exc}")
    finally:
        close_document(doc)
    return pages


def _remove_export_leftovers(folder: Path, before: set) -> bool:
    """删掉 folder 里导出前不存在的条目；有删不掉的返回 False。"""
    clean = True
    for path in folder.iterdir():
        if path.name in before:
            continue
        try:
            if path.is_dir():
                shutil.rmtree(path)
            else:
                path.unlink()
        except OSError:
            clean = False
    return clean


def _close_documents(word: WordApp, paths: Iterable[Path]) -> bool:
    """关掉仍开着的、全名是 paths 之一的文档（SaveAs 成功后文档全名已变成导出的 HTML 路径，
    所以调用方要同时传改写副本与导出路径），关完再核对一遍确实没有了。
    Word 不可用（COM 失效、已退出、文档关不掉）时返回 False。"""
    targets = {Path(path).resolve() for path in paths}

    def still_open():
        documents = word.word.Documents
        return [documents(index) for index in range(int(documents.Count), 0, -1)
                if Path(str(documents(index).FullName)).resolve() in targets]

    try:
        for document in still_open():
            document.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        return not still_open()
    except Exception:
        return False


def export_body_html(word: WordApp, src: Path, relinked: Optional[Path], filtered_html: Path) -> None:
    """正文导出筛选 HTML。有改写副本先导出它；Word 在这一步抛错时丢弃这次导出写出的全部文件、
    用原副本在同一路径重导（Word 不可用或残留文件删不掉时先 restart 重建实例）。
    只接 Exception：用户取消（KeyboardInterrupt、进程被结束）不重试。"""
    if relinked is None:
        word.export_filtered_html(src, filtered_html)
        return
    folder = filtered_html.parent
    before = {path.name for path in folder.iterdir()}
    try:
        word.export_filtered_html(relinked, filtered_html)
        return
    except Exception as exc:
        print(f"  [relink] export failed, re-exporting original: {type(exc).__name__}: {exc}")
    healthy = _close_documents(word, (relinked, filtered_html))
    clean = _remove_export_leftovers(folder, before)
    if not (healthy and clean):
        word.restart()
        if not _remove_export_leftovers(folder, before):
            raise ProcessingError("改写副本导出失败后，残留的导出文件无法删除")
    word.export_filtered_html(src, filtered_html)


def process_one_file(src: Path, out_dir: Path, cleaner: HtmlCleaner, word: WordApp, cfg: dict,
                    summary_extractor=None) -> dict:
    stem = safe_stem_ascii(src)
    temp_dir = Path(tempfile.mkdtemp(prefix="wechat_docx_clean_"))
    filtered_html = temp_dir / f"{stem}.html"
    style_probe_docx = src if src.suffix.lower() == ".docx" else temp_dir / f"{stem}_style_probe.docx"

    try:
        if style_probe_docx != src:
            word.export_docx(src, style_probe_docx)
        cleaner.set_heading_outline(extract_docx_heading_outline(style_probe_docx, cfg["matching"]))
        cleaner.set_heading_text_occurrences(
            extract_docx_heading_text_occurrences(style_probe_docx, cfg["matching"]))
        # 数据表格转图片：PNG 渲染到 temp 目录(在 base64 内联 / asset 复制期间必须还在)
        cleaner.generated_dir = temp_dir
        cleaner.source_docx_path = style_probe_docx
        cleaner._valuation_cache = None

        # ---- 前置内容(摘要 + 内容目录图)。enable=False 时整段跳过, 行为同旧版 ----
        summary_html, toc_png_path = build_front_matter_assets(
            src, temp_dir, stem, cleaner, word, cfg, summary_extractor
        )

        export_body_html(word, src, relinked_export_source(src, temp_dir, cfg), filtered_html)
        clean_result = cleaner.clean_html_file(
            filtered_html, f"{stem}_wechat_clean_files",
            summary_html=summary_html, toc_png_path=toc_png_path,
        )
        if cleaner.failed_blocks:
            cleaner.report_failed_blocks(locate_pages(word, src, cleaner.failed_blocks))
        # 外链资产只有 docx 输出才需要(HTML 走 base64 内联)。不生成 docx 就别复制资产文件夹。
        if cfg["output"].get("generate_docx", False) and clean_result.asset_map:
            copy_used_assets(clean_result.asset_map, out_dir)
        cleaned_html = clean_result.html_text
        cleaned_docx_html = clean_result.docx_html_text

        result = {"source": str(src)}
        html_out = out_dir / f"{stem}_wechat_clean.html"
        if cfg["output"].get("generate_html", True):
            html_out.write_text(cleaned_html, encoding="utf-8")
            result["html"] = str(html_out)

        if cfg["output"].get("generate_docx", True):
            html_for_docx = out_dir / f"{stem}_wechat_clean_docx_source.html"
            html_for_docx.write_text(cleaned_docx_html, encoding="utf-8")
            docx_out = out_dir / f"{stem}_wechat_clean.docx"
            word.convert_html_to_docx(html_for_docx, docx_out)
            result["docx"] = str(docx_out)

        return result
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)



def process_summary_file(src: Path, out_dir: Path, extractor: SummaryExtractor, word: WordApp, cfg: dict) -> dict:
    result = {"source": str(src)}
    doc = None
    try:
        doc = word.open_document(src)
        html_text = extractor.extract_html(doc)
        if cfg["output"].get("generate_html", True):
            suffix = cfg["output"].get("output_suffix", "_summary_extract.html")
            html_out = out_dir / f"{safe_stem_ascii(src)}{suffix}"
            html_out.write_text(html_text, encoding="utf-8")
            result["html"] = str(html_out)
        return result
    finally:
        close_document(doc)
