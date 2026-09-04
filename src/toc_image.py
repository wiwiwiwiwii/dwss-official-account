# -*- coding: utf-8 -*-
"""
内容目录 -> 图片 (PNG)

把深度报告的「内容目录」(章节目录, **不含图表目录**) 渲染成一张清晰图片,
插在公众号文章的「摘要」与「正文」之间。微信编辑器无法保留 Word 的目录域
(点线引导符 + 页码 + 层级缩进), 整段目录转图最稳。

调用方(wechat_docx_cleaner) 负责把 Word 目录域解析成 entries 传进来:

    entries = [
        {"level": 1, "title": "1 行业概况",   "page": 5},
        {"level": 2, "title": "1.1 市场规模", "page": 6},
        {"level": 1, "title": "2 公司分析",   "page": 9},
        ...
    ]

- level: 1/2/3 (越大缩进越深); page: int 或 None(无页码时只画标题)。
- 本模块只依赖 Pillow, 与 `table_image.py` 同系统(超采样 + 字体回退 + 东吴蓝配色)。
"""

import re

from PIL import Image, ImageDraw, ImageFont


# ----------------------------- 可调样式 -----------------------------
FONT_REGULAR = r"C:\Windows\Fonts\msyh.ttc"      # 微软雅黑
FONT_BOLD    = r"C:\Windows\Fonts\msyhbd.ttc"    # 微软雅黑 粗体

SS = 2                   # 超采样倍数(清晰度), 与 table_image.py 一致

PAD            = 24      # 整图外边距(逻辑像素)
TITLE_BAR_H    = 44      # 顶部「目录」标题栏高度
TITLE_FONT_PX  = 19      # 标题栏字号
L1_FONT_PX     = 15      # 一级条目字号
SUB_FONT_PX    = 14      # 二级及以下条目字号
PAGE_FONT_PX   = 14      # 页码字号
INDENT_STEP    = 26      # 每级缩进(逻辑像素)
ROW_VPAD       = 8       # 每行上下内边距
LINE_GAP       = 4       # 标题折行时的行距
L1_TOP_GAP     = 12      # 一级条目之上的额外留白(分组感)
LEADER_GAP     = 8       # 标题/页码与点线之间的间隙
LEADER_MIN     = 30      # 点线最短长度(避免标题贴着页码)
MIN_W          = 680     # 整图最小宽度(逻辑像素)
MAX_W          = 1080    # 整图最大宽度(逻辑像素), 超长标题折行
MAX_PX         = 40_000_000  # 整张 PNG 像素上限(超采样后), 超过即放弃(交主程序兜底)

# 配色(东吴蓝系, 与 table_image.py / 估值表保持一致)
C_BAR_BG   = (46, 90, 156)       # 标题栏蓝
C_BAR_FG   = (255, 255, 255)
C_L1_FG    = (38, 74, 130)       # 一级条目(深蓝, 加粗)
C_SUB_FG   = (51, 55, 61)        # 二级及以下(近黑)
C_LEADER   = (178, 190, 206)     # 点线引导符
C_PAGE_FG  = (90, 99, 115)       # 页码
C_BG       = (255, 255, 255)
# -------------------------------------------------------------------


# 字体回退链: 首选微软雅黑, 缺失时退到其它常见中文字体, 最后退 Pillow 内置位图字体,
# 避免在非标准 Windows 上渲染直接崩溃(崩了主程序会跳过目录图, 不影响正文/摘要)。
_FONT_FALLBACKS = [
    r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyhbd.ttc",
    r"C:\Windows\Fonts\simhei.ttf", r"C:\Windows\Fonts\simsun.ttc",
    r"C:\Windows\Fonts\Deng.ttf", r"C:\Windows\Fonts\arial.ttf",
]


def _font(path, px):
    for cand in [path] + _FONT_FALLBACKS:
        try:
            return ImageFont.truetype(cand, px)
        except Exception:
            continue
    try:
        return ImageFont.load_default()
    except Exception:
        return ImageFont.load_default()


def render_toc(entries, out_path, opts=None):
    """把内容目录 entries 渲染成 PNG。返回 (out_path, (w, h))。entries 为空抛 ValueError。"""
    opts = opts or {}
    title_text = str(opts.get("title", "目录") or "")
    show_bar   = bool(opts.get("show_title_bar", True))
    show_pages = bool(opts.get("show_page_numbers", True))

    entries = [e for e in (entries or []) if str(e.get("title", "")).strip()]
    if not entries:
        raise ValueError("empty toc entries")

    s = SS
    f_title = _font(FONT_BOLD, TITLE_FONT_PX * s)
    f_l1    = _font(FONT_BOLD, L1_FONT_PX * s)
    f_sub   = _font(FONT_REGULAR, SUB_FONT_PX * s)
    f_page  = _font(FONT_REGULAR, PAGE_FONT_PX * s)

    pad         = PAD * s
    indent_step = INDENT_STEP * s
    row_vpad    = ROW_VPAD * s
    line_gap    = LINE_GAP * s
    l1_top_gap  = L1_TOP_GAP * s
    leader_gap  = LEADER_GAP * s
    leader_min  = LEADER_MIN * s
    title_bar_h = TITLE_BAR_H * s

    probe = ImageDraw.Draw(Image.new("RGB", (4, 4)))

    def tw(text, font):
        return probe.textlength(text, font=font)

    def font_for(level):
        return f_l1 if level <= 1 else f_sub

    def color_for(level):
        return C_L1_FG if level <= 1 else C_SUB_FG

    has_page = show_pages and any(e.get("page") is not None for e in entries)
    page_w = 0.0
    if has_page:
        for e in entries:
            if e.get("page") is not None:
                page_w = max(page_w, tw(str(e["page"]), f_page))
    page_col = (leader_gap + leader_min + leader_gap + page_w) if has_page else 0.0

    # 整图宽度: 取「最长(缩进 + 标题)行」, 预留页码列, 夹在 [MIN_W, MAX_W]。
    natural = 0.0
    for e in entries:
        level = max(1, int(e.get("level", 1)))
        indent = pad + (level - 1) * indent_step
        natural = max(natural, indent + tw(str(e["title"]).strip(), font_for(level)))
    width = int(min(max(natural + page_col + pad, MIN_W * s), MAX_W * s))
    right_edge = width - pad
    page_x_right = right_edge                              # 页码右端对齐到右内边距
    leader_right = (right_edge - page_w - leader_gap) if has_page else right_edge

    # 标题折行(按字符, 兼容中文无空格)
    def wrap(text, font, max_w):
        text = re.sub(r"\s+", " ", str(text)).strip()
        if not text:
            return [""]
        out, cur = [], ""
        for ch in text:
            if cur and tw(cur + ch, font) > max_w:
                out.append(cur)
                cur = ch
            else:
                cur += ch
        if cur:
            out.append(cur)
        return out or [""]

    rows = []
    for e in entries:
        level = max(1, int(e.get("level", 1)))
        indent = pad + (level - 1) * indent_step
        font = font_for(level)
        page = e.get("page") if has_page else None
        page_str = str(page) if (page is not None) else None
        if page_str is not None:
            title_budget = (leader_right - leader_gap) - indent
        else:
            title_budget = right_edge - indent
        title_budget = max(title_budget, 60 * s)
        lines = wrap(e["title"], font, title_budget)
        rows.append({
            "indent": indent, "font": font, "color": color_for(level),
            "lines": lines, "page": page_str, "is_l1": level <= 1,
        })

    # 高度
    def row_text_h(r):
        n = len(r["lines"])
        return n * r["font"].size + (n - 1) * line_gap

    total_h = float(title_bar_h if show_bar else pad)
    for i, r in enumerate(rows):
        if r["is_l1"] and i > 0:
            total_h += l1_top_gap
        total_h += row_vpad + row_text_h(r) + row_vpad
    total_h += pad
    total_h = int(total_h)

    if width * total_h > MAX_PX:
        raise ValueError(f"toc image too large to render: {width}x{total_h}px")

    img = Image.new("RGB", (width, total_h), C_BG)
    d = ImageDraw.Draw(img)

    def draw_leaders(x0, x1, yc):
        if x1 - x0 < 4 * s:
            return
        dot_r = max(1, int(round(1.1 * s)))
        step = max(5 * s, dot_r * 5)
        x = x0
        while x <= x1:
            d.ellipse([x, yc - dot_r, x + 2 * dot_r, yc + dot_r], fill=C_LEADER)
            x += step

    y = 0
    if show_bar:
        d.rectangle([0, 0, width, title_bar_h], fill=C_BAR_BG)
        if title_text:
            tt_w = tw(title_text, f_title)
            d.text(((width - tt_w) / 2, (title_bar_h - f_title.size) / 2 - s),
                   title_text, font=f_title, fill=C_BAR_FG)
        y = title_bar_h
    else:
        y = pad

    for i, r in enumerate(rows):
        if r["is_l1"] and i > 0:
            y += l1_top_gap
        y += row_vpad
        n = len(r["lines"])
        font = r["font"]
        for li, ln in enumerate(r["lines"]):
            ty = y + li * (font.size + line_gap)
            d.text((r["indent"], ty), ln, font=font, fill=r["color"])
            if li == n - 1 and r["page"] is not None:
                line_w = tw(ln, font)
                lx0 = r["indent"] + line_w + leader_gap
                lx1 = leader_right - leader_gap
                yc = ty + font.size * 0.72
                draw_leaders(lx0, lx1, yc)
                px = page_x_right - tw(r["page"], f_page)
                py = ty + (font.size - f_page.size)
                d.text((px, py), r["page"], font=f_page, fill=C_PAGE_FG)
        y += row_text_h(r) + row_vpad

    if s != 1:
        img = img.resize((max(1, width // s), max(1, total_h // s)), Image.LANCZOS)
    img.save(out_path, "PNG")
    return out_path, img.size


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    demo = [
        {"level": 1, "title": "1 行业概况：韩国消费板块景气回升", "page": 4},
        {"level": 2, "title": "1.1 化妆品：出口链持续高增", "page": 5},
        {"level": 2, "title": "1.2 食品饮料：龙头份额稳固", "page": 8},
        {"level": 1, "title": "2 重点公司分析", "page": 12},
        {"level": 2, "title": "2.1 爱茉莉太平洋", "page": 12},
        {"level": 3, "title": "2.1.1 盈利预测与估值", "page": 14},
        {"level": 1, "title": "3 风险提示", "page": 18},
    ]
    path, size = render_toc(demo, "toc_demo.png")
    print(f"[OK] {path}  {size[0]}x{size[1]}px")
