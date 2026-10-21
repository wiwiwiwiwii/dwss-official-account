# -*- coding: utf-8 -*-
"""
通用表格 -> 图片 (PNG)

把深度报告正文里的**任意数据表格**渲染成一张清晰图片，方便直接插入公众号文章
（微信编辑器对复杂表格支持很差，整表转图最稳）。

与周报版 `valuation_table_image.py` 的区别：
- 周报版是**专表**渲染器：写死「估值表」的 11 个逻辑列、年份子表头、板块分隔。
- 本模块是**通用**渲染器：列数、合并单元格（横向 colspan / 纵向 rowspan）、表头行数、
  对齐、换行都按传入的网格结构自适应。深度报告的表格五花八门（政策表、盈利预测表、
  可比公司估值表的二级表头…），所以兼容性要求比周报大得多。

本模块**只依赖 Pillow**（不依赖 bs4/docx）。调用方（wechat_docx_cleaner）负责把
Word 导出的 HTML 表格解析成下面的 `spec` 结构后传进来，便于单测与解耦。

spec 结构::

    {
      "ncols": int,            # 网格总列数
      "nrows": int,            # 网格总行数（不含已被抽走的题注/数据来源行）
      "header_rows": int,      # 顶部按表头着色的行数（>=0）
      "cells": [               # 只列「起始单元格」(origin)；被合并覆盖的格不出现
        {
          "r": int, "c": int,          # 起始行/列（0 基）
          "rs": int, "cs": int,        # 纵向跨行 / 横向跨列（>=1）
          "lines": [str, ...],         # 单元格文本，按显式换行(<br>/<p>)拆好的多行
          "align": "l"|"c"|"r"|None,   # 水平对齐；None 表示由渲染器按内容判定
          "bold": bool,                # 是否加粗（表头行会强制加粗）
        }, ...
      ],
    }
"""

import sys

from PIL import Image, ImageDraw, ImageFont


# ----------------------------- 可调样式 -----------------------------
FONT_REGULAR = r"C:\Windows\Fonts\msyh.ttc"      # 微软雅黑
FONT_BOLD    = r"C:\Windows\Fonts\msyhbd.ttc"    # 微软雅黑 粗体

BODY_FONT_PX   = 14      # 正文(数据)字号(逻辑像素)
HEADER_FONT_PX = 14      # 表头字号

SS = 2                   # 超采样倍数(清晰度)

ROW_PAD_Y   = 7          # 单元格上下内边距
CELL_PAD_X  = 9          # 单元格左右内边距
MIN_COL_W   = 34         # 列最小宽度(逻辑像素)
MAX_COL_W   = 300        # 单列内容超过此宽度即换行(逻辑像素)
LINE_GAP    = 3          # 行内多行文字的行距(逻辑像素)

# 默认只画横向网格线(行分隔)，不画竖网格线。需要竖线时把它设为 True。
SHOW_VGRID  = False

THUMB_H     = 56         # 单元格内图片缩略图目标高度(逻辑像素)
MAX_TABLE_W = 1600       # 整表最大宽度(逻辑像素)；超宽表按比例压缩列宽，避免 PNG 过大/内存爆
MAX_TABLE_PX = 60_000_000  # 整张 PNG 的像素上限(超采样后)；超过即放弃渲染(交主程序文字兜底)，防 OOM

# 配色 (东吴蓝系，与估值表保持一致)
C_HEAD_BG    = (46, 90, 156)     # 表头蓝
C_HEAD_FG    = (255, 255, 255)
C_ROW_A      = (255, 255, 255)   # 数据行 交替色 A
C_ROW_B      = (242, 246, 252)   # 数据行 交替色 B
C_BODY_FG    = (33, 37, 41)
C_BORDER     = (200, 210, 224)   # 内部网格线
C_FRAME      = (150, 175, 205)   # 外框/表头分隔线
# -------------------------------------------------------------------



# 字体回退链：首选微软雅黑，缺失时退到其它常见中文字体，最后退到 Pillow 内置位图字体，
# 避免在非标准 Windows 上整张表渲染直接崩掉(崩了会落到主程序的文字兜底)。
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


def _load_thumb(path, target_h, max_w):
    """把单元格图片读出来并按目标高度缩放(超宽则反过来按 max_w 限制)。失败返回 None。"""
    try:
        im = Image.open(path)
        im.load()
    except Exception:
        return None
    if im.mode not in ("RGB", "RGBA"):
        im = im.convert("RGBA" if "A" in im.getbands() else "RGB")
    ow, oh = im.size
    if ow <= 0 or oh <= 0:
        return None
    th = max(1, int(target_h))
    tw = max(1, int(round(ow * th / oh)))
    if tw > max_w:
        tw = max(1, int(max_w))
        th = max(1, int(round(oh * tw / ow)))
    try:
        return im.resize((tw, th), Image.LANCZOS)
    except Exception:
        return None


def render_table(spec, out_path, missing=None):
    """把通用网格 spec 渲染成 PNG。返回 (out_path, (w, h))。
    missing 是列表时，读不出的单元格图片（该格按无图留空）逐个记进去，交调用方告警。"""
    cells = spec.get("cells") or []
    ncols = int(spec.get("ncols") or 0)
    nrows = int(spec.get("nrows") or 0)
    header_rows = max(0, int(spec.get("header_rows") or 0))
    if ncols <= 0 or nrows <= 0 or not cells:
        raise ValueError("empty table spec")

    s = SS
    f_body = _font(FONT_REGULAR, BODY_FONT_PX * s)
    f_bold = _font(FONT_BOLD, BODY_FONT_PX * s)
    f_head = _font(FONT_BOLD, HEADER_FONT_PX * s)

    pad_x = CELL_PAD_X * s
    pad_y = ROW_PAD_Y * s
    line_gap = LINE_GAP * s
    min_col_w = MIN_COL_W * s
    max_col_w = MAX_COL_W * s
    thumb_h = THUMB_H * s

    probe = ImageDraw.Draw(Image.new("RGB", (4, 4)))

    def tw(text, font):
        return probe.textlength(text, font=font)

    def font_for(cell):
        return f_head if cell["_header"] else (f_bold if cell.get("bold") else f_body)

    # 预处理：标注表头区；加载图片单元格的缩略图(失败则当无图)
    for cell in cells:
        cell["_header"] = cell["r"] < header_rows
        cell["lines"] = [ln for ln in (cell.get("lines") or []) if ln != ""]
        imgs = []
        for path in (cell.get("images") or []):
            thumb = _load_thumb(path, thumb_h, max_col_w)
            if thumb is not None:
                imgs.append(thumb)
            elif missing is not None:
                missing.append(path)
        cell["_imgs"] = imgs

    def cell_natural_w(cell):
        # 图文同格取「图片宽」与「文字宽」的较大者，避免窄缩略图把长文字挤成逐字折行
        w = 0.0
        if cell["_imgs"]:
            w = max(im.width for im in cell["_imgs"])
        if cell["lines"]:
            font = font_for(cell)
            w = max(w, min(max((tw(ln, font) for ln in cell["lines"]), default=0.0), max_col_w))
        return w + 2 * pad_x

    # ---------- 1) 列宽：先按单列单元格的自然宽度，再保证跨列单元格放得下 ----------
    col_w = [min_col_w] * ncols
    for cell in cells:
        if cell["cs"] != 1:
            continue
        col_w[cell["c"]] = max(col_w[cell["c"]], cell_natural_w(cell))

    for cell in cells:
        if cell["cs"] <= 1:
            continue
        need = cell_natural_w(cell)
        span = range(cell["c"], cell["c"] + cell["cs"])
        have = sum(col_w[c] for c in span)
        if need > have:
            add = (need - have) / cell["cs"]
            for c in span:
                col_w[c] += add

    col_w = [int(round(w)) for w in col_w]
    # 超宽表：按比例压缩列宽再折行(在折行之前压，文字会重新换行而不是被截断)。
    # 每列不压到 min_col_w 以下，避免数字/英文被逐字拆得难读；列实在太多时宁可整体偏宽
    # (展示端 <img max-width:100%> 会再整体缩放)，也不把单列压成几像素。
    max_total = MAX_TABLE_W * s
    if sum(col_w) > max_total and sum(col_w) > 0:
        factor = max_total / sum(col_w)
        col_w = [max(int(min_col_w), int(round(w * factor))) for w in col_w]
    col_x = [0]
    for w in col_w:
        col_x.append(col_x[-1] + w)
    total_w = col_x[-1]

    # 列宽可能因整表限宽被二次压缩。图片必须以最终单元格内宽为准再次缩小，
    # 否则居中时会越过单元格边界，覆盖相邻列。
    for cell in cells:
        inner_w = sum(col_w[cell["c"]:cell["c"] + cell["cs"]]) - 2 * pad_x
        fitted = []
        for im in cell["_imgs"]:
            if im.width > inner_w:
                fit_w = max(1, int(inner_w))
                fit_h = max(1, int(round(im.height * fit_w / im.width)))
                im = im.resize((fit_w, fit_h), Image.LANCZOS)
            fitted.append(im)
        cell["_imgs"] = fitted

    # ---------- 2) 文字换行：按各单元格分到的宽度逐字折行 ----------
    def wrap(lines, font, max_w):
        out = []
        for seg in lines:
            seg = seg.strip()
            if not seg:
                continue
            cur = ""
            for ch in seg:
                if cur and tw(cur + ch, font) > max_w:
                    out.append(cur)
                    cur = ch
                else:
                    cur += ch
            if cur:
                out.append(cur)
        return out or [""]

    for cell in cells:
        # 图片格也照常折行文字：同格既有图又有字时，缩略图在上、文字在下，文字不丢。
        if cell["lines"]:
            font = font_for(cell)
            inner_w = sum(col_w[cell["c"]:cell["c"] + cell["cs"]]) - 2 * pad_x
            cell["_wrapped"] = wrap(cell["lines"], font, max(inner_w, 1))
        else:
            cell["_wrapped"] = []

    # ---------- 3) 行高：图片块(若有) + 文字块(若有)，两者都有时中间留一个行距 ----------
    def img_block_h(cell):
        if not cell["_imgs"]:
            return 0
        n = len(cell["_imgs"])
        return sum(im.height for im in cell["_imgs"]) + (n - 1) * line_gap

    def text_block_h(cell):
        n = len(cell["_wrapped"])
        if n == 0:
            return 0
        return n * font_for(cell).size + (n - 1) * line_gap

    def cell_content_h(cell):
        ih = img_block_h(cell)
        th = text_block_h(cell)
        gap = line_gap if (ih and th) else 0
        return ih + gap + th + 2 * pad_y

    row_h = [int(BODY_FONT_PX * s + 2 * pad_y)] * nrows
    for cell in cells:
        if cell["rs"] != 1:
            continue
        row_h[cell["r"]] = max(row_h[cell["r"]], cell_content_h(cell))
    for cell in cells:
        if cell["rs"] <= 1:
            continue
        span = range(cell["r"], cell["r"] + cell["rs"])
        have = sum(row_h[r] for r in span)
        need = cell_content_h(cell)
        if need > have:
            add = int(round((need - have) / cell["rs"]))
            for r in span:
                row_h[r] += add

    row_y = [0]
    for h in row_h:
        row_y.append(row_y[-1] + h)
    total_h = row_y[-1]

    # OOM 兜底：极端表(超多行/列)像素量过大时直接放弃，让主程序退回文字兜底，别把内存撑爆
    if total_w * total_h > MAX_TABLE_PX:
        raise ValueError(f"table too large to render: {total_w}x{total_h}px")

    # ---------- 4) 绘制 ----------
    img = Image.new("RGB", (total_w, total_h), (255, 255, 255))
    d = ImageDraw.Draw(img)

    def draw_text_lines(x0, x1, ty, lines, font, color, align):
        line_step = font.size + line_gap
        for ln in lines:
            lw = tw(ln, font)
            if align == "r":
                tx = x1 - pad_x - lw
            elif align == "l":
                tx = x0 + pad_x
            else:
                tx = x0 + (x1 - x0 - lw) / 2
            d.text((tx, ty), ln, font=font, fill=color)
            ty += line_step

    for cell in cells:
        x0 = col_x[cell["c"]]
        x1 = col_x[cell["c"] + cell["cs"]]
        y0 = row_y[cell["r"]]
        y1 = row_y[cell["r"] + cell["rs"]]
        is_header = cell["_header"]
        if is_header:
            bg = C_HEAD_BG
            fg = C_HEAD_FG
        else:
            data_row = cell["r"] - header_rows
            bg = C_ROW_A if data_row % 2 == 0 else C_ROW_B
            fg = C_BODY_FG
        d.rectangle([x0, y0, x1, y1], fill=bg)

        align = cell.get("align")
        if align not in ("l", "c", "r"):
            if is_header:
                align = "c"
            elif len(cell["_wrapped"]) > 1:
                align = "l"
            else:
                align = "c"

        # 图片块(若有)在上、文字块(若有)在下，整体在单元格内垂直居中
        imgs = cell["_imgs"]
        lines = cell["_wrapped"]
        ih = img_block_h(cell)
        th = text_block_h(cell)
        gap = line_gap if (ih and th) else 0
        cy = y0 + (y1 - y0 - (ih + gap + th)) / 2
        for im in imgs:
            ix = int(x0 + (x1 - x0 - im.width) / 2)
            if im.mode == "RGBA":
                img.paste(im, (ix, int(cy)), im)
            else:
                img.paste(im, (ix, int(cy)))
            cy += im.height + line_gap
        # 图片循环每张后都加了一个 line_gap，正好充当图文之间的间隔；文字从 cy 起绘。
        if lines:
            draw_text_lines(x0, x1, cy, lines, font_for(cell), fg, align)

    # 边框：默认只画横线(行分隔)，不画竖网格线。逐个 origin 单元格画其上/下边，
    # 因此纵向合并(rowspan)的单元格中间不会被横线穿过；横向相邻单元格共享同一条横线。
    bw = max(1, s // 2)
    for cell in cells:
        x0 = col_x[cell["c"]]
        x1 = col_x[cell["c"] + cell["cs"]]
        y0 = row_y[cell["r"]]
        y1 = row_y[cell["r"] + cell["rs"]]
        if SHOW_VGRID:
            d.rectangle([x0, y0, x1 - 1, y1 - 1], outline=C_BORDER, width=bw)
        else:
            d.line([x0, y0, x1, y0], fill=C_BORDER, width=bw)          # 顶边(横)
            d.line([x0, y1 - 1, x1, y1 - 1], fill=C_BORDER, width=bw)  # 底边(横)

    # 表头下分隔线(更显眼) + 上下外框横线
    if 0 < header_rows < nrows:
        hy = row_y[header_rows]
        d.line([0, hy, total_w, hy], fill=C_FRAME, width=bw)
    if SHOW_VGRID:
        d.rectangle([0, 0, total_w - 1, total_h - 1], outline=C_FRAME, width=bw)
    else:
        d.line([0, 0, total_w, 0], fill=C_FRAME, width=bw)
        d.line([0, total_h - 1, total_w, total_h - 1], fill=C_FRAME, width=bw)

    if s != 1:
        img = img.resize((max(1, total_w // s), max(1, total_h // s)), Image.LANCZOS)
    img.save(out_path, "PNG")
    return out_path, img.size


if __name__ == "__main__":
    # 自测：拼一个带横向/纵向合并的小表，渲染出 PNG
    sys.stdout.reconfigure(encoding="utf-8")
    demo = {
        "ncols": 5,
        "nrows": 4,
        "header_rows": 2,
        "cells": [
            {"r": 0, "c": 0, "rs": 2, "cs": 1, "lines": ["名称"], "align": None, "bold": True},
            {"r": 0, "c": 1, "rs": 1, "cs": 2, "lines": ["归母净利润"], "align": None, "bold": True},
            {"r": 0, "c": 3, "rs": 1, "cs": 2, "lines": ["P/E"], "align": None, "bold": True},
            {"r": 1, "c": 1, "rs": 1, "cs": 1, "lines": ["2025A"], "align": None, "bold": False},
            {"r": 1, "c": 2, "rs": 1, "cs": 1, "lines": ["2026E"], "align": None, "bold": False},
            {"r": 1, "c": 3, "rs": 1, "cs": 1, "lines": ["2025A"], "align": None, "bold": False},
            {"r": 1, "c": 4, "rs": 1, "cs": 1, "lines": ["2026E"], "align": None, "bold": False},
            {"r": 2, "c": 0, "rs": 1, "cs": 1, "lines": ["豪悦护理"], "align": None, "bold": False},
            {"r": 2, "c": 1, "rs": 1, "cs": 1, "lines": ["2.3"], "align": None, "bold": False},
            {"r": 2, "c": 2, "rs": 1, "cs": 1, "lines": ["3.0"], "align": None, "bold": False},
            {"r": 2, "c": 3, "rs": 1, "cs": 1, "lines": ["20.9"], "align": None, "bold": False},
            {"r": 2, "c": 4, "rs": 1, "cs": 1, "lines": ["16.3"], "align": None, "bold": False},
            {"r": 3, "c": 0, "rs": 1, "cs": 1, "lines": ["可比公司平均"], "align": None, "bold": False},
            {"r": 3, "c": 1, "rs": 1, "cs": 1, "lines": ["-"], "align": None, "bold": False},
            {"r": 3, "c": 2, "rs": 1, "cs": 1, "lines": ["-"], "align": None, "bold": False},
            {"r": 3, "c": 3, "rs": 1, "cs": 1, "lines": ["29.8"], "align": None, "bold": False},
            {"r": 3, "c": 4, "rs": 1, "cs": 1, "lines": ["17.0"], "align": None, "bold": False},
        ],
    }
    path, size = render_table(demo, "table_demo.png")
    print(f"[OK] {path}  {size[0]}x{size[1]}px")
