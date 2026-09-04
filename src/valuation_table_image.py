# -*- coding: utf-8 -*-
"""
估值表 -> 图片 (PNG)

把周报 .docx 里的"行业公司估值表"渲染成一张清晰图片，方便直接插入公众号文章，
手机/电脑都能看。不再需要人工截图。

特性
- 自适应：每个板块标的增减、总行数变化都自动处理（按内容逐行识别，不依赖固定行号）。
- 年份自适应：表头年份(2025A/2026E/2027E)从文档里读，跨年不用改代码。
- 注释/数据来源字体比正文小 ~20%（NOTE_FONT_SCALE 可调）。
- 2 倍超采样 + LANCZOS 缩放，文字锐利。

用法
    python valuation_table_image.py "行业周报XXXX.docx"
    python valuation_table_image.py "报告.docx" -o 估值表.png
可一次拖多个文件。输出默认与 docx 同目录，文件名 *_估值表.png
"""

import os
import re
import argparse

import docx
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont


# ----------------------------- 可调样式 -----------------------------
FONT_REGULAR = r"C:\Windows\Fonts\msyh.ttc"      # 微软雅黑
FONT_BOLD    = r"C:\Windows\Fonts\msyhbd.ttc"    # 微软雅黑 粗体

BODY_FONT_PX   = 15      # 正文(数据)字号(逻辑像素)
TITLE_FONT_PX  = 19      # 标题字号
HEADER_FONT_PX = 15      # 表头字号
NOTE_FONT_SCALE = 0.8    # 注释/数据来源字体 = 正文 * 0.8  (即小 20%)

SS = 2                   # 超采样倍数(清晰度)

ROW_PAD_Y   = 7          # 单元格上下内边距
CELL_PAD_X  = 10         # 单元格左右内边距
MIN_COL_W   = 42         # 列最小宽度

# 配色 (东吴蓝系)
C_TITLE_BG   = (31, 78, 121)     # 标题深蓝
C_TITLE_FG   = (255, 255, 255)
C_HEAD_BG    = (46, 90, 156)     # 表头蓝
C_HEAD_FG    = (255, 255, 255)
C_SECTION_BG = (218, 230, 244)   # 板块分隔浅蓝
C_SECTION_FG = (31, 78, 121)
C_ROW_A      = (255, 255, 255)   # 数据行 交替色 A
C_ROW_B      = (242, 246, 252)   # 数据行 交替色 B
C_BODY_FG    = (33, 37, 41)
C_BORDER     = (200, 210, 224)
C_NOTE_FG    = (90, 90, 90)
C_NOTE_BG    = (250, 251, 253)

# 评级文字颜色
RATING_COLORS = {
    "买入": (192, 57, 43), "增持": (211, 84, 0),
    "中性": (120, 120, 120), "减持": (39, 110, 60), "卖出": (39, 110, 60),
}
C_RATING_DEFAULT = (120, 120, 120)   # 未评级 / - 等
# -------------------------------------------------------------------


YEAR_RE = re.compile(r"^\s*\d{4}\s*[AEFae]\s*$")
CAPTION_RE = re.compile(r"^\s*表\s*(?:\d+|[一二三四五六七八九十]+)\s*[：:]")
# 备注/来源行判定的**独立 CLI 默认值**：与主 cleaner 的 source_regex 同口径(行首锚定 + 紧跟冒号)，
# 避免"资料来源说明"这类无冒号子串被误判成来源行而丢掉真实数据行。
# 主程序(cleaner)调用 parse_table 时会传入 self.re_source(源自 config.json)，自动同步、不漂移；
# 仅本模块当独立脚本(process/main)运行时才用到这个默认值——改 config 的 source_regex 时建议一并对齐这里。
NOTES_RE = re.compile(r"^\s*(?:数据来源|资料来源|资料及数据来源|[Ss]ources?)\s*[：:]")


def _font(path, px):
    return ImageFont.truetype(path, px)


_T = qn("w:t")
_BR = qn("w:br")
_CR = qn("w:cr")


def _para_text(p):
    """段落文本，软换行(w:br/w:cr)转为 \\n。"""
    buf = []
    for node in p.iter():
        if node.tag == _T:
            buf.append(node.text or "")
        elif node.tag in (_BR, _CR):
            buf.append("\n")
    return "".join(buf)


def _cell_text(tc):
    """单元格文本，保留换行(\\n)。备注单元格(数据来源/注1/2/3)靠它分行。"""
    paras = tc.findall(qn("w:p"))
    if paras:
        return "\n".join(_para_text(p) for p in paras).strip()
    return "".join(n.text or "" for n in tc.iter(_T)).strip()


def _expand_row(tr, ncols):
    """按 gridSpan 把一行展开成 ncols 个网格列。
    返回 (vals, primary)：
      vals[i]    第 i 个网格列的文字(合并单元格在其跨度内重复)
      primary[i] 该网格列是否为某单元格的"起始列"(True=真实逻辑列；
                 False=gridSpan 合并产生的续列，或没有任何单元格覆盖的空列)
    """
    vals = [""] * ncols
    primary = [False] * ncols
    ci = 0
    for tc in tr.findall(qn("w:tc")):
        span = 1
        tcPr = tc.find(qn("w:tcPr"))
        if tcPr is not None:
            gs = tcPr.find(qn("w:gridSpan"))
            if gs is not None:
                span = int(gs.get(qn("w:val")))
        txt = _cell_text(tc)
        for k in range(span):
            if ci < ncols:
                vals[ci] = txt
                primary[ci] = (k == 0)
                ci += 1
    return vals, primary


def content_signature(rows):
    def clean(value):
        value = re.sub(r"\[Table_\w+\]", "", value or "")
        value = re.sub(r"\s+", "", value)
        # 题注编号「表X：」来自 Word SEQ 域：HTML 导出渲染域结果，
        # python-docx 只读到域外文字，两侧必然不等，剥掉后再比对。
        return re.sub(r"^[表图]\s*(?:\d+|[一二三四五六七八九十]+)\s*[：:]", "", value)

    result = []
    for row in rows:
        cells = tuple(clean(value) for value in row if clean(value))
        if cells:
            result.append(cells)
    return tuple(result)


def table_content_signature(table):
    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        return ()
    ncols = len(grid.findall(qn("w:gridCol")))
    rows = []
    for tr in tbl.findall(qn("w:tr")):
        vals, primary = _expand_row(tr, ncols)
        rows.append([value for value, is_primary in zip(vals, primary) if is_primary])
    return content_signature(rows)


def find_valuation_table(doc, keyword="估值表", require_keyword=False, signature=None):
    """找估值表：关键词出现在前 3 行任意单元格者优先。
    require_keyword=False(独立 CLI 默认)：无命中时退化为行数最多的表。
    require_keyword=True(主程序调用)：仅在关键词与内容签名均唯一命中时返回；
    零张、多张或签名不一致均返回 None，交由通用数据表渲染。
    与主程序的 _is_valuation_table(同样查前 3 行)保持一致，避免两边选到不同的表。"""
    best = None
    matches = []
    for t in doc.tables:
        hit = any(keyword in c.text for r in t.rows[:3] for c in r.cells)
        if hit:
            matches.append(t)
        if best is None or len(t.rows) > len(best.rows):
            best = t
    if require_keyword:
        if signature is not None:
            matches = [table for table in matches
                       if table_content_signature(table) == signature]
        return matches[0] if len(matches) == 1 else None
    return matches[0] if matches else best


def parse_table(table, notes_re=None):
    """
    解析估值表 -> 结构化数据。完全按内容逐行识别，自适应板块行数变化。
    返回 dict: title, group_header(list), year_labels(list3),
               sections(list of (section_name, [data_row,...])), notes(list[str])
    其中 data_row 是 11 个逻辑列的列表。

    notes_re: 备注/来源行的判定正则(.search)。主程序(cleaner)传入自己的 self.re_source
    (源自 config.json 的 source_regex)，使两边口径自动同步、不漂移；独立 CLI 不传则用模块默认 NOTES_RE。
    """
    if notes_re is None:
        notes_re = NOTES_RE
    tbl = table._tbl
    ncols = len(tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol")))
    raw = [_expand_row(tr, ncols) for tr in tbl.findall(qn("w:tr"))]

    title = ""
    notes = []
    group_header = None        # vals
    year_header = None
    body_rows = []             # ("section", name) | ("data", vals, primary)

    for vals, primary in raw:
        nonempty = [c for c in vals if c != ""]
        if not nonempty:
            continue
        allsame = len(set(nonempty)) == 1
        c0 = vals[0] if vals else ""

        if (not title
                and (any(CAPTION_RE.match(c) for c in nonempty)
                     or any("估值表" in c for c in nonempty))):
            title = "".join(c for c, is_primary in zip(vals, primary)
                            if c and is_primary)
            continue
        if any(notes_re.search(c) for c in nonempty):
            note = "".join(c for c, is_primary in zip(vals, primary)
                           if c and is_primary)
            if note:
                notes.append(note)
            continue
        if c0 == "代码":                       # 组表头行
            group_header = vals
            continue
        if c0 == "" and any(YEAR_RE.match(c) for c in vals):   # 年份子表头行
            year_header = (vals, primary)
            continue
        if allsame:                            # 板块分隔行
            body_rows.append(("section", nonempty[0]))
            continue
        # 数据行
        body_rows.append(("data", vals, primary))

    # --- 用 gridSpan 结构(而非"数值相同")判断要保留哪些网格列 ---
    # 真实逻辑列 = 在任一数据行里是某单元格的起始列(primary)。
    # 合并产生的续列 / 没有单元格覆盖的空列 都不是 primary，丢弃。
    # 这样不会因"两列数值恰好相同"或"某列对所有公司恰好为空"而误删真实列。
    data_primaries = [item[2] for item in body_rows if item[0] == "data"]
    keep = [c for c in range(ncols)
            if any(c < len(p) and p[c] for p in data_primaries)]
    if not keep:
        keep = list(range(ncols))
    if len(keep) != 11:
        raise ValueError(
            f"valuation table has {len(keep)} logical columns; expected 11"
        )

    def project(vals):
        vals = list(vals) + [""] * (ncols - len(vals))
        return [vals[c] for c in keep]

    if not title:
        raise ValueError("valuation table caption not found")
    if group_header is None:
        raise ValueError("valuation table group header not found")
    if year_header is None:
        raise ValueError("valuation table year header not found")

    gh = project(group_header)
    required_header_cols = (0, 1, 2, 3, 4, 7, 10)
    if (any(not gh[c].strip() for c in required_header_cols)
            or len(set(gh[4:7])) != 1
            or len(set(gh[7:10])) != 1):
        raise ValueError("valuation table group header is incomplete or misaligned")

    # 年份行的合并单元格(如 P/E 下两两合并的年份格)起始列与数据行 primary 列
    # 不一定重合，不能按 keep 投影核对；在年份行自身的 primary 单元格上验证：
    # 全部是年份格式、恰为两组各 3 个且两组同序(归母净利润组与 P/E 组各一份)。
    # 必须是 6 个：只有一组年份时 column_layout 会把它复制到另一组，
    # 输出源表不存在的年份口径，违反"不编造"原则，因此直接解析失败回退。
    year_vals, year_primary = year_header
    year_cells = [c.strip() for c, is_primary in zip(year_vals, year_primary)
                  if is_primary and c.strip()]
    if (len(year_cells) != 6
            or not all(YEAR_RE.match(y) for y in year_cells)
            or year_cells[:3] != year_cells[3:]
            or len(set(year_cells[:3])) != 3):
        raise ValueError("valuation table year header is incomplete or misaligned")
    year_labels = year_cells[:3]

    # 组装 sections
    sections = []
    cur_name, cur_rows = None, []
    for item in body_rows:
        if item[0] == "section":
            if cur_name is not None or cur_rows:
                sections.append((cur_name, cur_rows))
            cur_name, cur_rows = item[1], []
        else:
            cur_rows.append(project(item[1]))
    if cur_name is not None or cur_rows:
        sections.append((cur_name, cur_rows))

    return {
        "title": title,
        "group_header": gh,
        "year_labels": year_labels,
        "sections": sections,
        "notes": notes,
        "ncols": len(keep),
    }


def _stack_unit(label):
    """把"名称（单位）"拆成上下两行：总市值（亿元） -> 总市值\\n（亿元）。"""
    label = label.replace("\n", "")
    m = re.match(r"^(.*?)([（(].*[）)])\s*$", label)
    if m and m.group(1):
        return m.group(1) + "\n" + m.group(2)
    return label


def _stack_label(label):
    """无括号的两字+两字标签居中折行：投资评级 -> 投资\\n评级。"""
    label = label.replace("\n", "")
    if len(label) == 4:
        return label[:2] + "\n" + label[2:]
    return label


# ---- 11 个逻辑列的显示属性 (顺序须与 keep 一致：代码,简称,市值,收盘,净利*3,PE*3,评级) ----
def column_layout(data):
    gh = data["group_header"]
    yrs = data["year_labels"]
    n = data["ncols"]
    # 期望 11 列；若结构异常则退化为通用居中
    if n == 11:
        # 表头文字直接取自文档，并把"名称（单位）"拆成两行，与原始 Word 表格一致
        c2, c3 = _stack_unit(gh[2]), _stack_unit(gh[3])
        rating = gh[10] if "\n" in gh[10] else _stack_label(gh[10])
        col_top = [gh[0], gh[1], c2, c3, gh[4], "", "", gh[7], "", "", rating]
        col_sub = ["", "", "", "", yrs[0], yrs[1], yrs[2], yrs[0], yrs[1], yrs[2], ""]
        # 全表居中，与原始 Word 表格一致
        align   = ["c"] * 11
        # 跨列分组: (起列, 跨几列, 顶部文字)
        spans = [(0, 1, gh[0]), (1, 1, gh[1]), (2, 1, c2), (3, 1, c3),
                 (4, 3, gh[4]), (7, 3, gh[7]), (10, 1, rating)]
        merged_top = {0, 1, 2, 3, 10}   # 这些列上下表头合并(纵向)
        return col_top, col_sub, align, spans, merged_top
    raise ValueError(f"valuation table has {n} logical columns; expected 11")


def rating_color(text):
    t = text.strip()
    return RATING_COLORS.get(t, C_RATING_DEFAULT)


def render(data, out_path, draw_title=False, draw_notes=False):
    """渲染估值表为图片。

    draw_title / draw_notes 默认 False：图片里只含表格本体(表头+数据)。
    标题与备注/数据来源交由外层(公众号 HTML)按图表说明的方式处理。
    """
    s = SS
    f_body  = _font(FONT_REGULAR, BODY_FONT_PX * s)
    f_bold  = _font(FONT_BOLD,    BODY_FONT_PX * s)
    f_head  = _font(FONT_BOLD,    HEADER_FONT_PX * s)
    f_title = _font(FONT_BOLD,    TITLE_FONT_PX * s)
    f_note  = _font(FONT_REGULAR, max(1, int(round(BODY_FONT_PX * NOTE_FONT_SCALE))) * s)

    col_top, col_sub, align, spans, merged_top = column_layout(data)
    ncol = data["ncols"]

    # 用于测量
    probe = Image.new("RGB", (10, 10))
    pdraw = ImageDraw.Draw(probe)

    def tw(text, font):
        # 多行文字取最宽一行
        return max((pdraw.textlength(ln, font=font) for ln in text.split("\n")), default=0.0)

    pad_x = CELL_PAD_X * s
    pad_y = ROW_PAD_Y * s

    # ---- 计算列宽 ----
    col_w = [MIN_COL_W * s] * ncol
    # 数据
    for _, rows in data["sections"]:
        for r in rows:
            for c in range(ncol):
                col_w[c] = max(col_w[c], tw(r[c], f_body) + 2 * pad_x)
    # 子表头年份
    for c in range(ncol):
        if col_sub[c]:
            col_w[c] = max(col_w[c], tw(col_sub[c], f_head) + 2 * pad_x)
    # 单列顶表头(未分组的列)
    for (start, span, text) in spans:
        if span == 1:
            col_w[start] = max(col_w[start], tw(text, f_head) + 2 * pad_x)
    # 分组顶表头需 >= 组内列宽之和
    for (start, span, text) in spans:
        if span > 1:
            need = tw(text, f_head) + 2 * pad_x
            have = sum(col_w[start:start + span])
            if need > have:
                add = (need - have) / span
                for c in range(start, start + span):
                    col_w[c] += add
    col_w = [int(round(w)) for w in col_w]
    total_w = sum(col_w)
    col_x = [0]
    for w in col_w:
        col_x.append(col_x[-1] + w)

    # ---- 行高 ----
    line_h = (BODY_FONT_PX * s) + 2 * pad_y
    body_line_step = f_body.size + int(2 * s)

    def body_row_h(values):
        max_lines = max((len((text or "").split("\n")) for text in values), default=1)
        return line_h + (max_lines - 1) * body_line_step
    title_h = (TITLE_FONT_PX * s) + 2 * pad_y + 4 * s
    head_h = (HEADER_FONT_PX * s) + 2 * pad_y
    head_total = head_h * 2

    # 估算 notes 高度(自动换行)
    note_font_px = max(1, int(round(BODY_FONT_PX * NOTE_FONT_SCALE))) * s
    note_line_h = note_font_px + int(6 * s)

    def wrap(text, font, max_w):
        lines = []
        for seg in text.split("\n"):
            seg = seg.strip()
            if not seg:
                lines.append("")
                continue
            cur = ""
            for ch in seg:
                if tw(cur + ch, font) <= max_w:
                    cur += ch
                else:
                    lines.append(cur)
                    cur = ch
            lines.append(cur)
        return lines

    notes = data["notes"]
    notes_text = "\n".join(notes) if isinstance(notes, (list, tuple)) else notes
    note_lines = wrap(notes_text, f_note, total_w - 2 * pad_x) if (draw_notes and notes_text) else []
    notes_h = (len(note_lines) * note_line_h + 2 * pad_y + int(4 * s)) if note_lines else 0
    if not draw_title:
        title_h = 0

    # 数据/分隔总高
    body_h = 0
    for name, rows in data["sections"]:
        if name:
            body_h += body_row_h([name])
        body_h += sum(body_row_h(r) for r in rows)

    total_h = title_h + head_total + body_h + notes_h

    # OOM 兜底：与 table_image.MAX_TABLE_PX 同门限，超大估值表放弃专用渲染，
    # 由调用方退回通用数据表转图，防止 Image.new 撑爆内存。
    if total_w * total_h > 60_000_000:
        raise ValueError(f"valuation table too large to render: {total_w}x{total_h}px")

    img = Image.new("RGB", (total_w, total_h), (255, 255, 255))
    d = ImageDraw.Draw(img)

    def cell_text(x0, x1, y0, y1, text, font, color, al):
        if not text:
            return
        lines = text.split("\n")
        line_step = font.size + int(2 * s)
        block_h = line_step * len(lines)
        ty = y0 + (y1 - y0 - block_h) / 2 - int(1 * s)
        for ln in lines:
            lw = tw(ln, font)
            if al == "c":
                tx = x0 + (x1 - x0 - lw) / 2
            elif al == "r":
                tx = x1 - pad_x - lw
            else:
                tx = x0 + pad_x
            d.text((tx, ty), ln, font=font, fill=color)
            ty += line_step

    y = 0
    # 标题(默认不画，交给外层公众号 HTML)
    if draw_title:
        d.rectangle([0, y, total_w, y + title_h], fill=C_TITLE_BG)
        cell_text(0, total_w, y, y + title_h, data["title"], f_title, C_TITLE_FG, "c")
        y += title_h

    # 表头(两行)
    head_y0 = y
    d.rectangle([0, head_y0, total_w, head_y0 + head_total], fill=C_HEAD_BG)
    for (start, span, text) in spans:
        x0 = col_x[start]
        x1 = col_x[start + span]
        if start in merged_top and span == 1:
            cell_text(x0, x1, head_y0, head_y0 + head_total, text, f_head, C_HEAD_FG, "c")
        else:
            cell_text(x0, x1, head_y0, head_y0 + head_h, text, f_head, C_HEAD_FG, "c")
    # 子表头(年份)
    for c in range(ncol):
        if col_sub[c]:
            cell_text(col_x[c], col_x[c + 1], head_y0 + head_h,
                      head_y0 + head_total, col_sub[c], f_head, C_HEAD_FG, "c")
    # 分组表头(归母净利润/P/E)与年份子表头之间：底部横线 + 年份之间竖线
    # 注意：竖线只在子表头一行内，不穿过上方合并的分组单元格
    for (start, span, _t) in spans:
        if span > 1:
            d.line([col_x[start], head_y0 + head_h, col_x[start + span], head_y0 + head_h],
                   fill=(120, 150, 200), width=max(1, s // 2))
            for c in range(start + 1, start + span):
                d.line([col_x[c], head_y0 + head_h, col_x[c], head_y0 + head_total],
                       fill=(120, 150, 200), width=max(1, s // 2))
    y += head_total

    # 数据 + 板块
    ri = 0
    for name, rows in data["sections"]:
        if name:
            row_h = body_row_h([name])
            d.rectangle([0, y, total_w, y + row_h], fill=C_SECTION_BG)
            cell_text(0, total_w, y, y + row_h, name, f_bold, C_SECTION_FG, "l")
            d.line([0, y, total_w, y], fill=C_BORDER, width=max(1, s // 2))
            y += row_h
            ri = 0
        for r in rows:
            row_h = body_row_h(r)
            bg = C_ROW_A if ri % 2 == 0 else C_ROW_B
            d.rectangle([0, y, total_w, y + row_h], fill=bg)
            for c in range(ncol):
                color = C_BODY_FG
                font = f_body
                if c == ncol - 1:        # 评级列上色
                    color = rating_color(r[c])
                    font = f_bold
                cell_text(col_x[c], col_x[c + 1], y, y + row_h, r[c], font, color, align[c])
            y += row_h
            ri += 1

    # 竖向网格线。合并的分组表头(归母净利润/P/E)内部不画竖线，
    # 其内部列边界只从年份子表头下方(数据区)开始画。
    internal = set()
    for (start, span, _t) in spans:
        if span > 1:
            for c in range(start + 1, start + span):
                internal.add(c)
    grid_bottom = y
    for c in range(1, ncol):
        top = (head_y0 + head_total) if c in internal else head_y0
        d.line([col_x[c], top, col_x[c], grid_bottom], fill=C_BORDER, width=max(1, s // 2))
    # 外框
    d.rectangle([0, 0, total_w - 1, total_h - 1], outline=C_BORDER, width=max(1, s // 2))
    d.line([0, head_y0 + head_total, total_w, head_y0 + head_total],
           fill=(120, 150, 200), width=max(1, s // 2))

    # 注释
    if note_lines:
        d.rectangle([0, y, total_w, y + notes_h], fill=C_NOTE_BG)
        ny = y + pad_y
        for ln in note_lines:
            d.text((pad_x, ny), ln, font=f_note, fill=C_NOTE_FG)
            ny += note_line_h

    # 超采样缩放
    if s != 1:
        img = img.resize((total_w // s, total_h // s), Image.LANCZOS)
    img.save(out_path, "PNG")
    return out_path, img.size


def process(docx_path, out_path=None):
    doc = docx.Document(docx_path)
    table = find_valuation_table(doc)
    if table is None:
        raise SystemExit("未找到任何表格")
    data = parse_table(table)
    if out_path is None:
        base = os.path.splitext(docx_path)[0]
        out_path = base + "_估值表.png"
    path, size = render(data, out_path)
    n_companies = sum(len(rows) for _, rows in data["sections"])
    print(f"[OK] {os.path.basename(docx_path)} -> {os.path.basename(path)}  "
          f"({size[0]}x{size[1]}px, {len(data['sections'])} 个板块 / {n_companies} 只标的)")
    return path


def main():
    ap = argparse.ArgumentParser(description="把周报docx里的估值表渲染成图片")
    ap.add_argument("docx", nargs="+", help="一个或多个 .docx 文件")
    ap.add_argument("-o", "--out", help="输出 png 路径(仅单文件时有效)")
    args = ap.parse_args()
    for p in args.docx:
        out = args.out if (args.out and len(args.docx) == 1) else None
        try:
            process(p, out)
        except Exception as e:
            print(f"[FAIL] {p}: {e}")


if __name__ == "__main__":
    main()
