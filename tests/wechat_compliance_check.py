"""离线检查生成的文章 HTML 是否符合公众号《内容排版规范》中可静态判定的条目。

用法：python tests/wechat_compliance_check.py 文件.html [更多文件...]（可写通配符；剪贴板片段可直接检查）
退出码：0 无违规；1 有违规；2 参数或读文件出错。

规则出处与编号见 docs/wechat-mp-official/checklist.md（C1、C2…）。每份 HTML 查两种视图：

- 文件视图：成品 HTML 里写明的内联样式。
- 复制视图：结果页「复制正文」把 <body> 换成 <section>，再把每个元素的计算
  font-family、text-align 等写进内联样式（src/web/result.html 的 copyFragment）。
  这里按继承规则静态推算这两项：没有任何祖先声明时，浏览器写入的是 text-align:start
  和浏览器默认字体，两者都会被编辑器判为违规。

需要真实浏览器排版才能判定的条目（多屏宽度差异、叠字、高度溢出、深色模式对比度）
只给 [WARN] 或 [NOTE] 提示，以官方 CLI 为准。只用标准库；装了 Pillow 时额外检查透明 PNG。
"""
from __future__ import annotations

import base64
import glob
import re
import sys
from html.parser import HTMLParser
from pathlib import Path

OFFICIAL_FONT_PREFIX = "mp-quote, PingFang SC, system-ui, -apple-system"
ALIGN_OK = {"left", "center", "right", "justify"}
VOID = {"br", "img", "meta", "hr", "input", "link", "source", "wbr"}
BLOCK_TAGS = {"section", "div", "p", "ul", "ol", "li", "table", "pre", "blockquote",
              "h1", "h2", "h3", "h4", "h5", "h6", "figure", "header", "footer"}
SHELL_TAGS = {"html", "head", "meta", "title", "body"}
ARTICLE_TAGS = {"section", "p", "br", "span", "strong", "b", "em", "i", "u", "ul", "ol", "li", "img"}
URL_ATTRS = {"src", "href", "background", "poster", "xlink:href"}
NESTING_LIMIT = 10


def normalize_font_family(value: str) -> str:
    """与官方检测引擎同口径：去引号、合并空白。"""
    return re.sub(r"\s+", " ", re.sub(r"[\"']", "", value or "")).strip()


def parse_style(style: str) -> dict:
    out = {}
    for decl in (style or "").split(";"):
        if ":" in decl:
            name, value = decl.split(":", 1)
            name = name.strip().lower()
            if name:
                out[name] = value.strip()
    return out


class Node:
    def __init__(self, tag, attrs, parent):
        self.tag = tag
        self.attrs = {k.lower(): (v if v is not None else "") for k, v in attrs}
        self.parent = parent
        self.children = []
        self.style = parse_style(self.attrs.get("style", ""))

    def elements(self):
        return [c for c in self.children if isinstance(c, Node)]

    def walk(self):
        for child in self.elements():
            yield child
            yield from child.walk()

    def text(self):
        parts = []
        for child in self.children:
            parts.append(child if isinstance(child, str) else child.text())
        return "".join(parts)

    def has_text(self):
        return bool(re.sub(r"[\s​\xa0]+", "", self.text()))

    def own_text(self):
        return any(isinstance(c, str) and re.sub(r"[\s​\xa0]+", "", c) for c in self.children)

    def label(self):
        snippet = re.sub(r"\s+", " ", self.text()).strip()[:24]
        return f"<{self.tag}>" + (f"「{snippet}」" if snippet else "")


class Tree(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node("#root", [], None)
        self.cur = self.root
        self.raw_tags = []

    def handle_starttag(self, tag, attrs):
        node = Node(tag, attrs, self.cur)
        self.cur.children.append(node)
        self.raw_tags.append(tag)
        if tag not in VOID:
            self.cur = node

    def handle_startendtag(self, tag, attrs):
        node = Node(tag, attrs, self.cur)
        self.cur.children.append(node)
        self.raw_tags.append(tag)

    def handle_endtag(self, tag):
        n = self.cur
        while n is not self.root and n.tag != tag:
            n = n.parent
        if n is not self.root:
            self.cur = n.parent

    def handle_data(self, data):
        self.cur.children.append(data)


class Report:
    """同一规则、同一问题的命中合并成一行，附命中数和前三处位置。"""

    def __init__(self):
        self._hits = {"FAIL": {}, "WARN": {}, "NOTE": {}}

    def _add(self, level, rule, msg, where):
        self._hits[level].setdefault((rule, msg), []).append(where)

    def fail(self, rule, msg, where=""):
        self._add("FAIL", rule, msg, where)

    def warn(self, rule, msg, where=""):
        self._add("WARN", rule, msg, where)

    def note(self, rule, msg, where=""):
        self._add("NOTE", rule, msg, where)

    def lines(self, level):
        out = []
        for (rule, msg), places in self._hits[level].items():
            places = [w for w in places if w]
            text = f"{rule} {msg}"
            if places:
                more = f" 等 {len(places)} 处" if len(places) > 3 else ""
                text += f"：{'；'.join(places[:3])}{more}"
            out.append(text)
        return out

    @property
    def fails(self):
        return self.lines("FAIL")

    @property
    def warns(self):
        return self.lines("WARN")

    @property
    def notes(self):
        return self.lines("NOTE")

    def rules(self, level="FAIL"):
        return {rule for rule, _ in self._hits[level]}


def _px(value, font_px):
    """把长度换成 px；em/% 以 font_px 为基准；无单位数字按倍数（line-height 用）。"""
    m = re.fullmatch(r"(-?[\d.]+)(px|em|%|pt)?", (value or "").strip())
    if not m:
        return None
    num, unit = float(m.group(1)), m.group(2)
    if unit == "px":
        return num
    if unit == "pt":
        return num * 4 / 3
    if unit == "%":
        return num * font_px / 100
    return num * font_px


def _inherited(node, prop):
    while node is not None and node.tag != "#root":
        if prop in node.style:
            return node.style[prop]
        node = node.parent
    return None


def _font_px(node):
    value = _inherited(node, "font-size")
    m = re.fullmatch(r"([\d.]+)px", (value or "").strip())
    return float(m.group(1)) if m else 16.0


def _png_alpha(data: bytes):
    """返回 True/False 表示是否有透明像素；无法判断返回 None。"""
    if not data.startswith(b"\x89PNG"):
        return False
    try:
        from io import BytesIO
        from PIL import Image
        with Image.open(BytesIO(data)) as image:
            if not (image.mode in ("RGBA", "LA", "PA") or "transparency" in image.info):
                return False
            return image.convert("RGBA").getchannel("A").getextrema()[0] < 255
    except ImportError:
        color_type = data[25] if len(data) > 25 else 0
        return True if color_type in (4, 6) or b"tRNS" in data[:4096] else False
    except Exception:
        return None


def audit(html_text: str) -> Report:
    report = Report()
    tree = Tree()
    tree.feed(html_text)
    nodes = list(tree.root.walk())
    body = next((n for n in nodes if n.tag == "body"), None)
    # 没有 <body> 即剪贴板片段（「复制正文」写出的 HTML）：计算样式已逐元素写明，
    # 复制视图直接按片段自身检查，不再模拟 <body>→<section>
    fragment = body is None
    if fragment:
        body = tree.root
    article = list(body.walk())

    # ---- 标签、属性、资源（C13~C16）----
    for n in nodes:
        if n.tag in ("script", "style", "link", "iframe", "object", "embed", "form", "input"):
            report.fail("C13", f"出现 <{n.tag}>：编辑器不执行脚本、不保留样式表与嵌入对象")
        for name, value in n.attrs.items():
            if name.startswith("on"):
                report.fail("C13", "带事件属性", f"<{n.tag}> {name}")
            urls = [value] if name in URL_ATTRS else []
            if name == "style":
                urls = re.findall(r"url\(\s*['\"]?([^'\")]+)", value)
            for url in urls:
                if not re.match(r"\s*(data|https):", url or "", re.I):
                    report.fail("C14", f"<{n.tag}> {name} 资源不是 data: 或 https:（{(url or '')[:60]}）")
    for n in article:
        if n.tag not in ARTICLE_TAGS:
            report.fail("C15", f"正文出现白名单外标签（白名单：{' '.join(sorted(ARTICLE_TAGS))}）", f"<{n.tag}>")
        if "id" in n.attrs:
            report.note("C15", "带 id，编辑器可能丢弃", f"<{n.tag} id={n.attrs['id']!r}>")
        if "class" in n.attrs and n.attrs["class"] != "list-paddingleft-2":
            report.note("C15", "带 class，样式不会随粘贴保留", f"<{n.tag} class={n.attrs['class']!r}>")

    # ---- 1、CSS 属性 ----
    for n in article:
        st = n.style
        raw = (n.attrs.get("style") or "").lower()
        if "!important" in raw:
            report.fail("C11", "使用 !important（规范 4.5.2）", n.label())
        if n.tag == "img" and re.match(r"0(\.0+)?$", st.get("opacity", "1")):
            report.fail("C1", "opacity 为 0（规范 1.1）", n.label())
        if re.sub(r"\s", "", st.get("caret-color", "")) in ("rgba(0,0,0,0)", "transparent"):
            report.fail("C2", "caret-color 透明（规范 1.2）", n.label())
        if re.fullmatch(r"0(px)?", st.get("height", "").strip()) and n.has_text():
            report.fail("C6", "height:0 且含文字（规范 1.5.1）", n.label())
        ta = st.get("text-align", "").strip().lower()
        if ta in ("start", "end"):
            report.fail("C7", f"text-align:{ta}（规范 1.6）", n.label())
        ff = st.get("font-family")
        if ff is not None and not normalize_font_family(ff).startswith(OFFICIAL_FONT_PREFIX):
            report.fail("C10", f"font-family:{ff}（规范 3，只放行平台默认字体栈）", n.label())
        if "gradient(" in raw and n.has_text():
            report.fail("C12", "文字背景使用渐变（规范 4.1.2）", n.label())
        if st.get("position", "").strip().lower() in ("absolute", "fixed") or "transform" in st:
            report.fail("C12", "使用绝对定位或变换，打乱深色模式遍历顺序（规范 4.2.2）", n.label())
        lh = st.get("line-height")
        if lh is not None and n.has_text():
            fs = _font_px(n)
            lh_px = _px(lh, fs)
            if lh_px is not None and lh_px < fs * 0.95:
                report.fail("C3", "line-height 小于字号（规范 1.3；单行文字不算违规，需在浏览器确认是否折行）",
                            f"{n.label()} {lh}/{fs:g}px")
        width = st.get("width", "").strip().lower()
        if (n.tag not in ("img", "svg") and width and not re.search(r"auto|100%", width)
                and "data-ignore-width" not in n.attrs):
            report.warn("C4", f"固定宽度 width:{width}（规范 1.4，需三种屏宽实测）", n.label())
        margin_left = _px(st.get("margin-left", ""), _font_px(n))
        if margin_left is not None and margin_left > 120:
            report.warn("C4", f"margin-left:{st['margin-left']} 可能在窄屏溢出（规范 1.4.2）", n.label())
        height = st.get("height", "").strip().lower()
        if (height and height != "auto" and not re.fullmatch(r"0(px)?", height)
                and st.get("overflow", st.get("overflow-y", "")).strip().lower() in ("hidden", "clip")
                and n.has_text()):
            report.warn("C6", f"固定高度 {height} 且裁剪溢出，文字折行会被裁掉（规范 1.5.2）", n.label())
        for attr in ("data-no-dark", "data-ignore-dm", "data-ignore-width"):
            if attr in n.attrs:
                report.note("C12", f"使用豁免属性 {attr}，确认是有意为之", n.label())

    # ---- <pre>、SVG 动画、leaf 节点（C8、C9、C17）----
    for n in nodes:
        if n.tag == "pre" and n.has_text():
            report.fail("C8", "出现 <pre>：窄屏不换行会被截断（规范 1.8）；代码块用 section+p+br")
        if n.tag == "animate" and "touchstart" in n.attrs.get("begin", "") and "click" not in n.attrs.get("begin", ""):
            report.fail("C17", "SVG 动画 begin 只写了 touchstart（规范 1.7）")
        if n.tag == "svg":
            report.fail("C17", "正文内嵌 SVG：深色模式不转换，固定颜色可能看不见（规范 4.4.1）")
        if n.tag == "span" and "leaf" in n.attrs and any(c.tag in BLOCK_TAGS for c in n.walk()):
            report.fail("C9", "span[leaf] 里有块级元素（规范 2.2）")
        if n.tag == "section" and "nodeleaf" in n.attrs and any(c.tag != "img" for c in n.elements()):
            report.fail("C9", "section[nodeleaf] 里有 img 以外的元素（规范 2.3）")

    # ---- 2.1 冗余嵌套：同标签、同样式、单子节点的连续链 ----
    for n in article:
        chain = 1
        cur = n
        while True:
            kids = cur.elements()
            if len(kids) != 1 or cur.own_text():
                break
            kid = kids[0]
            if kid.tag != n.tag or kid.attrs.get("style", "") != n.attrs.get("style", "") or kid.tag == "img":
                break
            chain += 1
            cur = kid
        if chain > NESTING_LIMIT:
            report.fail("C9", f"同样式单子节点连续嵌套超过 {NESTING_LIMIT} 层（规范 2.1）", f"<{n.tag}> {chain} 层")
            break

    # ---- 复制视图：text-align、font-family 的计算值 ----
    for n in ([] if fragment else [body]) + article:
        if n.tag == "br":
            continue
        where = "<body→section>" if n is body else n.label()
        ta = (_inherited(n, "text-align") or "start").strip().lower()
        if ta not in ALIGN_OK:
            report.fail("C7", f"复制正文后计算对齐为 {ta}（规范 1.6）", where)
        if n.has_text() or n is body:
            ff = _inherited(n, "font-family")
            if ff is None:
                report.fail("C10", "复制正文后写入浏览器默认字体（规范 3）", where)
            elif not normalize_font_family(ff).startswith(OFFICIAL_FONT_PREFIX):
                report.fail("C10", "复制正文后计算字体不是平台默认字体栈（规范 3）", where)

    # ---- 图片（C5、C14、C16）----
    images = [n for n in article if n.tag == "img"]
    total = 0
    largest = 0
    for i, img in enumerate(images, 1):
        src = img.attrs.get("src", "")
        if not src.startswith("data:"):
            report.fail("C14", f"图片没有以 data: 内嵌（{src[:60]}）", f"图片 {i}")
            continue
        m = re.match(r"data:([\w/+.-]+);base64,(.*)", src, re.S)
        if not m:
            report.fail("C14", "图片的 data: URI 不是 base64", f"图片 {i}")
            continue
        mime = m.group(1).lower()
        try:
            data = base64.b64decode(m.group(2))
        except Exception:
            report.fail("C14", "图片的 base64 无法解码", f"图片 {i}")
            continue
        total += len(data)
        largest = max(largest, len(data))
        if mime not in ("image/png", "image/jpeg", "image/gif"):
            report.warn("C16", f"图片格式 {mime}，编辑器转存只认 png/jpeg/gif", f"图片 {i}")
        if mime == "image/png" and _png_alpha(data):
            report.fail("C16", "图片是透明底 PNG，深色模式底色 #191919 上的深色内容看不见（规范 4.3.2）", f"图片 {i}")
        if "data-w" not in img.attrs:
            report.note("C5", "图片没有 data-w（规范 1.4.3 建议项；data: 图片即时加载，不影响检测）", f"图片 {i}")
    for n in article:
        bg = n.style.get("background-image", "")
        if "url(data:" in bg.replace(" ", ""):
            report.warn("C16", "背景图是 data: URI；粘贴后编辑器是否转存背景图未经验证", n.label())
    if images:
        report.note("C16", f"{len(images)} 张图片，内嵌 {total / 1024:.0f} KB，最大 {largest / 1024:.0f} KB")
    return report


def main(argv):
    if not argv:
        print(__doc__)
        return 2
    worst = 0
    names = []
    for name in argv:
        # PowerShell 不替原生程序展开通配符，这里自行展开；无匹配按读文件出错处理，避免门禁漏查
        if glob.has_magic(name):
            matched = sorted(glob.glob(name))
            if not matched:
                print(f"[ERROR] {name}: 通配符没有匹配到文件")
                return 2
            names += matched
        else:
            names.append(name)
    for name in names:
        path = Path(name)
        try:
            text = path.read_bytes().decode("utf-8-sig")
        except OSError as error:
            print(f"[ERROR] {path}: {error}")
            return 2
        report = audit(text)
        print(f"== {path}")
        for line in report.notes:
            print("[NOTE]", line)
        for line in report.warns:
            print("[WARN]", line)
        for line in report.fails:
            print("[FAIL]", line)
        print(f"[{'FAIL' if report.fails else 'PASS'}] {len(report.fails)} 项违规，{len(report.warns)} 项需实测")
        worst = max(worst, 1 if report.fails else 0)
    return worst


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main(sys.argv[1:]))
