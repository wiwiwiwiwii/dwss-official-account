from __future__ import annotations

import argparse
import base64
import json
import re
import sys
import tempfile
import traceback
import xml.etree.ElementTree as ElementTree
import zipfile
from pathlib import Path


import converter as engine


MODES = ("summary", "body", "weekly")
AUTO = "auto"
FRAMEWORK_MODES = {"weekly": ["weekly"], "comment": ["summary"], "depth": ["summary", "body"]}
KIND_LABELS = {"weekly": "周报", "comment": "点评", "depth": "深度"}
MODE_LABELS = {"summary": "摘要", "body": "正文", "weekly": "周报"}
TITLE_TAGS = ("Indicator_SecondTitle", "Indicator_ReportTitle")
CAPTION_RE = re.compile(r"^\s*(图|表)\s*(\d+)\s*[：:]")
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class Progress:
    """输出工作台结果页读取的 [STEP i/n] 行。"""

    def __init__(self) -> None:
        self.total = 0
        self.done = 0

    def plan(self, total: int) -> None:
        self.total = total

    def step(self, text: str) -> None:
        self.done += 1
        print(f"[STEP {self.done}/{max(self.total, self.done)}] {text}", flush=True)

    @staticmethod
    def note(text: str) -> None:
        print(text, flush=True)


def body_matching(root: Path) -> dict:
    return engine.load_config("body", root / "configs").get("matching", {})


def body_heading_styles(root: Path) -> list:
    names = []
    matching = body_matching(root)
    for level in (1, 2, 3):
        for name in matching.get(f"heading_level_{level}_style_names", []) or []:
            if str(name).strip():
                names.append(str(name).strip())
    return names


def classify(headers, titles, has_body_headings: bool):
    """返回（提取框架, 标题类型）。框架看文档结构，标题类型看封面页眉的措辞。

    带正文的点评存在：页眉写“公司点评报告”，正文却使用【正文】各级标题。
    这种报告要按深度框架提取，标题仍是点评。
    周报信号：页眉任意位置含“周报”，或标题控件以“周报”结尾（如“xx行业跟踪周报”）；
    标题控件正文中间偶然出现“周报”二字不算，避免误判整个提取框架。
    """
    header_text = re.sub(r"\s+", "", "".join(headers or ""))
    title_texts = [re.sub(r"\s+", "", t) for t in (titles or [])]
    text = header_text + "".join(title_texts)
    if "周报" in header_text or any(t.endswith("周报") for t in title_texts):
        return "weekly", "weekly"
    framework = "depth" if has_body_headings else "comment"
    if "点评" in text:
        return framework, "comment"
    if "深度" in text:
        return framework, "depth"
    return framework, framework


MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"


def _xml_texts(element) -> str:
    """可见文字。跳过 `mc:Fallback`：页眉里的文本框被存成 AlternateContent 时，
    Choice 与 Fallback 各存一份同样的文字，两份都算会把「万科A（000002）」
    读成「万科A（000002）万科A（000002）」，页眉兜底出来的公司名随即匹配不上三表标题。"""
    texts = []

    def walk(node):
        for child in node:
            if child.tag == MC_FALLBACK:
                continue
            if child.tag == "{%s}t" % WORD_NS:
                texts.append(child.text or "")
            walk(child)

    walk(element)
    return "".join(texts)


def docx_parts(source: Path):
    """返回（页眉文字列表, 首页标题控件文字列表）；不是可读的 .docx 时返回 None。

    页眉走包内 XML 而不是 Word COM：旧模板把页眉画在文本框里，
    COM 的 Headers(...).Range.Text 读到的是空串。
    """
    if source.suffix.lower() != ".docx":
        return None
    namespace = {"w": WORD_NS}
    try:
        with zipfile.ZipFile(str(source)) as archive:
            headers = [_xml_texts(ElementTree.fromstring(archive.read(name)))
                       for name in archive.namelist() if re.match(r"word/header\d*\.xml$", name)]
            document = ElementTree.fromstring(archive.read("word/document.xml"))
    except Exception:
        return None
    titles = []
    for control in document.findall(".//w:sdt", namespace):
        properties = control.find("w:sdtPr", namespace)
        tag = properties.find("w:tag", namespace) if properties is not None else None
        if tag is not None and tag.get("{%s}val" % WORD_NS) in TITLE_TAGS:
            titles.append(_xml_texts(control))
    return headers, titles





def header_company(source: Path) -> str:
    """旧模板缺副标题控件时，用页眉中“证券研究报告”之前的公司名兜底。"""
    parts = docx_parts(source)
    for text in (parts[0] if parts else []):
        cleaned = re.sub(r"\s+", " ", text).strip()
        lead = cleaned.split("证券研究报告")[0].strip()
        if lead and lead != cleaned:
            return lead
    return ""


def resolve_main_title(main_title: str, subtitle: str, source: Path) -> str:
    """公司名字段：优先用标题控件，旧模板缺该控件时退回页眉里的公司名。

    只在有副标题时兜底——没有副标题的报告本就退回裸报告标题，不拼造公司名。"""
    return main_title or (header_company(source) if subtitle else "")


def detect_kind_from_docx(root: Path, source: Path):
    parts = docx_parts(source)
    if parts is None:
        return None
    return classify(parts[0], parts[1], bool(engine.extract_docx_heading_outline(source, body_matching(root))))


def detect_kind_with_word(root: Path, word, source: Path) -> str:
    """.doc 走 Word：读页眉与标题控件，再用 Find 判断是否存在【正文】标题样式。"""
    document = word.open_document(source)
    try:
        headers_text, titles_text = [], []
        sections = document.Sections
        for index in range(1, min(int(sections.Count), 3) + 1):
            headers = sections(index).Headers
            for which in (1, 2, 3):
                try:
                    headers_text.append(str(headers(which).Range.Text))
                except Exception:
                    pass
        try:
            controls = document.ContentControls
            for index in range(1, int(controls.Count) + 1):
                control = controls(index)
                if str(control.Tag) in TITLE_TAGS:
                    titles_text.append(str(control.Range.Text))
        except Exception:
            pass
        found = False
        for name in body_heading_styles(root):
            try:
                finder = document.Content.Find
                finder.ClearFormatting()
                finder.Text = ""
                finder.Format = True
                finder.Style = name
                found = bool(finder.Execute())
            except Exception:
                found = False
            if found:
                break
        return classify(headers_text, titles_text, found)
    finally:
        try:
            document.Close(0)
        except Exception:
            pass


def apply_detection(detected, toc: bool, progress):
    """把识别结果落成本次转换的提取模式、目录开关和标题类型。"""
    framework, title_kind = detected
    modes = list(FRAMEWORK_MODES[framework])
    toc = toc and "body" in modes
    if title_kind == framework:
        progress.note(f"[AUTO] 报告类型：{KIND_LABELS[title_kind]}")
    else:
        progress.note(f"[AUTO] 报告类型：{KIND_LABELS[title_kind]}（正文按{KIND_LABELS[framework]}框架提取）")
    return modes, toc, title_kind


LEVEL_CAPTION_RE = re.compile(r"^\s*(图|表)\s*%\d\s*[：:]")


def _w(name: str) -> str:
    return "{%s}%s" % (WORD_NS, name)


def _w_val(element, path: str):
    found = element.find(path) if element is not None else None
    return found.get(_w("val")) if found is not None else None


def _is_toc_link_paragraph(paragraph) -> bool:
    """目录条目：带指向 `_Toc` 书签的超链接（内容目录、图表目录都是）。"""
    return any((link.get(_w("anchor")) or "").startswith("_Toc") for link in paragraph.iter(_w("hyperlink")))


def _is_toc_control(element) -> bool:
    """顶层目录内容控件：文档部件库标为目录，或里面是 TOC 域 / `_Toc` 超链接。"""
    if element.tag != _w("sdt"):
        return False
    gallery = _w_val(element, f"{_w('sdtPr')}/{_w('docPartObj')}/{_w('docPartGallery')}")
    if gallery and "Contents" in gallery:
        return True
    return (_is_toc_link_paragraph(element)
            or any((code.text or "").strip().upper().startswith("TOC") for code in element.iter(_w("instrText"))))


def _cover_paragraph_ids(body) -> set:
    """东吴模板首页单独成节：首个段落级分节符早于第一个顶层目录控件时，它所在的顶层元素及之前的
    段落属于首页。正文管线不输出首页，首页的题注不算进期望。没有目录控件时不认首页节，
    以免正文中途的分节符把前面整段正文藏出自检范围。"""
    top = list(body)
    first_break = next((i for i, element in enumerate(top) if element.tag == _w("p")
                        and element.find(f"{_w('pPr')}/{_w('sectPr')}") is not None), None)
    first_toc = next((i for i, element in enumerate(top) if _is_toc_control(element)), None)
    if first_break is None or first_toc is None or first_toc < first_break:
        return set()
    return {id(p) for element in top[:first_break + 1] for p in element.iter(_w("p"))}


class CaptionNumbering:
    """自动编号题注的编号推算：段落的 numPr（直接设置，或沿样式 basedOn 链继承）指向的编号级别，
    lvlText 形如「图%1：」时按 (abstractNumId, ilvl) 顺序计数，起始值取 start；共用同一 abstractNum 的
    不同 numId 接续同一序列，带 startOverride 的 numId 第一次出现时从该值重新起算。不处理 lvlRestart。"""

    def __init__(self, styles, numbering):
        self.style_num, self.style_based = {}, {}
        for style in (styles.iter(_w("style")) if styles is not None else ()):
            style_id = style.get(_w("styleId"))
            num_pr = style.find(f"{_w('pPr')}/{_w('numPr')}")
            if num_pr is not None:
                self.style_num[style_id] = (_w_val(num_pr, _w("numId")), _w_val(num_pr, _w("ilvl")))
            based = _w_val(style, _w("basedOn"))
            if based:
                self.style_based[style_id] = based
        self.levels, self.abstract_of, self.start_override, self.counters = {}, {}, {}, {}
        self.restarted = set()
        if numbering is None:
            return
        for abstract in numbering.iter(_w("abstractNum")):
            for level in abstract.findall(_w("lvl")):
                match = LEVEL_CAPTION_RE.match(_w_val(level, _w("lvlText")) or "")
                if match:
                    key = (abstract.get(_w("abstractNumId")), level.get(_w("ilvl")))
                    self.levels[key] = (match.group(1), int(_w_val(level, _w("start")) or 1))
        for num in numbering.iter(_w("num")):
            num_id = num.get(_w("numId"))
            self.abstract_of[num_id] = _w_val(num, _w("abstractNumId"))
            for override in num.findall(_w("lvlOverride")):
                start = _w_val(override, _w("startOverride"))
                if start is not None:
                    self.start_override[(num_id, override.get(_w("ilvl")))] = int(start)

    def _num_pr(self, paragraph):
        """(numId, ilvl)：两项各自取段落直接设置的值，缺的沿样式 basedOn 链往上找。"""
        props = paragraph.find(_w("pPr"))
        num_pr = props.find(_w("numPr")) if props is not None else None
        num_id, level = (_w_val(num_pr, _w("numId")), _w_val(num_pr, _w("ilvl"))) if num_pr is not None else (None, None)
        style = _w_val(props, _w("pStyle"))
        for _ in range(20):
            if (num_id is not None and level is not None) or not style:
                break
            style_id, style_level = self.style_num.get(style, (None, None))
            num_id = num_id if num_id is not None else style_id
            level = level if level is not None else style_level
            style = self.style_based.get(style)
        return (num_id, level or "0") if num_id is not None else None

    def take(self, paragraph):
        """该段是自动编号题注时计数并返回 (图/表, 编号)，否则返回 None。"""
        num = self._num_pr(paragraph)
        if not num or num[0] in (None, "0") or num[0] not in self.abstract_of:
            return None
        key = (self.abstract_of[num[0]], num[1])
        if key not in self.levels:
            return None
        kind, start = self.levels[key]
        if num in self.start_override and num not in self.restarted:
            self.restarted.add(num)
            self.counters[key] = self.start_override[num] - 1
        self.counters[key] = self.counters.get(key, start - 1) + 1
        return kind, self.counters[key]


def _walk_paragraphs(node, in_control: bool = False):
    """按文档顺序给出 (段落, 是否在内容控件里)；跳过 `mc:Fallback`（与 Choice 重复的一份）。"""
    for child in node:
        if child.tag == MC_FALLBACK:
            continue
        if child.tag == _w("p"):
            yield child, in_control
        yield from _walk_paragraphs(child, in_control or child.tag == _w("sdtContent"))


def docx_caption_numbers(source: Path):
    """源 .docx 正文题注的「图N/表N」编号；读不了时返回 None。

    跳过目录条目（内容控件里的段落、`_Toc` 超链接）与首页节；自动编号题注的文字里没有编号，
    按 numbering.xml 推算（首页的自动编号题注照样计数，只是不算进期望）。"""
    if source.suffix.lower() != ".docx":
        return None
    try:
        with zipfile.ZipFile(str(source)) as archive:
            names = set(archive.namelist())
            document = ElementTree.fromstring(archive.read("word/document.xml"))
            styles, numbering = (ElementTree.fromstring(archive.read(name)) if name in names else None
                                 for name in ("word/styles.xml", "word/numbering.xml"))
    except Exception:
        return None
    body = document.find(_w("body"))
    cover = _cover_paragraph_ids(body) if body is not None else set()
    numbers = CaptionNumbering(styles, numbering)
    found = {"图": set(), "表": set()}
    for paragraph, in_control in _walk_paragraphs(document):
        if in_control or _is_toc_link_paragraph(paragraph):
            continue
        caption = numbers.take(paragraph)
        if caption is None:
            match = CAPTION_RE.match(_xml_texts(paragraph))
            caption = (match.group(1), int(match.group(2))) if match else None
        if caption and id(paragraph) not in cover:
            found[caption[0]].add(caption[1])
    return found


def report_caption_check(source: Path, paragraph_texts: list, progress):
    """图表自检：对照源文档与成品的题注编号，跳过了哪些图/表直接点名。

    只认段落行首的题注（正文中“详见图2：”这类交叉引用不算保留）。
    返回：全部保留 -> 摘要文字；有跳过 -> ""；无法自检 -> None。"""
    expected = docx_caption_numbers(source)
    if expected is None or not (expected["图"] or expected["表"]):
        return None
    kept = {"图": set(), "表": set()}
    for text in paragraph_texts:
        match = CAPTION_RE.match(text)
        if match:
            kept[match.group(1)].add(int(match.group(2)))
    parts, missing = [], []
    for kind in ("图", "表"):
        if not expected[kind]:
            continue
        miss = sorted(expected[kind] - kept[kind])
        if miss:
            missing.append("、".join(f"{kind}{n}" for n in miss))
        else:
            parts.append(f"{len(expected[kind])} 张{kind}")
    if not parts and not missing:
        return None
    if missing:
        progress.note(f"[warn] 图表自检：{'，'.join(missing)}被跳过，请人工检查并补充")
        return ""
    progress.note(f"图表自检：{'、'.join(parts)}全部保留")
    return "、".join(parts)


SOURCE_LINE_RE = re.compile(r"^\s*(?:数据来源|资料来源|资料及数据来源|[Ss]ources?|注\s*\d*|单位)\s*[：:]")


def loose_key(text: str) -> str:
    return re.sub(r"[^0-9A-Za-z一-鿿]+", "", text or "").lower()


def docx_body_paragraphs(source: Path):
    """源 .docx 顶层正文段落的文本列表，以及每个顶层目录控件之后第一段在列表里的下标。

    表格/文本框里的段落和目录条目（`_Toc` 超链接）不算。读不了返回 None。"""
    if source.suffix.lower() != ".docx":
        return None
    try:
        with zipfile.ZipFile(str(source)) as archive:
            root = ElementTree.fromstring(archive.read("word/document.xml"))
    except Exception:
        return None
    body = root.find("w:body", {"w": WORD_NS})
    if body is None:
        return None
    paragraphs, toc_ends = [], []
    for child in body:
        if _is_toc_control(child):
            toc_ends.append(len(paragraphs))
        if child.tag != "{%s}p" % WORD_NS or _is_toc_link_paragraph(child):
            continue
        boxed = {id(t) for box in child.iter("{%s}txbxContent" % WORD_NS)
                 for t in box.iter("{%s}t" % WORD_NS)}
        text = "".join(t.text or "" for t in child.iter("{%s}t" % WORD_NS) if id(t) not in boxed)
        text = re.sub(r"\s+", " ", text).strip()
        if text:
            paragraphs.append(text)
    return paragraphs, toc_ends


def report_text_check(root: Path, source: Path, plain_text: str, progress):
    """文字自检：以「已进入成品的第一段～最后一段」为包络，包络内没匹配上的正文段落点名并给出章节定位。

    封面、目录、免责声明天然落在包络外；整表转图的表内文字不参与（只查顶层段落）。
    包络起点前移到第一个已匹配段之前最后一个目录控件之后：目录与第一章之间的导语整段丢失时也能报出。"""
    parsed = docx_body_paragraphs(source)
    if not parsed or not parsed[0]:
        return None
    paragraphs, toc_ends = parsed
    out_key = loose_key(plain_text)
    keys = [loose_key(p) for p in paragraphs]
    matched = [bool(k) and k in out_key for k in keys]
    anchors = [i for i, ok in enumerate(matched) if ok and len(keys[i]) >= 12]
    if not anchors:
        return None
    outline_keys = {}
    try:
        for entry in engine.extract_docx_heading_outline(source, body_matching(root)):
            outline_keys[loose_key(entry.get("text", ""))] = entry.get("text", "")
    except Exception:
        pass
    total = 0
    misses = []
    last_heading = ""
    begin = max([end for end in toc_ends if end <= anchors[0]], default=anchors[0])
    for i in range(begin, anchors[-1] + 1):
        text, key = paragraphs[i], keys[i]
        if key in outline_keys:
            last_heading = outline_keys[key]
        if len(key) < 12 or CAPTION_RE.match(text) or SOURCE_LINE_RE.match(text):
            continue
        total += 1
        if not matched[i]:
            snippet = text[:15] + ("…" if len(text) > 15 else "")
            misses.append(f"「{snippet}」" + (f"（{last_heading} 节下）" if last_heading else ""))
    if misses:
        shown = "、".join(misses[:3]) + (f" 等{len(misses)}段" if len(misses) > 3 else "")
        progress.note(f"[warn] 文字自检：正文 {total} 段中疑似跳过 {len(misses)} 段：{shown}，请人工检查并补充")
        return ""
    progress.note(f"文字自检：正文 {total} 段全部保留")
    return f"正文 {total} 段"


def summary_fragment(html: str, *, depth_style: bool = False) -> str:
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "html.parser")
    section = soup.new_tag("section")
    if depth_style:
        section["style"] = "margin:0 0 0.6em 0;"
        for child in list(soup.body.contents):
            section.append(child.extract())
        return str(section)
    # These are the fixed body/p/strong/u rules emitted by SummaryExtractor.
    for style in soup.find_all("style"):
        for selector, declarations in re.findall(r"([^{}]+)\{([^{}]+)\}", style.get_text()):
            selector = selector.strip()
            if selector == "body":
                declarations = re.sub(r"(?:^|;)\s*(?:margin|padding|max-width)\s*:[^;]*", "", declarations)
                section["style"] = declarations.strip().strip(";")
            else:
                for node in soup.body.select(selector):
                    node["style"] = declarations.strip() + ";" + node.get("style", "")
    for child in list(soup.body.contents):
        section.append(child.extract())
    return str(section)


FORECAST_NOTE_RE = re.compile(r"^注\s*\d*[：:]")
FORECAST_SOURCE_RE = re.compile(r"^(?:数据来源|资料来源|资料及数据来源|source|sources)\s*[：:]", re.IGNORECASE)
FORECAST_BOLD_RE = re.compile(r"font-weight:\s*(?:[6-9]00|bold)")


def _forecast_note_like(text: str) -> bool:
    """空段、带冒号的来源段（数据来源/资料来源/Source：）或编号注（注：/注1：）。"""
    if not text:
        return True
    return bool(FORECAST_SOURCE_RE.match(text) or FORECAST_NOTE_RE.match(text))


def _is_bold(tag) -> bool:
    return tag.name in ("b", "strong") or bool(FORECAST_BOLD_RE.search(tag.get("style") or ""))


FORECAST_KEYWORD = "三大财务预测表"


def _all_text_bold(block) -> bool:
    """段内每一段可见文字都落在加粗节点内（容器层级 <p><span><strong>… 不限）。"""
    from bs4 import Comment

    if _is_bold(block):
        return True
    for node in block.find_all(string=True):
        # 注释不是可见文字：算进来会让本该剥除的真标题漏判
        if isinstance(node, Comment) or not str(node).strip():
            continue
        parent = node.parent
        while parent is not None and parent is not block.parent:
            if getattr(parent, "name", None) and _is_bold(parent):
                break
            parent = parent.parent
        else:
            return False
    return True


# 关键词之前只写证券简称：711 份真实报告的 63 种三表标题，前缀最长的「京东集团-SW」
# 只有 7 个字符，没有一种带代码或句读。上限取 16 是给拉丁文简称留的余量
# （`Amazon.com, Inc.` 恰好 16）。只允许「像名字」的字符：中西文字、数字，以及证券
# 简称里真实出现的标点（`·・-–—&.!！*+/／'’` 与空格，覆盖 `*ST` 前缀、`Yum!`、
# `McDonald's`），且至少要有一个文字或数字（纯标点不算名字）。逗号不在名单内：
# 简称里没有它，中文叙述句里遍地都是（「…盈利改善，安克创新三大财务预测表」）
FORECAST_NAME_MAX = 16
FORECAST_NAME_ALLOWED_RE = re.compile(r"^(?=.*[^\W_])(?:[^\W_]|[·・\-–—&.!！*+/／'’ 　])+$")


def _forecast_name_like(prefix: str) -> bool:
    """关键词之前的文字是否形如证券简称（空串也算：裸关键词标题）。"""
    if not prefix:
        return True
    return (2 <= len(prefix) <= FORECAST_NAME_MAX
            and bool(FORECAST_NAME_ALLOWED_RE.match(prefix)))


def _forecast_heading_like(block, text: str) -> bool:
    """真实三表标题段：整段加粗，文字是「证券简称 + 三大财务预测表」或裸关键词。

    这个标题由模板宏按证券简称字段生成，形态固定，所以**不再**要求它与本报告的
    公司名字段对上：实测语料里 62 份报告的封面控件/页眉停留在模板自带的另一家公司
    或占位文字上（如标题写「珀莱雅」而页眉仍是「大商股份（603605）」），字段比对
    只会让这些报告漏剥，字段本身并不比标题更可信。
    判定改由位置承担：该段须整段加粗，其后直到正文结束只剩表格图与来源/注段
    （见 strip_forecast_tables），也就是它必须是正文最后一张表的表题。"""
    if not text.endswith(FORECAST_KEYWORD) or _forecast_note_like(text):
        return False
    return _forecast_name_like(text[:-len(FORECAST_KEYWORD)]) and _all_text_bold(block)


def locate_forecast_tables(soup) -> list:
    """找出正文末尾的「XX三大财务预测表」区块（标题段 + 表格图片 + 来源/注），返回其顶层块列表。

    只认**最早**的加粗标题候选：其后直到正文结束只允许纯图片段（去图后无文字）
    与来源/注说明段，且至少要有一个图片段；验证失败即返回空列表，不再尝试更靠后的
    候选（防止真三表保留却误报“已省略”，或把半截正文当三表附给点评）。"""
    body = soup.body if soup.body is not None else soup
    blocks = [child for child in body.children if getattr(child, "name", None)]
    for start, block in enumerate(blocks):
        text = re.sub(r"\s+", "", block.get_text())
        if block.name != "p" or not _forecast_heading_like(block, text):
            continue
        has_image = False
        for tail in blocks[start + 1:]:
            if tail.name != "p":
                return []
            tail_text = re.sub(r"\s+", "", tail.get_text())
            if tail.find("img"):
                if tail_text:      # 图片段夹带文字：不是纯表格图，宁可整体保留
                    return []
                has_image = True
            elif not _forecast_note_like(tail_text):
                return []
        if not has_image:
            return []
        return blocks[start:]
    return []


def strip_forecast_tables(soup) -> bool:
    """深度框架默认路径：按设置省略文末三表区块。返回是否移除。"""
    doomed = locate_forecast_tables(soup)
    for block in doomed:
        block.decompose()
    return bool(doomed)


def extract_forecast_tables(soup) -> list:
    """点评框架勾选三表时的反向路径：把同一判据认出的区块从正文成品里摘出来，
    供拼到摘要之后。找不到返回空列表，正文成品不动。"""
    blocks = locate_forecast_tables(soup)
    for block in blocks:
        block.extract()
    return blocks


def self_check_hint(modes: list, checks_ran: bool, forecast_appended: bool = False) -> str:
    """完成但无自检结论（灰色信息条）时的提示语，按模式区分。"""
    if not any(m in modes for m in ("body", "weekly")):
        if forecast_appended:
            return "确认摘要文字与三表后复制到公众号编辑器"
        return "点评摘要不含图表，确认文字内容后复制到公众号编辑器"
    if not checks_ran:
        return "未能自动对账图表与文字，请人工核对后复制到公众号编辑器"
    return ""


def wants_forecast_append(modes: list, forecast_tables: bool) -> bool:
    """点评框架（只取摘要）勾选三表时，另跑一遍正文管线把文末三表摘出来附在摘要后。
    正文/周报框架不走这条：正文框架的三表本就在成品里，由 strip 决定去留。"""
    return bool(forecast_tables) and modes == ["summary"]


def plan_steps(modes: list, forecast_append: bool) -> int:
    return 3 + len(modes) + (1 if forecast_append else 0)


def note_modes(progress, modes: list, forecast_append: bool) -> None:
    """[modes] 行供结果页按实际框架估算完成时间；点评附三表要多跑一遍正文管线，
    以 forecast 伪模式告知前端把这一遍算进预估。"""
    progress.note("[modes] " + ",".join(modes + (["forecast"] if forecast_append else [])))


def forecast_tail_problems(cleaner) -> list:
    """清洗器记录的丢图问题里，落在三表区块范围内的那些。

    这个范围是按**文字**推出来的保守区间，起点可能早于真正的标题段（见下），
    所以「区块之内」是宽口径，不是严格等于摘取器认定的那几块。

    记录都在 `HtmlCleaner` 里就地产生并带上源块序号：主判据是逐地址对账——源块里
    这个图片地址出现几次、成品就该用掉几次，少用一次记一条（记在真正进入 spec 的
    时点，不是读到的时点：表格的题注/来源行会连图一起被剪掉）；另有图片地址解析
    不出来、表格转图时单元格里的图读不到、整块渲染后一张图都不剩三种情形单独记录。
    判据建在渲染器自己的记录上，而不是另写一套「源文档里该有几张图」去和它对账——
    那要把渲染器的分类规则（嵌套表算一张、表内图并进表图、兼容分支只算一支）
    复制一份，两处迟早对不上。

    序号用来把判断大致限定在区块附近：报告别处的丢图不该让用户白白丢掉这个功能。
    起点取**最早**一个够格当标题的已渲染块，与摘取器只认最早候选的规则一致：
    文字以关键词结尾、不是注/来源段、其前缀形如证券简称。取最早而不是最晚，是因为
    这里只有文字、没有加粗信息：正文里若先有一句「详见文末XX三大财务预测表」，
    起点就落在它上面，范围比真区块宽——代价是多一次「放弃并告警」，
    换来的是文末一句「注：详见XX三大财务预测表」不会把起点推到真标题之后、
    把区块内的丢图记录挡在范围外。"""
    problems = list(getattr(cleaner, "render_problems", []))
    if not problems:
        return []
    start = None
    for idx, text in getattr(cleaner, "rendered_block_texts", []):
        text = re.sub(r"\s+", "", text)
        if not text.endswith(FORECAST_KEYWORD) or _forecast_note_like(text):
            continue
        if _forecast_name_like(text[:-len(FORECAST_KEYWORD)]):
            start = idx
            break
    if start is None:
        return [p["message"] for p in problems]      # 定位不到区块：一律从严
    return [p["message"] for p in problems if p["idx"] >= start]


def render_forecast_tables(root: Path, source: Path, word, temporary: Path) -> list:
    """按正文管线转换全文，只取文末三表区块（与深度框架的剥除判据完全相同）。

    这一遍的引擎输出全部截住，不进任务日志：全文里大半内容（封面、首页估值表）转完
    就丢，它们的 skipped/fallback 提示与成品无关，不该在点评结果页上亮红条。

    整遍失败即放弃（抛出说明，调用方转成一条告警、仍输出摘要）。渲染器一贯的策略是
    「失败就跳过这一张、继续往下」，区块里若有一张表没渲染出来、另一张还在，位置判据
    仍然通过，半截三表就会被当成完整的附上去。两层判据：

    ① 区块范围内的丢图记录（`forecast_tail_problems`）：清洗器按「源块里这个地址出现
       几次、成品就该用掉几次」记账并就地记录，与源文档格式无关，`.doc` 同样有效；
       按源块序号取用，范围是按文字推出的保守区间（起点可能早于真标题段，宁宽不窄）。
    ② 摘出的区块里每张图必须是内嵌的 `data:` 图，不能是死链。

    这一遍的引擎输出仍然整体截住不进任务日志（全文里大半内容转完就丢，它们的提示与
    成品无关），但**不再**拿日志行当判据：真正的丢图都已在 ① 里带序号记下，
    拿全局日志行判断只会让区块之外的问题误伤一次本来完整的附加。

    覆盖范围：核对的是「渲染器有没有把源块里的图弄丢」，不是「成品与源文档逐项对应」。
    渲染器把某张表换成另一张图而自身毫无察觉，不在覆盖范围内。"""
    import io
    from contextlib import redirect_stdout
    from bs4 import BeautifulSoup

    cfg = engine.load_config("body", root / "configs")
    cfg["output"].update(generate_html=True, generate_docx=False)
    cfg.setdefault("front_matter", {}).update(enable=False, include_summary=False, include_toc_image=False)
    part_dir = temporary / "forecast"
    part_dir.mkdir()
    cleaner = engine.HtmlCleaner(cfg)
    captured = io.StringIO()
    with redirect_stdout(captured):
        result = engine.process_one_file(source, part_dir, cleaner, word, cfg)
    rendered = Path(result["html"]).read_text(encoding="utf-8")
    blocks = extract_forecast_tables(BeautifulSoup(rendered, "html.parser"))
    images = [image for block in blocks for image in block.find_all("img")]
    problems = forecast_tail_problems(cleaner)
    if problems:
        raise RuntimeError("三表区块的转换有内容丢失，为免附上残缺的三表而放弃：" + problems[0])
    for image in images:
        if not str(image.get("src") or "").startswith("data:"):
            raise RuntimeError("三表区块里有未内嵌的图片，为免附上残缺的三表而放弃。")
    return blocks


def convert(root: Path, source: Path, modes: list, output: Path, toc: bool,
            forecast_tables: bool = False) -> dict:
    from bs4 import BeautifulSoup

    output.parent.mkdir(parents=True, exist_ok=True)
    old_temp = tempfile.tempdir
    progress = Progress()
    auto = modes == [AUTO]
    detected = None
    title_kind = ""
    forecast_append = False
    if auto:
        progress.note("正在识别报告类型…")
        detected = detect_kind_from_docx(root, source)
        if detected:
            modes, toc, title_kind = apply_detection(detected, toc, progress)
    if not auto or detected:
        forecast_append = wants_forecast_append(modes, forecast_tables)
        note_modes(progress, modes, forecast_append)
    progress.plan(5 if auto and not detected else plan_steps(modes, forecast_append))
    try:
        with tempfile.TemporaryDirectory(prefix="conversion_", dir=output.parent, ignore_cleanup_errors=True) as temporary:
            tempfile.tempdir = temporary
            fragments = []
            forecast_blocks = []
            final_html = None
            progress.step("启动 Word 引擎")
            with engine.WordApp(image_ppi=engine.resolve_image_ppi(engine.load_config("body", root / "configs"))) as word:
                if auto and not detected:
                    detected = detect_kind_with_word(root, word, source)
                    modes, toc, title_kind = apply_detection(detected, toc, progress)
                    forecast_append = wants_forecast_append(modes, forecast_tables)
                    note_modes(progress, modes, forecast_append)
                    progress.plan(plan_steps(modes, forecast_append))
                configs = {mode: engine.load_config(mode, root / "configs")
                           for mode in MODES if mode in modes}
                if "weekly" in configs:
                    image = base64.b64encode((root / "assets" / "weekly-heading.png").read_bytes()).decode("ascii")
                    configs["weekly"]["output"]["weekly_heading_level_1_background_image"] = "data:image/png;base64," + image
                progress.step("载入报告版式")
                for mode, cfg in configs.items():
                    progress.step(f"转换{MODE_LABELS[mode]}（Word 处理中，长报告需要数分钟）")
                    cfg["output"].update(generate_html=True, generate_docx=False)
                    if mode != "summary":
                        cfg.setdefault("front_matter", {}).update(
                            enable=toc, include_summary=False, include_toc_image=toc)
                    part_dir = Path(temporary) / mode
                    part_dir.mkdir()
                    converter = (engine.SummaryExtractor(cfg) if mode == "summary"
                                 else engine.HtmlCleaner(cfg))
                    process = engine.process_summary_file if mode == "summary" else engine.process_one_file
                    if mode == "summary" and len(configs) > 1:
                        # 摘要 + 正文/周报：摘要取不到时照常交付正文，告警请人工补充摘要；只转摘要时摘要就是全部内容，照常失败
                        try:
                            result = process(source, part_dir, converter, word, cfg)
                        except Exception as exc:
                            progress.note(f"[warn] 摘要未能提取，已只输出正文，请对照原文人工补充摘要：{type(exc).__name__}: {exc}")
                            continue
                    else:
                        result = process(source, part_dir, converter, word, cfg)
                    rendered = Path(result["html"]).read_text(encoding="utf-8")
                    if mode == "summary":
                        fragments.append(summary_fragment(rendered, depth_style="body" in modes))
                        if toc and len(configs) == 1:
                            fragments.append(engine.build_content_toc_fragment(
                                source, part_dir, word, engine.load_config("body", root / "configs")))
                    final_html = rendered
                    print(f"  -> {MODE_LABELS[mode]}完成", flush=True)
                if forecast_append:
                    progress.step("提取文末三大财务预测表（按正文管线转换全文）")
                    try:
                        forecast_blocks = render_forecast_tables(root, source, word, Path(temporary))
                    except Exception as exc:
                        # 摘要已经拿到手：三表这一遍失败只告警，不让整份点评作废
                        progress.note(f"[warn] 未能提取文末三大财务预测表：{exc}")
                    else:
                        if forecast_blocks:
                            print("  -> 三表完成", flush=True)
                        else:
                            progress.note("[warn] 未找到文末三大财务预测表，仅输出摘要")
            progress.step("合并内容并写出 HTML")
            soup = BeautifulSoup(final_html, "html.parser")
            main_title = resolve_main_title(word.main_title, word.subtitle, source)
            stripped_forecast = ("body" in modes and not forecast_tables
                                 and strip_forecast_tables(soup))
            if stripped_forecast:
                progress.note("已按设置省略文末三大财务预测表")
            if modes == ["summary"]:
                soup.body.clear()
                for style in soup.find_all("style"):
                    body_rule = re.search(r"\bbody\s*\{([^{}]+)\}", style.get_text())
                    if body_rule:
                        soup.body["style"] = body_rule.group(1).strip()
                    style.decompose()
            for fragment in reversed(fragments):
                soup.body.insert(0, BeautifulSoup(fragment, "html.parser"))
            for block in forecast_blocks:
                soup.body.append(block)
            if forecast_blocks:
                progress.note("已按设置在摘要后附上文末三大财务预测表")
            if not soup.body.get_text(strip=True) and not soup.body.find("img"):
                raise RuntimeError("转换结果为空，未生成 HTML。")
            output.write_text(str(soup), encoding="utf-8")
            self_check = ""
            ran = []
            if any(m in modes for m in ("body", "weekly")):
                # 只对账正文/周报部分：摘要逐字引用正文原句时，合并全文会掩盖正文缺段
                body_soup = BeautifulSoup(final_html, "html.parser")
                if stripped_forecast:
                    # 自检必须对账实际写出的内容：同样剥除三表后再对账
                    strip_forecast_tables(body_soup)
                plain_text = body_soup.get_text(" ")
                paragraph_texts = [p.get_text() for p in body_soup.find_all("p")]
                results = []
                for label, check in (("图表自检", lambda: report_caption_check(source, paragraph_texts, progress)),
                                     ("文字自检", lambda: report_text_check(root, source, plain_text, progress))):
                    try:
                        results.append(check())
                    except Exception as exc:
                        print(f"  [warn] {label}未能完成：{type(exc).__name__}: {exc}", flush=True)
                        results.append(None)
                ran = [r for r in results if r is not None]
                if ran and all(ran):
                    self_check = "；".join(ran)
            return {"title": word.report_title, "main_title": main_title,
                    "subtitle": word.subtitle, "report_kind": KIND_LABELS.get(title_kind, ""),
                    "self_check": self_check, "forecast_stripped": bool(stripped_forecast),
                    "forecast_appended": bool(forecast_blocks),
                    "self_check_hint": self_check_hint(modes, bool(ran), bool(forecast_blocks))}
    finally:
        tempfile.tempdir = old_temp


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--modes", required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--metadata", type=Path)
    parser.add_argument("--toc", action="store_true")
    parser.add_argument("--forecast-tables", action="store_true")
    args = parser.parse_args()
    modes = args.modes.split(",")
    if modes != [AUTO]:
        if not modes or any(mode not in MODES for mode in modes) or len(set(modes)) != len(modes):
            parser.error("--modes 必须是 auto，或 summary、body、weekly 的非重复组合")
        if "body" in modes and "weekly" in modes:
            parser.error("body 与 weekly 不能同时选择")
    source = args.source.resolve()
    output = args.output.resolve()
    if not source.is_file() or source.suffix.lower() not in {".doc", ".docx"}:
        parser.error("--source 必须是现有 .doc 或 .docx 文件")
    if output == source or output.suffix.lower() != ".html":
        parser.error("--output 必须是独立的 .html 文件")
    metadata = args.metadata.resolve() if args.metadata else None
    if metadata and (metadata in {source, output} or metadata.suffix.lower() != ".json"):
        parser.error("--metadata 必须是独立的 .json 文件")
    try:
        title_fields = convert(args.root.resolve(), source, modes, output, args.toc, args.forecast_tables)
        if metadata:
            try:
                metadata.parent.mkdir(parents=True, exist_ok=True)
                metadata.write_text(json.dumps(title_fields, ensure_ascii=False), encoding="utf-8")
            except OSError as exc:
                print(f"  [warn] 标题等信息未能写出：标题请手动填写，自检与文末三表的处理情况无法显示（{exc}）", flush=True)
        print(f"HTML：{output}", flush=True)
        return 0
    except Exception:
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
