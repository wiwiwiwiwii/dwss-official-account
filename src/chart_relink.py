# -*- coding: utf-8 -*-
"""
正文导出前把本地图表的数据源改写为「链接到不存在的外部工作簿、不自动更新」。

Word 导出筛选 HTML 时，数据没有链接外部工作簿的图表（内嵌 xlsx 包、内嵌旧式 xls、没有数据）会逐张向
Excel / excelcnv 发同步 COM 调用并等它返回（诊断见 tests/perf-v117/saveas-diagnosis.md）。改写后 Word 按图表 XML
里的缓存渲染，不再等待；缓存值与格式一字不动。方案与守卫见 tests/perf-v117/design.md 第二节。

纯 zip/XML 处理，不依赖 Word：
- 按 [Content_Types].xml 的 drawingml.chart+xml 枚举图表部件（chartex 跳过）；
- 只做定点文本插入/替换，不整份重序列化 XML；改写后重新解析并核对「除 externalData 外整棵树不变」，
  不通过则该图表原样保留；
- 命中守卫（缺缓存、ptCount 不符、透视图；文档里有依赖路径的域或相对路径的外部关系）的图表或整份文档不改写；
- 其余 zip 条目按原顺序原样复制。

用法（调用方）：
    result = relink_local_charts(src, dst, missing_target)
    if result["changed"]: 用 dst 导出，否则用 src。
"""

from __future__ import annotations

import io
import posixpath
import re
import shutil
import zipfile
from collections import Counter
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
OLE_LINK_TYPE = REL_NS + "/oleObject"
HYPERLINK_TYPE = REL_NS + "/hyperlink"

CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CHARTEX_SUFFIX = "chartex+xml"  # application/vnd.ms-office.chartex+xml
# 会被扫描域代码的部件：正文、页眉页脚、脚注尾注（文本框在这些部件内部）
FIELD_PART_SUFFIXES = (
    "wordprocessingml.document.main+xml", "wordprocessingml.template.main+xml",
    "ms-word.document.macroEnabled.main+xml", "ms-word.template.macroEnabledTemplate.main+xml",
    "wordprocessingml.header+xml", "wordprocessingml.footer+xml",
    "wordprocessingml.footnotes+xml", "wordprocessingml.endnotes+xml",
)
# 依赖文档路径或外部文件的域：命中即整份不改写（改写副本与原副本不在同一目录）
PATH_FIELDS = re.compile(r"\b(FILENAME|INCLUDEPICTURE|INCLUDETEXT|LINK|DDE|DDEAUTO)\b", re.IGNORECASE)
# 图表根元素下 chart 之后允许出现的子元素（schema 顺序）；externalData 插在 txPr 之后、printSettings 之前
AFTER_CHART = ("spPr", "txPr", "externalData", "printSettings", "userShapes", "extLst")
INSERT_BEFORE = ("printSettings", "userShapes", "extLst")

MAX_ENTRIES = 5000
MAX_PART_BYTES = 100 * 1024 * 1024

START_TAG = r"<{name}(?:\s+[^\s=/>]+\s*=\s*(?:\"[^\"]*\"|'[^']*'))*\s*/?>"
ROOT_START = re.compile(r"<(?![?!])[^\s/>]+(?:\s+[^\s=/>]+\s*=\s*(?:\"[^\"]*\"|'[^']*'))*\s*/?>")
XMLNS_ATTR = re.compile(r"\sxmlns(?::([\w.\-]+))?\s*=\s*(?:\"([^\"]*)\"|'([^']*)')")
ATTR = r"\s{name}\s*=\s*(?:\"([^\"]*)\"|'([^']*)')"


class Skip(Exception):
    """守卫命中或无法安全改写：该图表（或整份文档）留在旧路径。"""


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _ns(tag: str) -> str:
    return tag[1:].split("}", 1)[0] if tag.startswith("{") else ""


def _decode(data: bytes) -> Tuple[str, bytes]:
    """只处理 UTF-8（可带 BOM）；返回 (文本, BOM)。"""
    bom = b""
    if data.startswith(b"\xef\xbb\xbf"):
        bom, data = data[:3], data[3:]
    head = data[:200].decode("ascii", "replace")
    declared = re.match(r"<\?xml[^>]*encoding\s*=\s*[\"']([^\"']+)[\"']", head)
    if declared and declared.group(1).lower().replace("_", "-") not in {"utf-8", "utf8"}:
        raise Skip("encoding")
    try:
        return data.decode("utf-8"), bom
    except UnicodeDecodeError:
        raise Skip("encoding")


def _attr(tag_text: str, name: str) -> Optional[str]:
    found = re.search(ATTR.format(name=re.escape(name)), tag_text)
    if not found:
        return None
    return found.group(1) if found.group(1) is not None else found.group(2)


def _root_start(text: str) -> re.Match:
    found = ROOT_START.search(text)
    if not found:
        raise Skip("no_root")
    return found


def _prefix_for(start_tag: str, uri: str) -> Optional[str]:
    """根起始标签上绑定到 uri 的前缀：'' 表示默认命名空间，None 表示未在根上声明。"""
    for match in XMLNS_ATTR.finditer(start_tag):
        value = match.group(2) if match.group(2) is not None else match.group(3)
        if value == uri:
            return match.group(1) or ""
    return None


def _declared_prefixes(start_tag: str) -> set:
    return {m.group(1) or "" for m in XMLNS_ATTR.finditer(start_tag)}


def _qname(prefix: str, local: str) -> str:
    return f"{prefix}:{local}" if prefix else local


def _rels_name(part: str) -> str:
    folder, base = posixpath.split(part)
    return posixpath.join(folder, "_rels", base + ".rels")


# ----------------------------- 守卫 -----------------------------

def chart_guard(root: ET.Element) -> Optional[str]:
    """单张图表的守卫：返回命中原因，None 表示可改写。"""
    if any(_local(e.tag) == "pivotSource" for e in root.iter()):
        return "pivot"
    for parent in root.iter():
        children = list(parent)
        if any(_local(c.tag) == "f" for c in children):
            if not any(_local(c.tag).endswith("Cache") for c in children):
                return "missing_cache"
    for cache in root.iter():
        if not _local(cache.tag).endswith("Cache"):
            continue
        count = next((c for c in cache if _local(c.tag) == "ptCount"), None)
        try:
            total = int(count.get("val")) if count is not None else None
        except (TypeError, ValueError):
            return "bad_ptcount"
        for pt in cache.iter():
            if _local(pt.tag) != "pt":
                continue
            try:
                idx = int(pt.get("idx"))
            except (TypeError, ValueError):
                return "bad_pt"
            if total is None or idx >= total or idx < 0:
                return "pt_out_of_range"
    return None


def _field_instructions(data: bytes) -> List[str]:
    """按文档顺序收集复杂域（fldChar begin…separate/end 之间的 instrText，可跨多个 run、可嵌套）与 fldSimple 的指令。"""
    out: List[str] = []
    stack: List[dict] = []
    instr_tags = {f"{{{W_NS}}}instrText", f"{{{W_NS}}}delInstrText"}
    for _, elem in ET.iterparse(io.BytesIO(data), events=("end",)):
        tag = elem.tag
        if tag == f"{{{W_NS}}}fldChar":
            kind = elem.get(f"{{{W_NS}}}fldCharType")
            if kind == "begin":
                stack.append({"text": [], "done": False})
            elif kind in {"separate", "end"} and stack:
                field = stack[-1]
                if not field["done"]:
                    out.append("".join(field["text"]))
                    field["done"] = True
                if kind == "end":
                    stack.pop()
        elif tag in instr_tags:
            if stack and not stack[-1]["done"]:
                stack[-1]["text"].append(elem.text or "")
            else:
                out.append(elem.text or "")
        elif tag == f"{{{W_NS}}}fldSimple":
            out.append(elem.get(f"{{{W_NS}}}instr") or "")
    out.extend("".join(f["text"]) for f in stack if not f["done"])
    return out


def _is_relative_target(target: str) -> bool:
    if re.match(r"^[A-Za-z][A-Za-z0-9+.\-]*:", target) or target.startswith(("\\\\", "//")):
        return False
    return True


def document_guard(archive: zipfile.ZipFile, content_types: Dict[str, str]) -> Optional[str]:
    """整份文档的守卫：依赖路径的域、相对路径的外部关系（hyperlink 除外）。"""
    for part, kind in content_types.items():
        if kind.endswith(FIELD_PART_SUFFIXES):
            data = _read(archive, part)
            if data is None:
                continue
            try:
                instructions = _field_instructions(data)
            except ET.ParseError:
                return "field_part_unparsable"
            if any(PATH_FIELDS.search(text) for text in instructions):
                return "path_field"
    for info in archive.infolist():
        if not info.filename.endswith(".rels"):
            continue
        data = _read(archive, info.filename)
        try:
            rels = ET.fromstring(data)
        except ET.ParseError:
            return "rels_unparsable"
        for rel in rels:
            if rel.get("TargetMode") != "External" or rel.get("Type") == HYPERLINK_TYPE:
                continue
            if _is_relative_target(rel.get("Target") or ""):
                return "relative_external"
    return None


def _read(archive: zipfile.ZipFile, name: str) -> Optional[bytes]:
    try:
        info = archive.getinfo(name)
    except KeyError:
        return None
    if info.file_size > MAX_PART_BYTES:
        raise Skip("part_too_large")
    return archive.read(info)


# ----------------------------- 单张图表改写 -----------------------------

def _without_external_data(root: ET.Element) -> bytes:
    clone = ET.fromstring(ET.tostring(root))
    for child in list(clone):
        if child.tag == f"{{{CHART_NS}}}externalData":
            clone.remove(child)
    return ET.tostring(clone)


def _rels_entries(rels_text: str) -> Tuple[str, List[re.Match]]:
    root = _root_start(rels_text)
    prefix = _prefix_for(root.group(0), PKG_REL_NS)
    if prefix is None:
        raise Skip("rels_namespace")
    pattern = re.compile(START_TAG.format(name=re.escape(_qname(prefix, "Relationship")) + r"(?=[\s/>])"))
    return prefix, list(pattern.finditer(rels_text))


def _new_rel(prefix: str, rid: str, target: str) -> str:
    return (f'<{_qname(prefix, "Relationship")} Id="{rid}" Type="{OLE_LINK_TYPE}" '
            f'Target="{_xml_attr(target)}" TargetMode="External"/>')


def _xml_attr(value: str) -> str:
    return value.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")


def relink_chart(chart_bytes: bytes, rels_bytes: Optional[bytes], target_uri: str) -> Tuple[str, bytes, bytes]:
    """返回 (动作, 新图表 XML, 新 rels)；动作为 retargeted / inserted / linked。不能安全改写时抛 Skip。"""
    chart_text, chart_bom = _decode(chart_bytes)
    try:
        root = ET.fromstring(chart_bytes)
    except ET.ParseError:
        raise Skip("chart_unparsable")
    if root.tag != f"{{{CHART_NS}}}chartSpace":
        raise Skip("not_chartspace")
    rels_text, rels_bom = _decode(rels_bytes) if rels_bytes is not None else (None, b"")
    rels_root = None
    if rels_text is not None:
        try:
            rels_root = ET.fromstring(rels_bytes)
        except ET.ParseError:
            raise Skip("rels_unparsable")
    relations = {r.get("Id"): r for r in rels_root} if rels_root is not None else {}

    root_tag = _root_start(chart_text)
    cp = _prefix_for(root_tag.group(0), CHART_NS)
    if cp is None:
        raise Skip("chart_namespace_not_on_root")
    children = list(root)
    names = [(_ns(c.tag), _local(c.tag)) for c in children]
    try:
        chart_index = names.index((CHART_NS, "chart"))
    except ValueError:
        raise Skip("no_chart_element")
    tail = names[chart_index + 1:]
    if any(ns != CHART_NS or name not in AFTER_CHART for ns, name in tail):
        raise Skip("unexpected_layout")
    order = [AFTER_CHART.index(name) for _, name in tail]
    if order != sorted(order) or len(set(order)) != len(order):
        raise Skip("unexpected_layout")
    external = next((c for c in children if c.tag == f"{{{CHART_NS}}}externalData"), None)
    ext_q = re.escape(_qname(cp, "externalData"))
    if external is not None:
        linked = relations.get(external.get(f"{{{REL_NS}}}id") or "")
        if linked is not None and linked.get("TargetMode") == "External":
            return "linked", chart_bytes, rels_bytes or b""
    reason = chart_guard(root)
    if reason:
        raise Skip(reason)

    if external is not None:
        rid = external.get(f"{{{REL_NS}}}id")
        if not rid:
            raise Skip("externaldata_no_id")
        rel = relations.get(rid)
        if rel is None:
            raise Skip("externaldata_rel_missing")
        starts = list(re.finditer(START_TAG.format(name=ext_q + r"(?=[\s/>])"), chart_text))
        if len(starts) != 1:
            raise Skip("ambiguous_text")
        start = starts[0]
        auto = _qname(cp, "autoUpdate")
        if start.group(0).endswith("/>"):
            new_element = start.group(0)[:-2].rstrip() + f'><{auto} val="0"/></{_qname(cp, "externalData")}>'
            span = (start.start(), start.end())
        else:
            close = re.compile(rf"</{ext_q}\s*>").search(chart_text, start.end())
            if not close:
                raise Skip("ambiguous_text")
            inner = chart_text[start.end():close.start()]
            auto_tags = list(re.finditer(START_TAG.format(name=re.escape(auto) + r"(?=[\s/>])"), inner))
            if len(auto_tags) > 1:
                raise Skip("ambiguous_text")
            if auto_tags:
                tag = auto_tags[0].group(0)
                if _attr(tag, "val") is not None:
                    fixed = re.sub(ATTR.format(name="val"), ' val="0"', tag, count=1)
                else:
                    fixed = tag[:-2].rstrip() + ' val="0"/>' if tag.endswith("/>") else tag[:-1] + ' val="0">'
                inner = inner[:auto_tags[0].start()] + fixed + inner[auto_tags[0].end():]
            else:
                inner = f'<{auto} val="0"/>' + inner
            new_element = start.group(0) + inner + chart_text[close.start():close.end()]
            span = (start.start(), close.end())
        new_chart = chart_text[:span[0]] + new_element + chart_text[span[1]:]
        prefix, entries = _rels_entries(rels_text)
        hits = [m for m in entries if _attr(m.group(0), "Id") == rid]
        if len(hits) != 1 or not hits[0].group(0).endswith("/>"):
            raise Skip("rels_unusual")
        new_rels = rels_text[:hits[0].start()] + _new_rel(prefix, rid, target_uri) + rels_text[hits[0].end():]
        action = "retargeted"
    else:
        rid = _unique_id(set(relations))
        rp = _prefix_for(root_tag.group(0), REL_NS)
        declaration = ""
        if not rp:
            taken = _declared_prefixes(root_tag.group(0))
            rp = next(p for p in ("r", "rl", "rrel", "relns") if p not in taken)
            declaration = f' xmlns:{rp}="{REL_NS}"'
        new_element = (f'<{_qname(cp, "externalData")}{declaration} {rp}:id="{rid}">'
                       f'<{_qname(cp, "autoUpdate")} val="0"/></{_qname(cp, "externalData")}>')
        closes = list(re.finditer(rf"</{re.escape(_qname(cp, 'chart'))}\s*>", chart_text))
        if len(closes) != 1:
            raise Skip("ambiguous_text")
        position = None
        for _, name in tail:
            if name in INSERT_BEFORE:
                found = re.compile(START_TAG.format(name=re.escape(_qname(cp, name)) + r"(?=[\s/>])")).search(
                    chart_text, closes[0].end())
                if not found:
                    raise Skip("ambiguous_text")
                position = found.start()
                break
        if position is None:
            ends = list(re.finditer(rf"</{re.escape(_qname(cp, 'chartSpace'))}\s*>", chart_text))
            if len(ends) != 1:
                raise Skip("ambiguous_text")
            position = ends[0].start()
        new_chart = chart_text[:position] + new_element + chart_text[position:]
        if rels_text is None:
            new_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                        f'<Relationships xmlns="{PKG_REL_NS}">{_new_rel("", rid, target_uri)}</Relationships>')
        else:
            prefix, _ = _rels_entries(rels_text)
            rels_root_tag = _root_start(rels_text)
            closing = list(re.finditer(rf"</{re.escape(_qname(prefix, 'Relationships'))}\s*>", rels_text))
            if rels_root_tag.group(0).endswith("/>") and not closing:
                new_rels = (rels_text[:rels_root_tag.start()] + rels_root_tag.group(0)[:-2].rstrip() + ">"
                            + _new_rel(prefix, rid, target_uri) + f"</{_qname(prefix, 'Relationships')}>"
                            + rels_text[rels_root_tag.end():])
            elif len(closing) == 1:
                at = closing[0].start()
                new_rels = rels_text[:at] + _new_rel(prefix, rid, target_uri) + rels_text[at:]
            else:
                raise Skip("rels_unusual")
        action = "inserted"

    chart_out = chart_bom + new_chart.encode("utf-8")
    rels_out = rels_bom + new_rels.encode("utf-8")
    _verify(root, chart_text, chart_out, relations, rels_out, rid, target_uri)
    return action, chart_out, rels_out


def _unique_id(taken: set) -> str:
    index = 1
    while f"rIdRelink{index}" in taken:
        index += 1
    return f"rIdRelink{index}"


def _verify(old_root, old_text, chart_out, old_relations, rels_out, rid, target_uri) -> None:
    """改写后重新解析：externalData 指向新外链且 autoUpdate=0、位置合乎 schema 顺序，其余整棵树与原文不变；
    根起始标签逐字不变；rels 里只有 rid 这一条关系变化。"""
    try:
        new_root = ET.fromstring(chart_out)
        new_rels = ET.fromstring(rels_out)
    except ET.ParseError:
        raise Skip("verify_failed")
    new_text = chart_out.decode("utf-8-sig")
    if _root_start(new_text).group(0) != _root_start(old_text).group(0):
        raise Skip("verify_failed")
    if _without_external_data(new_root) != _without_external_data(old_root):
        raise Skip("verify_failed")
    children = list(new_root)
    externals = [i for i, c in enumerate(children) if c.tag == f"{{{CHART_NS}}}externalData"]
    if len(externals) != 1:
        raise Skip("verify_failed")
    external = children[externals[0]]
    autos = [c for c in external if c.tag == f"{{{CHART_NS}}}autoUpdate"]
    if external.get(f"{{{REL_NS}}}id") != rid or len(autos) != 1 or autos[0].get("val") != "0":
        raise Skip("verify_failed")
    before = [_local(c.tag) for c in children[:externals[0]]]
    after = [_local(c.tag) for c in children[externals[0] + 1:]]
    if "chart" not in before or any(n not in INSERT_BEFORE for n in after) or \
            any(n in INSERT_BEFORE for n in before):
        raise Skip("verify_failed")
    relations = {r.get("Id"): dict(r.attrib) for r in new_rels}
    expected = {k: dict(v.attrib) for k, v in old_relations.items()}
    expected[rid] = {"Id": rid, "Type": OLE_LINK_TYPE, "Target": target_uri, "TargetMode": "External"}
    if relations != expected or len(list(new_rels)) != len(relations):
        raise Skip("verify_failed")


# ----------------------------- 整份文档 -----------------------------

def _content_types(archive: zipfile.ZipFile) -> Tuple[Dict[str, str], set]:
    data = _read(archive, "[Content_Types].xml")
    if data is None:
        raise Skip("no_content_types")
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        raise Skip("content_types_unparsable")
    overrides, defaults = {}, set()
    for item in root:
        if _local(item.tag) == "Override":
            overrides[(item.get("PartName") or "").lstrip("/")] = item.get("ContentType") or ""
        elif _local(item.tag) == "Default":
            defaults.add((item.get("Extension") or "").lower())
    return overrides, defaults


def relink_local_charts(src: Path, dst: Path, missing_target: Path) -> dict:
    """把 src 里可安全改写的本地图表改为外链到 missing_target（一个从不创建的本机路径），写出 dst。

    返回 {"changed", "document_skip", "charts": [{"part", "action", "reason"}], "counts"}；
    changed 为 False 时不写 dst，调用方用原文件。"""
    result = {"changed": False, "document_skip": None, "charts": [], "counts": {}}
    target_uri = Path(missing_target).resolve().as_uri()
    replacements: Dict[str, bytes] = {}
    try:
        with zipfile.ZipFile(src) as archive:
            infos = archive.infolist()
            names = [i.filename for i in infos]
            if len(infos) > MAX_ENTRIES:
                raise Skip("too_many_entries")
            if len(set(names)) != len(names):
                raise Skip("duplicate_entries")
            overrides, defaults = _content_types(archive)
            charts = sorted(p for p, kind in overrides.items() if kind == CHART_CONTENT_TYPE)
            for part, kind in sorted(overrides.items()):
                if kind.lower().endswith(CHARTEX_SUFFIX):
                    result["charts"].append({"part": part, "action": "skipped", "reason": "chartex"})
            if not charts:
                raise Skip("no_charts")
            reason = document_guard(archive, overrides)
            if reason:
                raise Skip(reason)
            for part in charts:
                entry = {"part": part, "action": "skipped", "reason": None}
                try:
                    chart_bytes = _read(archive, part)
                    if chart_bytes is None:
                        raise Skip("part_missing")
                    rels_name = _rels_name(part)
                    rels_bytes = _read(archive, rels_name)
                    if rels_bytes is None and "rels" not in defaults:
                        raise Skip("no_rels_content_type")
                    action, chart_out, rels_out = relink_chart(chart_bytes, rels_bytes, target_uri)
                    entry.update(action=action, reason=None)
                    if action != "linked":
                        replacements[part] = chart_out
                        replacements[rels_name] = rels_out
                except Skip as skip:
                    entry["reason"] = str(skip)
                result["charts"].append(entry)
            if replacements:
                _write(archive, infos, replacements, dst)
                result["changed"] = True
    except Skip as skip:
        result["document_skip"] = str(skip)
    except zipfile.BadZipFile:
        result["document_skip"] = "bad_zip"
    counts = Counter(c["action"] if c["action"] != "skipped" else f"skipped:{c['reason']}" for c in result["charts"])
    if result["document_skip"]:
        counts[f"document:{result['document_skip']}"] += 1
    result["counts"] = dict(counts)
    return result


def _write(archive: zipfile.ZipFile, infos, replacements: Dict[str, bytes], dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    pending = dict(replacements)
    with zipfile.ZipFile(dst, "w") as out:
        for info in infos:
            clone = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            clone.compress_type = info.compress_type
            clone.external_attr = info.external_attr
            if info.filename in pending:
                out.writestr(clone, pending.pop(info.filename))
                continue
            with archive.open(info) as reader, out.open(clone, "w") as writer:
                shutil.copyfileobj(reader, writer, 1 << 20)
        for name, data in pending.items():  # 原来没有 rels 的图表新增的 rels 部件
            clone = zipfile.ZipInfo(name, date_time=infos[0].date_time if infos else (1980, 1, 1, 0, 0, 0))
            clone.compress_type = zipfile.ZIP_DEFLATED
            out.writestr(clone, data)
