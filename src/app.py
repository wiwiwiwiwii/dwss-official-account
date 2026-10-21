from __future__ import annotations

import argparse
import ctypes
import importlib
import json
import math
import os
import re
from pathlib import Path
import secrets
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import unquote, urlsplit
import webbrowser

import maintenance


HERE = Path(__file__).resolve().parent
APP_VERSION = "v2.1.0"
PACKAGES = {"pywin32": "win32com.client", "beautifulsoup4": "bs4", "python-docx": "docx", "Pillow": "PIL.Image"}
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
MAX_FILE = 100 * 1024 * 1024
ITEM_TIMEOUT = 30 * 60
DEFAULT_TITLE_PREFIX = "【东吴商社】"
SETTINGS_PATH = Path(os.environ.get("APPDATA") or tempfile.gettempdir()) / "报告转公众号" / "settings.json"
REPORT_TYPES = ("点评", "深度", "周报")
AUTO_MODE = "auto"
AUTO_TYPE = "自动"
TITLE_TYPES = REPORT_TYPES + (AUTO_TYPE,)
STEP_LINE = re.compile(r"^\[STEP (\d+)/(\d+)\]\s*(.*)$")
UPDATE_REPO = "wiwiwiwiwii/dwss-official-account"
UPDATE_API = f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest"
UPDATE_PAGE = f"https://github.com/{UPDATE_REPO}/releases/latest"
VERSION_RE = re.compile(r"(\d+)(?:\.(\d+))?(?:\.(\d+))?")
PRERELEASE_RE = re.compile(r"alpha|beta|rc", re.IGNORECASE)


def version_tuple(text):
    """取版本号的可比较形态：`v0.13.0 beta` → (0, 13, 0, 0)，`v0.13.0` → (0, 13, 0, 1)。
    末位是“正式版”标记：同号正式版排在 beta 之后，beta 用户才会收到正式版提示。取不到返回 None。"""
    text = str(text or "")
    found = VERSION_RE.search(text)
    if not found:
        return None
    return tuple(int(part or 0) for part in found.groups()) + (0 if PRERELEASE_RE.search(text) else 1,)


def fetch_latest_release(timeout=3):
    """查 GitHub 最新 Release。只发一个不含任何文档内容的 GET；失败一律返回 None。"""
    from urllib.request import Request, urlopen

    try:
        request = Request(UPDATE_API, headers={"Accept": "application/vnd.github+json",
                                               "User-Agent": "dwss-official-account"})
        with urlopen(request, timeout=timeout) as response:
            data = json.loads(response.read(1 << 20).decode("utf-8", "replace"))
    except Exception:
        return None
    tag = str(data.get("tag_name") or "").strip()
    latest, current = version_tuple(tag), version_tuple(APP_VERSION)
    if not tag or not latest or not current or latest <= current:
        return None
    page = str(data.get("html_url") or "")
    return {"version": tag, "url": page if page.startswith("https://github.com/") else UPDATE_PAGE}


def compose_title(main_title, subtitle, report_type, prefix="【东吴商社】",
                  include_subtitle=True, subtitle_sep=" | ", include_type=True, type_sep="："):
    """标题装配。include_subtitle=False 是用户主动去副标题（保留前缀与类型），
    与「副标题数据缺失」的保守回退（裸报告标题）不同。"""
    main_title, subtitle = str(main_title).strip(), str(subtitle).strip()
    if not subtitle:
        return ""
    if include_subtitle:
        if not main_title:
            return subtitle
        ticker = r"(?:\d{4,6}(?:\.(?:SH|SZ|BJ|HK|SS))?|[A-Z]{1,5}(?:[.-][A-Z]{1,3})?)"
        main_title = re.sub(r"\s*[（(]\s*" + ticker + r"\s*[）)]", "", main_title).strip()
        main_title = re.sub(r"\s+\d{4,6}\.(?:SH|SZ|BJ|HK|SS)\s*$", "", main_title).strip()
        # 无括号无后缀的裸 A 股代码（如“安克创新 300866”）：限定 0/3/6/8/9 开头的六位，
        # 年份（1xxx/2xxx）和其他尾数不会被误删
        main_title = re.sub(r"\s+[03689]\d{5}\s*$", "", main_title).strip()
        if not main_title:
            return subtitle
    else:
        main_title = ""
    # 分字段查类型词：避免副标题末字与报告标题首字拼出并不存在的类型（如“新焦点”+“评估…”）
    has_type_word = report_type in main_title or report_type in subtitle
    type_text = report_type + type_sep if include_type and not has_type_word else ""
    if main_title:
        return prefix + main_title + subtitle_sep + type_text + subtitle
    return prefix + type_text + subtitle


def load_settings():
    try:
        data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {}
        return data
    except (OSError, ValueError):
        return {}


def load_title_prefix(settings=None):
    prefix = (load_settings() if settings is None else settings).get("title_prefix")
    if not single_line(prefix) or len(prefix) > 100:
        return DEFAULT_TITLE_PREFIX
    return prefix


def load_forecast_tables(settings=None):
    return (load_settings() if settings is None else settings).get("forecast_tables") is True


LINE_BREAK_CHARS = "\r\n\v\f\x1c\x1d\x1e\x85  "


def single_line(value):
    if not isinstance(value, str) or any(c in LINE_BREAK_CHARS for c in value):
        return False
    try:
        # 孤立代理项能通过换行检查却无法写成 UTF-8：在入口拒掉，
        # 否则要到落盘设置时才炸，那时任务线程已经起来了
        value.encode("utf-8")
    except UnicodeError:
        return False
    return True


def valid_sep(value, default):
    if not single_line(value) or len(value) > 10:
        return default
    return value


def load_title_options(settings=None):
    """标题装配四项设置（副标题/类型开关与各自分隔符），非法值回默认。"""
    data = load_settings() if settings is None else settings
    return {"include_subtitle": data.get("title_include_subtitle") is not False,
            "subtitle_sep": valid_sep(data.get("title_subtitle_sep"), " | "),
            "include_type": data.get("title_include_type") is not False,
            "type_sep": valid_sep(data.get("title_type_sep"), "：")}


def save_settings(prefix, forecast_tables, title_options=None):
    title_options = title_options or {}
    try:
        SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
        temp = SETTINGS_PATH.with_suffix(".json.tmp")
        payload = {"title_prefix": prefix, "forecast_tables": bool(forecast_tables)}
        if title_options:
            payload.update(title_include_subtitle=bool(title_options.get("include_subtitle", True)),
                           title_subtitle_sep=valid_sep(title_options.get("subtitle_sep"), " | "),
                           title_include_type=bool(title_options.get("include_type", True)),
                           title_type_sep=valid_sep(title_options.get("type_sep"), "："))
        temp.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        os.replace(temp, SETTINGS_PATH)
    except Exception:
        # 记忆设置是附带功能：无论何种失败都不得回传给调用方——此刻任务线程已经启动，
        # 抛出会让调用方误以为建任务失败而回滚状态并重复释放互斥锁
        pass


def kill_process_tree(pid):
    subprocess.run(["taskkill", "/T", "/F", "/PID", str(pid)], capture_output=True, creationflags=NO_WINDOW)


def find_word_pid(marker_token):
    """按 worker 打在 Word 主窗口上的标题标记实时定位其 WINWORD 进程号；没有则返回 None。

    kill 时刻现查窗口而不是存 PID：既不怕上报迟到，也不怕 PID 被回收复用。
    标记是工作台按任务项生成的随机值（经环境变量下发给 worker）。
    Word 打开文档后标题变成「文档名 - 标记」，所以枚举 OpusApp 窗口做包含匹配，
    不能用 FindWindow 精确匹配。"""
    user32 = ctypes.windll.user32
    marker = f"wechat-workbench-{marker_token}"
    found = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def visit(hwnd, _):
        cls = ctypes.create_unicode_buffer(64)
        user32.GetClassNameW(hwnd, cls, 64)
        if cls.value == "OpusApp":
            title = ctypes.create_unicode_buffer(512)
            user32.GetWindowTextW(hwnd, title, 512)
            if marker in title.value:
                pid = ctypes.c_ulong()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                if pid.value:
                    found.append(pid.value)
                    return False
        return True

    user32.EnumWindows(visit, None)
    return found[0] if found else None


def reap_word_instance(marker_token):
    """取消/超时后的后台收割：worker 已死，带其标记的 Word 窗口一出现就清理。

    Word 冷启动可能晚于 kill 数秒到十几秒才挂出窗口，同步等待会拖住取消响应，
    所以放到守护线程里扫最多 60 秒。标记是该任务项独有的随机值且 worker 已死，
    命中的窗口必属本程序。"""
    for _ in range(60):
        pid = find_word_pid(marker_token)
        if pid:
            kill_word_process(pid)
            time.sleep(1)
            if find_word_pid(marker_token) is None:
                return
        time.sleep(1)


def kill_word_process(pid):
    """只在该 PID 确认仍是 WINWORD 时结束它：绝不触碰用户自己打开的其他 Word。"""
    probe = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
                           capture_output=True, text=True, creationflags=NO_WINDOW)
    if "WINWORD.EXE" in (probe.stdout or "").upper():
        subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True, creationflags=NO_WINDOW)


# 路径内允许空格（如 Program Files）和中文段（如 工具软件）。空格后紧跟中文一般视为正文开始，
# 但其后同一词里还有路径分隔符（如 「2026 - 项目甲\报告.docx」）时仍属路径；
# 宁可多脱敏掉路径后紧跟的英文词，也不能泄露含空格路径的后半段
_PATH_CHARS = r"(?:[^\s<>:\"|?*\r\n'，。;；]|[ ](?![一-鿿])|[ ](?=[^\s\\/]*[\\/]))"
PATH_RE = re.compile(r"(?:\\\\\?\\)?[A-Za-z]:[\\/]" + _PATH_CHARS + r"*|\\\\" + _PATH_CHARS + "+")
# 堆栈与异常里带引号的路径整段脱敏（引号内可含任意空格与中文）
QUOTED_PATH_RE = re.compile(r"(?P<q>[\"'])(?:\\\\\?\\)?(?:[A-Za-z]:[\\/]|\\\\)(?:(?!(?P=q))[^\r\n])*(?P=q)")


# 已知失败形态 → 结果页上的一句原因与做法（按日志与错误信息匹配，先中先用）
FAILURE_HINTS = (
    (re.compile(r"Word\.Application\.\w+|0x800706BA|0x800706BE|0x80010108|-2147023174|-2147023170|-2147417848|"
                r"RPC 服务器不可用|远程过程调用失败|已与其客户端断开"),
     "Word 在转换途中退出或断开了连接。常见原因是内存加速、清理或安全类软件结束了后台的 Word；"
     "请暂停这类软件的加速或清理功能后重试。"),
    (re.compile(r"Could not locate summary content on page 1"),
     "第 1 页没有找到摘要（投资要点）。请确认报告使用东吴研报模板，且首页的投资要点不是空的。"),
    (re.compile(r"Invalid class string|无效的类字符串|-2147221005|-2147221164|没有注册类"),
     "没有找到可用的 Microsoft Word 桌面版，请确认已安装并激活（WPS 不能替代）。"),
    (re.compile(r"PermissionError|Permission denied|拒绝访问|被另一个程序使用|being used by another process"),
     "报告文件或临时目录被占用、无权访问。请关闭正在占用它的程序后重试。"),
)


EARLY_QUIT_NOTE = "Word 已提前退出"


def failure_hint(error, log=""):
    """先按最终错误分类（失败的真正原因）；分类不出时再看日志，已容错的告警与进度行（[warn]、->、
    「Word 已提前退出」）不算原因。"""
    log = "\n".join(line for line in str(log or "").splitlines()
                    if EARLY_QUIT_NOTE not in line and "[warn]" not in line and not line.lstrip().startswith("->"))
    for text in (str(error or ""), log):
        for pattern, hint in FAILURE_HINTS:
            if pattern.search(text):
                return hint
    return ""


# 报告里只列进程映像名，用于判断失败是否与这类常驻软件有关（不列路径、不列其他进程）
RESIDENT_TOOLS = re.compile(r"^(?:HipsTray|HipsDaemon|360Tray|360sd|ZhuDongFangYu|QQPCTray|QQPCRTP|kxetray|kxescore|"
                            r"LenovoPcManager\w*|LenovoVantage\w*|LeAppManager|SLBrowser)\.exe$", re.I)


def environment_lines():
    """错误报告的环境段：Word 版本、物理内存占用、Word 进程数、常驻的安全/管家类软件。取不到的项略去。"""
    lines = []
    if os.name != "nt":
        return lines
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CurVer") as key:
            word = winreg.QueryValue(key, None)
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration") as key:
                word += " · " + str(winreg.QueryValueEx(key, "VersionToReport")[0])
        except OSError:
            pass
        lines.append(f"Word：{word}")
    except OSError:
        pass
    try:
        class MemoryStatus(ctypes.Structure):
            _fields_ = [("length", ctypes.c_ulong), ("load", ctypes.c_ulong)] + [
                (name, ctypes.c_ulonglong) for name in
                ("total", "avail", "page_total", "page_avail", "virtual_total", "virtual_avail", "extended")]
        status = MemoryStatus()
        status.length = ctypes.sizeof(MemoryStatus)
        if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
            lines.append(f"内存（生成报告时）：占用 {status.load}%，共 {status.total / 2 ** 30:.0f} GB")
    except Exception:
        pass
    try:
        listing = subprocess.run(["tasklist", "/FO", "CSV", "/NH"], capture_output=True, timeout=10,
                                 creationflags=NO_WINDOW).stdout.decode("mbcs", "replace")
        names = [row.split('","')[0].strip('"') for row in listing.splitlines() if row.startswith('"')]
        lines.append(f"Word 进程数（生成报告时）：{sum(name.lower() == 'winword.exe' for name in names)}")
        tools = sorted({name for name in names if RESIDENT_TOOLS.match(name)})
        if tools:
            lines.append("常驻软件：" + "、".join(tools))
    except Exception:
        pass
    return lines


def build_error_report(job, item):
    """失败项的脱敏错误报告：只含定位 bug 所需的技术信息，绝不含用户报告内容。

    脱敏规则：文件名→<文件>；一切本机/网络路径→<本机路径>；含「」引文的行
    （自检点名的正文内容）整行剔除。日志仍只留内存，本报告由用户主动下载。"""
    def clean(text):
        lines = []
        for line in str(text or "").splitlines():
            if "「" in line or "」" in line:
                continue
            if item.get("name"):
                line = line.replace(item["name"], "<文件>")
                line = line.replace(Path(item["name"]).stem, "<文件>")
            line = QUOTED_PATH_RE.sub(lambda match: match["q"] + "<本机路径>" + match["q"], line)
            lines.append(PATH_RE.sub("<本机路径>", line))
        return lines
    import platform
    head = [
        "# 报告转公众号 错误报告（已脱敏，可直接转发开发者）",
        f"版本：{APP_VERSION}　打包：{bool(getattr(sys, 'frozen', False))}",
        f"系统：{platform.platform()}　Python：{sys.version.split()[0]}",
        f"时间：{time.strftime('%Y-%m-%d %H:%M:%S')}",
        f"模式：{','.join(job.get('modes', []))}　目录：{job.get('toc')}　渲染文末三表：{job.get('forecast_tables')}　标题类型：{job.get('title_type')}",
        f"任务项状态：{item.get('state')}　退出码：{item.get('exit_code')}",
    ] + environment_lines() + ([f"可能原因：{item['hint']}"] if item.get("hint") else []) + [
        "",
        "## 进度与日志（脱敏后）",
    ]
    body = clean("\n".join(item.get("log", [])))
    error = ["", "## 错误信息（脱敏后）"] + clean(item.get("error", "")) if item.get("error") else []
    return "\n".join(head + body + error) + "\n"


MAX_ZIP_ENTRIES = 5000
CHART_PART = re.compile(r"word/charts/chart\d+\.xml")


CENTRAL_HEADER_SIZE = 46
MAX_ZIP_RECORD_BOUND = 50000


def zip_entry_count(path):
    """在构造 ZipFile **之前**读出 EOCD 自报的条目数；读不出就返回 None。

    构造 ZipFile 会解析整个中央目录并为每个条目建 ZipInfo，条目极多的畸形包在那一步
    就能吃光内存，事后再数 infolist() 已经晚了。

    返回值是 EOCD 的自报值，**可能低报**，所以调用方仍须在构造后复查实际条目数；
    本函数负责的是让"构造"这一步本身有界：中央目录字节数 ÷ 单条记录最小长度是
    ZipFile 能建出多少 ZipInfo 的硬上界，越过硬线就直接返回 None、不再往下解析。
    该上界只当硬线用、不当条目数用——文件名等可变字段会让它系统性偏大。

    EOCD 记录须正好终止于文件末尾，且它指向的偏移处确实是中央目录，否则注释里的
    假签名就能冒充记录；任何 ZIP64 标记一律返回 None：真实 .docx 用不到。"""
    try:
        with open(path, "rb") as handle:
            handle.seek(0, os.SEEK_END)
            size = handle.tell()
            handle.seek(max(0, size - 66560))
            tail = handle.read()
            record, count, directory_bytes = None, 0, 0
            index = len(tail)
            while record is None:
                index = tail.rfind(b"PK\x05\x06", 0, index)
                if index < 0:
                    return None
                if len(tail) - index < 22:
                    continue
                candidate = tail[index:]
                # 注释长度必须让 EOCD 正好覆盖到文件末尾，否则这个签名是数据里的巧合
                if 22 + int.from_bytes(candidate[20:22], "little") != len(candidate):
                    continue
                count = int.from_bytes(candidate[10:12], "little")
                directory_bytes = int.from_bytes(candidate[12:16], "little")
                offset = int.from_bytes(candidate[16:20], "little")
                if count == 0xFFFF or directory_bytes == 0xFFFFFFFF or offset == 0xFFFFFFFF:
                    return None
                # 记录指向的中央目录必须真的在那儿：注释里自洽的假记录到这一步就出局
                if offset + directory_bytes + len(candidate) != size:
                    continue
                handle.seek(offset)
                if directory_bytes < CENTRAL_HEADER_SIZE or handle.read(4) != b"PK\x01\x02":
                    continue
                record = candidate
    except OSError:
        return None
    if directory_bytes // CENTRAL_HEADER_SIZE > MAX_ZIP_RECORD_BOUND:
        return None
    return count


def docx_stats(path):
    """预估“预计完成时间”用的轻量特征（仅 .docx；失败返回 None，前端只显示已运行时长）。

    全部取自压缩包，不调用 Word：图表数与其中「本地图表」数（数据不是链接到外部工作簿的图表：
    内嵌工作簿或没有数据）、内嵌对象数、docProps/app.xml 的页数，以及摘要的字数与它是否在摘要内容控件里
    （eta_summary_extent）。各步用哪些特征由拟合挑选（tests/perf-v117/refit_eta.py）。
    设有 zip 守卫（条目数、document.xml 与小部件的解压尺寸）防异常压缩比拖垮内存；
    调用方须在全局锁外执行本函数。"""
    try:
        import zipfile
        from xml.etree import ElementTree
        entries = zip_entry_count(path)
        if entries is None or entries > MAX_ZIP_ENTRIES:
            return None
        with zipfile.ZipFile(str(path)) as archive:
            infos = archive.infolist()
            if len(infos) > MAX_ZIP_ENTRIES:
                return None
            document = next((i for i in infos if i.filename == "word/document.xml"), None)
            if document is None or document.file_size > 100 * 1024 * 1024:
                return None
            media = [i for i in infos if i.filename.startswith("word/media/")]
            media_bytes = sum(i.file_size for i in media)
            # 按 ZipInfo 读而不是按名字读：同名重复条目时 read(name) 取的是最后一个，
            # 会绕过上面对第一个条目做的解压尺寸守卫
            xml = ElementTree.fromstring(archive.read(document))
            first = {}
            for info in infos:
                first.setdefault(info.filename, info)

            def small(name):
                info = first.get(name)
                return archive.read(info) if info is not None and info.file_size <= 1024 * 1024 else b""

            charts = [name for name in first if CHART_PART.fullmatch(name)]
            # 链接外部工作簿的图表：其关系文件里有 TargetMode="External" 的关系
            linked = sum(1 for name in charts
                         if b'TargetMode="External"' in small(name.replace("charts/", "charts/_rels/", 1) + ".rels"))
            pages = re.search(rb"<Pages>(\d{1,6})</Pages>", small("docProps/app.xml"))
        ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        summary_chars, summary_control = eta_summary_extent(xml)
        return {"size": path.stat().st_size, "media_count": len(media), "media_bytes": media_bytes,
                "summary_chars": summary_chars, "summary_control": summary_control,
                "paragraphs": len(list(xml.iter(ns + "p"))), "tables": len(list(xml.iter(ns + "tbl"))),
                "chars": sum(len(t.text or "") for t in xml.iter(ns + "t")),
                "charts": len(charts), "charts_local": len(charts) - linked,
                "embeddings": sum(1 for name in first if name.startswith("word/embeddings/")),
                "pages": int(pages.group(1)) if pages else None}
    except Exception:
        return None


# ---- 预计完成时间（ETA）：各步工作量模型 × 本机性能系数，转换中按已完成步骤向先验收缩修正 ----
# 算法见 tests/perf-v115/eta-quality.md；工作量系数、残差分位点、冷启动系数与步内外推开关由 tests/perf-v117/refit_eta.py
# 按批重拟合（tests/perf-v117/eta-refit.md），refit_apply.py 写进下面的常量。回放脚本直接调用 safe_item_eta / eta_learn / docx_stats。
LEGACY_ETA_PATH = SETTINGS_PATH.with_name("eta.json")   # beta 4 及以前写在程序外的系数文件，启动时删除
ETA_STREAM = "eta"           # 系数记录写在当前 EXE 的 NTFS 备用数据流 <exe>:eta；源码运行只记在内存
ETA_STREAM_LIMIT = 4096      # 数据流超过这个字节数视为损坏，按无记录处理
ETA_HISTORY = 10             # 保留最近几次转换的本机系数
ETA_K_BOUNDS = (0.2, 20.0)   # 本机系数的取值范围（超出的记录视为损坏）
ETA_RECENCY = 0.7            # 起算系数取加权中位数，每早一次的记录权重乘这个数
ETA_COLD_K = 0.9771          # 没有记录时的起算系数：工作量模型拟合数据里本机无负载各次系数的中位数
ETA_SHRINK = 40.0            # 在线修正时先验折合多少秒工作量的观测（n₀）
ETA_OUTLIER = 3.0            # 学习时排除与起算系数相差超过这个倍数的步骤
ETA_TIER_GROWTH = 1.3        # 当前步越过 P99 档后，每档再乘这个数
ETA_TIER_MARGIN = 0.5        # 离一档结束不到这么多秒就换下一档；预计值也不低于「已运行 + 这么多秒」
ETA_SUMMARY_CONTROLS = ("报告摘要", "AUTO_SAVE_Summary")   # 同 configs/summary.json 的 control_name_keywords
# 步内进度：worker 在摘要一步第一遍逐段判断结束时打印 `[SUB] summary 0/<总字数>`，第二遍逐段渲染每完成约 5% 字数
# 打印 `[SUB] summary <已渲染字数>/<总字数>`。按字数占比外推本步时长：已到 ETA_SUB_TRUST 的占比时完全按外推，
# 之前与模型中位数按占比几何加权，再分档。默认关闭：摘要提速后这一步只剩几秒，训练批次回放开与关几乎无差别，
# beta 6 开启时上修更多（tests/perf-v116/eta-quality.md、tests/perf-v117/eta-refit.md）。
ETA_SUB_LINE = re.compile(r"^\[SUB\] (summary) (\d{1,9})/(\d{1,9})\s*$")
ETA_SUB_ROLES = frozenset(("summary", "summary_only"))   # [SUB] summary 所在步骤的类型
ETA_SUB_TRUST = 0.2
ETA_SUB_ENABLED = False
# 工作量模型：本机（Core Ultra 7 258V 笔记本，接电，无负载）各步秒数，log W = b0 + Σ b·log1p(特征)。
# 数据：tests/batch-v117/survey.jsonl 的训练批次（摘要按格式分段渲染、正文导出前本地图表改为外链之后的实跑）。
# toc 与 toc_chars 沿用 beta 6：toc 代表 runs.json 与 batch-v115 两批运行条件之差，toc_chars 为指定点评模式含目录时的
# 目录图耗时（tests/perf-v116/eta-quality.md）；新数据全是工作台默认（自动、含目录），先扣除它们再拟合其余部分。
ETA_WORKLOAD = {
    "pre": (-1.1251, (('toc', 0.2329),)),
    "word": (0.18, (('toc', 0.826),)),
    "load": (-4.0, ()),
    "summary": (0.6628, (('toc', 0.9793),)),
    "summary_only": (0.8917, (('toc', 0.5703), ('toc_chars', 0.2187))),
    "body": (-0.787, (('pages', 0.3647), ('charts', 0.3915), ('toc', 1.9005))),
    "weekly": (1.554, ()),
    "merge": (-3.413, (('pages', 0.2727), ('charts', 0.505))),
}
# 各步类型「实际 ÷ 模型」的残差分位点（按报告留一），列依次对应 ETA_LEVELS
ETA_LEVELS = (0.5, 0.65, 0.8, 0.9, 0.95, 0.99)
ETA_RESIDUAL = {
    "pre": (1.0, 1.056, 1.127, 1.199, 1.262, 1.39),
    "word": (1.0, 1.039, 1.088, 1.137, 1.179, 1.262),
    "load": (1.0, 1.5, 2.0, 3.0, 5.0, 10.0),
    "summary": (1.007, 1.08, 1.173, 1.271, 1.358, 1.537),
    "summary_only": (1.007, 1.072, 1.155, 1.242, 1.318, 1.474),
    "body": (1.004, 1.053, 1.113, 1.174, 1.227, 1.334),
    "weekly": (1.0, 1.043, 1.096, 1.15, 1.196, 1.289),
    "merge": (0.983, 1.126, 1.324, 1.547, 1.759, 2.239),
}
ETA_TIER_LEVELS = (0.5, 0.65, 0.8, 0.9, 0.95, 0.99)   # 当前步超时时依次取用的档（之后每档 ×ETA_TIER_GROWTH）
ETA_HIGH_LEVEL = 0.95        # 范围上限：当前步取这一分位的档（主数字已到这一档时取下一档）
ETA_FUTURE_HIGH_LEVEL = 0.8  # 范围上限：后续各步取这一分位
ETA_ROLE_TYPE = {"forecast": "body"}    # 点评附三表多跑的那一遍就是正文管线
ETA_CONVERT = frozenset(("summary", "summary_only", "body", "weekly", "merge"))   # 参与本机系数的步骤类型
# 进度条段内曲线 f(x)，x = 本段已耗时 ÷ 本段预估：拐点前线性（按时推进时进度≈耗时比例），拐点后先以
# 一阶连续的指数软着陆到 MID，再以长尺度指数缓慢逼近 CAP（严重超时也持续可见地爬升，永不触顶）
ETA_BAR_KNEE = 0.85
ETA_BAR_SOFT = 0.12              # 软着陆段的尺度＝幅度，故拐点处斜率为 1
ETA_BAR_MID = 0.86
ETA_BAR_CAP = 0.97
ETA_BAR_TAIL = 5.0               # 末段指数的尺度（以 x 计）


def eta_summary_extent(root):
    """（摘要字数, 是否在摘要内容控件里 1/0），root 为 document.xml 的根元素。

    与 converter.SummaryExtractor 找摘要的顺序一致：先找标题或标记含 ETA_SUMMARY_CONTROLS 的内容控件
    （跳过只显示占位文字的），取其中的字数；找不到时转换器退回第 1 页，这里取 Word 上次保存时记下的
    第一个分页处（lastRenderedPageBreak 或手动分页符）之前的字数，没有分页记录时取全文字数。"""
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for sdt in root.iter(ns + "sdt"):
        props = sdt.find(ns + "sdtPr")
        content = sdt.find(ns + "sdtContent")
        if props is None or content is None or props.find(ns + "showingPlcHdr") is not None:
            continue
        label = " ".join(child.get(ns + "val") or "" for child in props if child.tag in (ns + "alias", ns + "tag"))
        label = re.sub(r"\s+", " ", label)
        if any(key in label for key in ETA_SUMMARY_CONTROLS):
            return sum(len(t.text or "") for t in content.iter(ns + "t")), 1
    chars = 0
    body = root.find(ns + "body")
    for element in (body.iter() if body is not None else ()):
        if element.tag == ns + "lastRenderedPageBreak" or (element.tag == ns + "br" and element.get(ns + "type") == "page"):
            break
        if element.tag == ns + "t":
            chars += len(element.text or "")
    return chars, 0


def eta_features(stats, toc=True):
    """工作量模型的特征（docx_stats 的结果换算）；拿不到或取值异常返回 None（只显示已运行）。"""
    if not isinstance(stats, dict):
        return None

    def number(key, required=False):
        value = stats.get(key)
        if value is None and not required:
            return 0.0
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value < 0:
            raise ValueError(key)
        return float(value)

    try:
        size, media_bytes, chars = number("size", True), number("media_bytes", True), number("chars", True)
        return {"text_mb": max(0.0, size - media_bytes) / 1e6, "media": number("media_count"),
                "media_mb": media_bytes / 1e6, "paragraphs": number("paragraphs"), "tables": number("tables"),
                "chars": chars / 1e4, "pages": number("pages"), "charts": number("charts"),
                "charts_local": number("charts_local"), "embeddings": number("embeddings"),
                "summary_chars": number("summary_chars") / 100, "summary_control": number("summary_control"),
                "toc": 1.0 if toc else 0.0, "toc_chars": chars / 1e4 if toc else 0.0}
    except ValueError:
        return None


def eta_roles(modes):
    """各段的步骤：段 0 = 任务开始→[STEP 1]（复制、启动 worker、识别），段 1 启动 Word，段 2 载入版式，
    之后每个模式一段（顺序同 convert_worker.MODES），点评附三表多一段正文管线，末段合并写出。"""
    modes = list(modes or [])
    summary = "summary" if "body" in modes else "summary_only"
    return (["pre", "word", "load"] + [summary if m == "summary" else m for m in ("summary", "body", "weekly") if m in modes]
            + (["forecast"] if "forecast" in modes else []) + ["merge"])


def eta_workload(role, features):
    """第 role 步在本机无负载时的预估秒数 W_s(x)。"""
    if role == "forecast":      # 三表那一遍按正文管线转换，但不出目录图
        features = dict(features, toc=0.0, toc_chars=0.0)
    intercept, coefs = ETA_WORKLOAD[ETA_ROLE_TYPE.get(role, role)]
    value = intercept + sum(b * math.log1p(features[name]) for name, b in coefs)
    return min(3600.0, max(0.01, math.exp(value)))


def eta_quantile(role, level):
    """第 role 步类型「实际 ÷ 模型」的 level 分位点（level 须是 ETA_LEVELS 之一）。"""
    return ETA_RESIDUAL[ETA_ROLE_TYPE.get(role, role)][ETA_LEVELS.index(level)]


def eta_empty_store():
    return []


def eta_prior_k(store):
    """起算系数 k₀：最近几次本机系数的加权中位数（越近权重越大）；没有记录时为 ETA_COLD_K。"""
    values = [v for v in store or [] if isinstance(v, float) and ETA_K_BOUNDS[0] <= v <= ETA_K_BOUNDS[1]]
    if not values:
        return ETA_COLD_K
    weighted = sorted((v, ETA_RECENCY ** age) for age, v in enumerate(reversed(values)))
    half, acc = sum(w for _, w in weighted) / 2, 0.0
    for value, weight in weighted:
        acc += weight
        if acc >= half:
            return value
    return weighted[-1][0]


def eta_arrivals(item, started):
    """[STEP i] 的到达时刻（自任务开始计）。按步号取首次到达：自动识别要用 Word 时总步数会中途改变，
    步号不变。"""
    marks = {}
    for step, _total, moment in item.get("step_times") or []:
        marks.setdefault(int(step), float(moment) - started)
    arrivals = []
    while len(arrivals) + 1 in marks:
        arrivals.append(max(marks[len(arrivals) + 1], arrivals[-1] if arrivals else 0.0))
    return arrivals


def eta_current_k(k0, roles, works, bounds, partial=None):
    """按已完成转换步的工作量加权向先验收缩：log k = (n₀·log k₀ + Σ W_j·log(T_j/W_j)) / (n₀ + Σ W_j)。
    partial =（当前步序号, 已完成占比 f, 外推的本步时长 D）：当前步按 f·W 的权重计入 D / W（步内进度）。"""
    num, den = ETA_SHRINK * math.log(k0), ETA_SHRINK
    observed = [(j, 1.0, max(bounds[j + 1] - bounds[j], 0.01)) for j in range(len(bounds) - 1)]
    if partial is not None:
        observed.append(partial)
    for j, share, spent in observed:
        if ETA_ROLE_TYPE.get(roles[j], roles[j]) in ETA_CONVERT:
            num += share * works[j] * math.log(spent / works[j])
            den += share * works[j]
    low, high = ETA_K_BOUNDS
    return min(high, max(low, math.exp(num / den)))


def eta_tiers(base, role, spent):
    """当前步的预计时长：按残差分位点分档（ETA_TIER_LEVELS，之后每档 ×ETA_TIER_GROWTH），取第一个比已耗时
    长出 ETA_TIER_MARGIN 秒以上的档——档与档之间预计值不随时钟走。返回（这一档, 范围上限用的档）。"""
    tiers = [base * eta_quantile(role, level) for level in ETA_TIER_LEVELS]
    while tiers[-1] <= spent + ETA_TIER_MARGIN:
        tiers.append(tiers[-1] * ETA_TIER_GROWTH)
    index = next(i for i, t in enumerate(tiers) if t > spent + ETA_TIER_MARGIN)
    upper = tiers[index + 1] if index + 1 < len(tiers) else tiers[index] * ETA_TIER_GROWTH
    high = base * eta_quantile(role, ETA_HIGH_LEVEL)
    return tiers[index], (high if high > tiers[index] else upper)


def eta_record_sub(item, match, moment):
    """记下一行 [SUB]（ETA_SUB_LINE 的匹配）：当前 [STEP] 步号、总字数、第一遍结束（0/总字数）时刻、最近一次的已完成字数与时刻。
    换了步骤或总字数就重新开始；已完成字数只增不减。调用方持有锁。"""
    done, total = int(match.group(2)), int(match.group(3))
    step = int((item.get("progress") or {}).get("step") or 0)
    if total <= 0 or done > total:
        return
    sub = item.get("sub_progress")
    if not isinstance(sub, dict) or sub.get("step") != step or sub.get("total") != total:
        sub = item["sub_progress"] = {"step": step, "total": total, "zero": None, "done": 0, "at": None}
    if done == 0:
        sub["zero"] = sub["zero"] if sub["zero"] is not None else moment
    elif done >= sub["done"]:
        sub["done"], sub["at"] = done, moment


def eta_sub_point(item, started):
    """当前步内进度（步号, 第一遍结束时刻或 None, 已完成占比, 这一占比的时刻），时刻自任务开始计；没有返回 None。"""
    sub = item.get("sub_progress")
    if not isinstance(sub, dict) or not sub.get("at") or not sub.get("total"):
        return None
    zero = sub.get("zero")
    return (int(sub["step"]), None if zero is None else float(zero) - started,
            min(1.0, sub["done"] / sub["total"]), float(sub["at"]) - started)


def eta_sub_extrapolate(sub, current, role, start):
    """按当前步的 [SUB] 进度外推本步时长：（外推时长 D, 已完成占比 f）；不适用返回 None。
    第一遍结束前的部分按实耗计，第二遍按字数占比线性外推：D = (t₀ − 本步开始) + (t − t₀) / f。"""
    if sub is None or sub[0] != current or ETA_ROLE_TYPE.get(role, role) not in ETA_SUB_ROLES:
        return None
    _, zero, share, moment = sub
    origin = zero if zero is not None and start <= zero <= moment else start
    if share <= 0 or moment <= origin:
        return None
    return (origin - start) + (moment - origin) / share, share


def eta_estimate(k0, roles, works, arrivals, elapsed, sub=None):
    """返回（预计总耗时 P50, 范围上限, 进度条原始值 0~1），都自任务开始计。

    已完成各步按实际耗时；本机系数 k 由已完成的转换步向起算系数 k₀ 收缩更新（启动类步骤不参与：
    实测 Word 启动快慢与后续转换无关）；当前步超出预估时归因于这一步工作量估低，不调 k、不放大后续
    步骤，按残差分位点分档（eta_tiers）；未开始各步按 k × W × P50。
    范围上限：当前步取 ETA_HIGH_LEVEL 分位的档（主数字已到这一档时取下一档），后续各步取 ETA_FUTURE_HIGH_LEVEL。
    进度条：各步水位按本报告起算时的各步预估（k₀ × W × P50）的累计占比，不随修正变化；
    步内为水位 + 占比 × f(已耗时 ÷ 本步预估)，f 见 eta_bar_curve，封顶在 1 之下，完成由页面推到 100%。"""
    last = len(works) - 1
    arrivals = [a for a in arrivals if a <= elapsed][:last]
    bounds = [0.0] + arrivals
    current = len(arrivals)
    start = bounds[-1]
    spent = max(0.0, elapsed - start)
    if sub is not None and sub[3] > elapsed:
        sub = None
    extrapolated = eta_sub_extrapolate(sub, current, roles[current], start) if ETA_SUB_ENABLED else None
    partial = (current, extrapolated[1], extrapolated[0]) if extrapolated else None
    k = eta_current_k(k0, roles, works, bounds, partial)
    base = k * works[current]
    tier, high_tier = eta_tiers(base, roles[current], spent)
    if extrapolated:
        # 步内进度：按字数占比 f 把本步的模型中位数几何加权地换成外推值（f ≥ ETA_SUB_TRUST 时完全按外推），
        # 再以加权结果为中位数分档。加权的两端在两行 [SUB] 之间都不变，档与档之间预计值不随时钟走
        span, share = extrapolated
        weight = min(1.0, share / ETA_SUB_TRUST)
        median = math.exp((1 - weight) * math.log(base * eta_quantile(roles[current], 0.5)) + weight * math.log(span))
        tier, high_tier = eta_tiers(median / eta_quantile(roles[current], 0.5), roles[current], spent)
    future = sum(k * works[j] * eta_quantile(roles[j], 0.5) for j in range(current + 1, last + 1))
    future_high = sum(k * works[j] * eta_quantile(roles[j], ETA_FUTURE_HIGH_LEVEL) for j in range(current + 1, last + 1))
    total = max(start + tier + future, elapsed + ETA_TIER_MARGIN)
    high = max(start + high_tier + future_high, total)
    plan = [k0 * w * eta_quantile(r, 0.5) for r, w in zip(roles, works)]
    whole = sum(plan)
    level = sum(plan[:current]) / whole
    expected = median if extrapolated else base * eta_quantile(roles[current], 0.5)
    climb = plan[current] / whole * eta_bar_curve(spent / expected)
    return total, high, min(level + climb, 0.999)


def eta_bar_curve(x):
    """段内进度 f(x) ∈ [0, ETA_BAR_CAP)：x ≤ 拐点时 f = x；之后一阶连续地转入两段指数。"""
    knee, soft, mid = ETA_BAR_KNEE, ETA_BAR_SOFT, ETA_BAR_MID
    if x <= knee:
        return max(0.0, x)
    turn = knee - soft * math.log(1 - (mid - knee) / soft)     # 软着陆段到达 MID 的 x
    if x <= turn:
        return knee + soft * (1 - math.exp(-(x - knee) / soft))
    return mid + (ETA_BAR_CAP - mid) * (1 - math.exp(-(x - turn) / ETA_BAR_TAIL))


def eta_job_plan(item, job_modes=None):
    """（各段步骤, 各段工作量）；无法预估返回 None。[modes] 行到达前退回任务选定的模式。
    自动模式下点评框架不出目录图（convert_worker.apply_detection 把目录开关限定在有正文的模式），
    「仅摘要」一步不计目录图（toc_chars 取 0）；指定点评模式含目录时才在这一步出目录图。toc 本身与其余各步
    同口径，代表拟合数据两批运行条件之差（见 ETA_WORKLOAD 的注释），照常取任务的目录开关。"""
    modes = item.get("detected_modes") or ([m for m in job_modes or [] if m != AUTO_MODE] or None)
    features = eta_features(item.get("stats"), item.get("toc", True))
    if not modes or features is None:
        return None
    roles = eta_roles(modes)
    no_toc = dict(features, toc_chars=0.0)
    auto = AUTO_MODE in (job_modes or ())
    return roles, [eta_workload(role, no_toc if auto and role == "summary_only" else features) for role in roles]


def safe_item_eta(item, now, store, job_modes=None):
    """轮询接口用的 item_eta：预估是附带功能，任何异常都只放弃这次预估（页面保持已显示的进度，
    尚无进度时显示不确定态），不得让 api/jobs 报错——那会让结果页收不到进度和结果。"""
    try:
        return item_eta(item, now, store, job_modes)
    except Exception:
        return None, None, None


def item_eta(item, now, store, job_modes=None):
    """结果页信息条用的（预计总耗时整秒, 进度条原始值 0~1, 范围上限整秒）；无法预估时全为 None
    （页面只显示已运行；进度条尚无进度时为不确定态，已有进度时保持已显示的值）。时刻全部取服务端时钟。"""
    plan = eta_job_plan(item, job_modes)
    started = item.get("started")
    if not plan or not started:
        return None, None, None
    roles, works = plan
    total, high, progress = eta_estimate(eta_prior_k(store), roles, works, eta_arrivals(item, started),
                                         max(0.0, now - started), eta_sub_point(item, started))
    return math.ceil(total), round(progress, 4), math.ceil(high)


def eta_learn(store, item, finished_at, job_modes=None):
    """一次成功转换后的本机系数（返回新的记录列表，只留最近 ETA_HISTORY 次）；无法计算时原样返回。

    本次系数 = 各转换步「实际 ÷ 工作量」按工作量加权的几何均值；与起算系数相差超过 ETA_OUTLIER 倍的
    步骤不参与（一份报告特异的步骤不带偏本机系数）。全部步骤都偏向同一侧时视为本机变了，按截到
    ETA_OUTLIER 倍计入，几次之内跟上。"""
    plan = eta_job_plan(item, job_modes)
    started = item.get("started")
    if not plan or not started or not isinstance(finished_at, (int, float)):
        return store
    roles, works = plan
    arrivals = eta_arrivals(item, started)
    if len(arrivals) != len(roles) - 1:
        return store
    bounds = [0.0] + arrivals + [max(finished_at - started, arrivals[-1] if arrivals else 0.0)]
    k0 = eta_prior_k(store)
    observed = [(works[j], max(bounds[j + 1] - bounds[j], 0.01) / works[j]) for j in range(len(roles))
                if ETA_ROLE_TYPE.get(roles[j], roles[j]) in ETA_CONVERT]
    kept = [(w, r) for w, r in observed if abs(math.log(r / k0)) <= math.log(ETA_OUTLIER)]
    if not kept and observed and (all(r > k0 for _, r in observed) or all(r < k0 for _, r in observed)):
        kept = [(w, min(k0 * ETA_OUTLIER, max(k0 / ETA_OUTLIER, r))) for w, r in observed]
    if not kept:
        return store
    low, high = ETA_K_BOUNDS
    k = math.exp(sum(w * math.log(r) for w, r in kept) / sum(w for w, _ in kept))
    return (list(store or []) + [round(min(high, max(low, k)), 4)])[-ETA_HISTORY:]


def eta_stream_path():
    """打包版：当前 EXE 的备用数据流；源码运行返回 None（绝不写 python.exe 的数据流）。"""
    if not getattr(sys, "frozen", False):
        return None
    return sys.executable + ":" + ETA_STREAM


def load_eta_store():
    path = eta_stream_path()
    if not path:
        return []
    try:
        with open(path, "rb") as handle:
            raw = handle.read(ETA_STREAM_LIMIT + 1)
        if len(raw) > ETA_STREAM_LIMIT:
            return []
        data = json.loads(raw.decode("utf-8"))
    except (OSError, ValueError, RecursionError):
        return []
    values = data.get("k") if isinstance(data, dict) else None
    if not isinstance(values, list):
        return []
    low, high = ETA_K_BOUNDS
    return [float(v) for v in values if isinstance(v, (int, float)) and not isinstance(v, bool)
            and math.isfinite(v) and low <= v <= high][-ETA_HISTORY:]


def save_eta_store(store):
    path = eta_stream_path()
    if not path:
        return
    try:
        payload = json.dumps({"k": [round(float(v), 4) for v in list(store)[-ETA_HISTORY:]]}).encode("utf-8")
        with open(path, "wb") as handle:
            handle.write(payload)
    except Exception:
        # 附带功能：EXE 所在盘不支持数据流、只读或被占用时，只影响下一次的起算系数
        pass


def remove_legacy_eta_file():
    """删除 beta 4 及以前写在 %APPDATA% 的 eta.json（只删这一个文件，settings.json 保留）。"""
    try:
        LEGACY_ETA_PATH.unlink()
    except OSError:
        pass


def python_exe():
    exe = Path(sys.executable)
    return str(exe.with_name("python.exe") if exe.name.lower() == "pythonw.exe" else exe)


def accessible_path(raw):
    path = str(Path(raw).absolute())
    if os.name == "nt" and not path.startswith("\\\\?\\"):
        path = "\\\\?\\UNC\\" + path[2:] if path.startswith("\\\\") else "\\\\?\\" + path
    return Path(path)


def process_alive(pid):
    if os.name != "nt":
        try:
            os.kill(pid, 0)
            return True
        except ProcessLookupError:
            return False
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.OpenProcess.restype = ctypes.c_void_p
    kernel.GetExitCodeProcess.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
    kernel.CloseHandle.argtypes = [ctypes.c_void_p]
    handle = kernel.OpenProcess(0x1000, False, pid)
    if not handle:
        return ctypes.get_last_error() != 87
    try:
        code = ctypes.c_ulong()
        return not kernel.GetExitCodeProcess(handle, ctypes.byref(code)) or code.value == 259
    finally:
        kernel.CloseHandle(handle)


def clean_orphans():
    root = Path(tempfile.gettempdir()).resolve()
    for candidate in root.glob("wechat-workbench-*"):
        if candidate.is_symlink() or not candidate.is_dir() or candidate.resolve().parent != root:
            continue
        try:
            owner = json.loads((candidate / "owner.json").read_text(encoding="utf-8"))
            if owner.get("app") == "wechat-workbench" and not process_alive(int(owner["pid"])):
                shutil.rmtree(candidate)
        except (OSError, ValueError, KeyError):
            pass


class Workbench:
    def __init__(self, engine_root):
        self.engine_root = engine_root
        self.token = secrets.token_urlsafe(24)
        self.lock = threading.RLock()
        self.operation = threading.Lock()
        self.jobs = {}
        self.last_job = None
        self.uploads = {}
        self.clients = {}
        self.last_activity = time.monotonic()
        self.install = {"state": "idle", "phase": "idle", "source": "", "log": ""}
        self.dependencies = []
        self.dep_time = 0
        self.closing = False
        self.server = None
        self.update = None      # 后台查到的新版本；查不到或已是最新时保持 None
        self.eta_store = None   # 最近几次的本机性能系数（打包版存于 EXE 的备用数据流），首次用到时载入

    def check_dependencies(self):
        with self.lock:
            if time.monotonic() - self.dep_time < 15 and self.dependencies:
                return self.dependencies
        bundled = bool(getattr(sys, "frozen", False))
        rows = [{"name": "Python", "ok": True, "detail": ("内置运行环境 · " if bundled else "") + sys.version.split()[0], "installable": False}]
        for name, module in PACKAGES.items():
            try:
                importlib.import_module(module)
                ok, detail = True, "程序内置，无需安装" if bundled else "已安装"
            except Exception as exc:
                ok, detail = False, str(exc)
            rows.append({"name": name, "ok": ok, "detail": detail, "installable": not bundled})
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
                pass
            word_ok = True
        except OSError:
            word_ok = False
        rows.append({"name": "Microsoft Word 桌面版", "ok": word_ok,
                     "detail": "已注册，转换时验证可用性" if word_ok else "请安装并激活桌面版 Word，WPS 不能替代此引擎",
                     "installable": False})
        try:
            import converter
            for mode in ("summary", "body", "weekly"):
                converter.load_config(mode, self.engine_root / "configs")
            if not (self.engine_root / "assets" / "weekly-heading.png").is_file():
                raise RuntimeError("缺少内置周报标题背景，请恢复完整程序文件。")
            core_ok, core_detail = True, "摘要、深度、周报共用核心，保留三套版式"
        except Exception as exc:
            core_ok, core_detail = False, str(exc)
        rows.append({"name": "转换核心与版式", "ok": core_ok, "detail": core_detail, "installable": False})
        if bundled:
            # 打包版对用户唯一有信息量的是 Word 前置；内置项只在异常（包损坏）时才显示
            rows = [r for r in rows if r["name"] == "Microsoft Word 桌面版" or not r["ok"]]
        with self.lock:
            self.dependencies, self.dep_time = rows, time.monotonic()
        return rows

    def append_install_log(self, line):
        with self.lock:
            self.install["log"] = (self.install["log"] + line + "\n")[-20000:]

    def run_install_command(self, args, phase, source):
        with self.lock:
            self.install.update(phase=phase, source=source)
        process = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                   text=True, encoding="utf-8", errors="replace", creationflags=NO_WINDOW,
                                   env=dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1"))
        for line in process.stdout:
            self.append_install_log(line.rstrip())
        return process.wait()

    def install_missing(self):
        try:
            with self.lock:
                self.install.update(phase="preparing", source="")
            missing = [r["name"] for r in self.check_dependencies() if r["installable"] and not r["ok"]]
            if not missing:
                self.append_install_log("所有 Python 依赖已就绪，无需安装。")
            else:
                probe = subprocess.run([python_exe(), "-m", "pip", "--version"], capture_output=True, creationflags=NO_WINDOW)
                if probe.returncode:
                    subprocess.run([python_exe(), "-m", "ensurepip", "--user"], check=True, capture_output=True, creationflags=NO_WINDOW)
                for source, index in [("PyPI 官方源", "https://pypi.org/simple"), ("清华大学国内镜像", "https://pypi.tuna.tsinghua.edu.cn/simple")]:
                    with tempfile.TemporaryDirectory(prefix="wechat-dependencies-") as directory:
                        self.append_install_log("正在使用 " + source + " 下载：" + "、".join(missing))
                        args = [python_exe(), "-m", "pip", "download", "--dest", directory, "--disable-pip-version-check", "--progress-bar", "off", "--timeout", "15", "--retries", "1", "--index-url", index] + missing
                        if self.run_install_command(args, "downloading", source) != 0:
                            continue
                        self.append_install_log("下载完成，正在安装：" + "、".join(missing))
                        args = [python_exe(), "-m", "pip", "install", "--user", "--no-index", "--find-links", directory, "--disable-pip-version-check", "--no-warn-script-location"] + missing
                        if self.run_install_command(args, "installing", source) == 0:
                            break
                else:
                    raise RuntimeError("官方源和国内镜像均安装失败，请查看上方日志后重试。")
                importlib.invalidate_caches()
                import site
                site.addsitedir(site.getusersitepackages())
                with self.lock:
                    self.dep_time = 0
                failed = [r["name"] for r in self.check_dependencies() if r["installable"] and not r["ok"]]
                if failed:
                    raise RuntimeError("安装后仍无法导入：" + "、".join(failed) + "。请退出后重新启动工作台。")
            with self.lock:
                self.install.update(state="done", phase="done")
        except Exception as exc:
            self.append_install_log(str(exc))
            with self.lock:
                self.install.update(state="error", phase="error")
        finally:
            self.operation.release()

    def create_job(self, request):
        modes = request.get("modes")
        if not isinstance(modes, list) or not modes:
            raise ValueError("请选择有效转换模式。")
        auto = modes == [AUTO_MODE]
        if not auto:
            if any(m not in ("summary", "body", "weekly") for m in modes):
                raise ValueError("请选择有效转换模式。")
            if "body" in modes and "weekly" in modes:
                raise ValueError("深度正文与周报不能同时选择。")
        report_type = request.get("report_type", AUTO_TYPE if auto else "周报" if "weekly" in modes else "深度" if "body" in modes else "点评")
        title_type = request.get("title_type", report_type)
        title_prefix = request.get("title_prefix", DEFAULT_TITLE_PREFIX)
        if not single_line(title_prefix) or len(title_prefix) > 100:
            raise ValueError("标题前缀须为100字以内的单行文字。")
        if "forecast_tables" in request and not isinstance(request["forecast_tables"], bool):
            raise ValueError("盈利预测三表开关必须是布尔值。")
        if "toc" in request and not isinstance(request["toc"], bool):
            raise ValueError("内容目录开关必须是布尔值。")
        for key in ("title_include_subtitle", "title_include_type"):
            if key in request and not isinstance(request[key], bool):
                raise ValueError("标题开关必须是布尔值。")
        for key in ("title_subtitle_sep", "title_type_sep"):
            if key in request and (not single_line(request[key]) or len(request[key]) > 10):
                raise ValueError("标题分隔符须为10字以内的单行文字。")
        title_options = {"include_subtitle": request.get("title_include_subtitle", True) is not False,
                         "subtitle_sep": request.get("title_subtitle_sep", " | "),
                         "include_type": request.get("title_include_type", True) is not False,
                         "type_sep": request.get("title_type_sep", "：")}
        if report_type not in TITLE_TYPES or title_type not in TITLE_TYPES:
            raise ValueError("请选择自动、点评、深度或周报标题样式。")
        if title_type == AUTO_TYPE and not auto:
            title_type = "周报" if "weekly" in modes else "深度" if "body" in modes else "点评"
        inputs = request.get("inputs")
        if not isinstance(inputs, list) or not 1 <= len(inputs) <= 30:
            raise ValueError("每批请选择 1 至 30 个 Word 文件。")
        items = []
        for entry in inputs:
            if not isinstance(entry, dict):
                raise ValueError("文件输入格式不正确。")
            if "upload" in entry:
                upload = self.uploads.get(entry["upload"])
                if not upload:
                    raise ValueError("上传已过期，请重新添加文件。")
                items.append({"name": upload["name"], "state": "queued", "upload": entry["upload"], "content": upload["content"]})
            else:
                raw = entry.get("path", "")
                if not isinstance(raw, str):
                    raise ValueError("路径必须是文字。")
                path = Path(raw.strip().strip('"'))
                if not path.is_absolute() or path.suffix.lower() not in (".doc", ".docx"):
                    raise ValueError("找不到 Word 文件，请检查完整路径：" + str(path))
                path = accessible_path(path)
                if not path.is_file():
                    raise ValueError("找不到 Word 文件，请检查完整路径：" + str(path))
                if path.stat().st_size > MAX_FILE:
                    raise ValueError("单个文件不能超过 100 MB：" + path.name)
                items.append({"name": path.name, "path": str(path), "state": "queued"})
        if not self.operation.acquire(blocking=False):
            raise ValueError("正在转换或安装依赖，请等待完成。")
        job_id = secrets.token_hex(12)
        previous_job = self.last_job
        try:
            job = {"id": job_id, "state": "queued", "items": items, "modes": list(dict.fromkeys(modes)),
                   "toc": bool(request.get("toc")), "forecast_tables": request.get("forecast_tables") is True,
                   "title_type": title_type, "title_prefix": title_prefix.strip(), "title_options": title_options,
                   "expires": time.monotonic() + 7200, "viewers": {}}
            with self.lock:
                self.jobs[job_id] = job
                self.last_job = job_id
            threading.Thread(target=self.run_job, args=(job,), daemon=True).start()
            # 任务确实开跑才记住设置：被拒绝或失败的提交不覆盖既有设置
            save_settings(title_prefix.strip(), job["forecast_tables"], title_options)
            return job_id
        except Exception:
            with self.lock:
                self.jobs.pop(job_id, None)
                # 恢复上一任务指针：新任务没能起来，主页的“返回结果页”仍应指向旧结果
                if self.last_job == job_id:
                    self.last_job = previous_job
            self.operation.release()
            raise

    def run_job(self, job):
        try:
            with self.lock:
                job["state"] = "running"
            for item in job["items"]:
                with self.lock:
                    if job.get("cancel") or job.get("unkillable"):
                        item.pop("content", None)
                        item.update(state="cancelled")
                        continue
                    item.update(state="running", started=time.time(), log=[], exit_code=None, step_times=[],
                                progress={"step": 0, "total": 0, "text": "准备文件"}, toc=bool(job["toc"]))
                try:
                    # ignore_cleanup_errors：取消/超时刚杀掉 Word 时文件可能仍被短暂占用，
                    # 残留目录由下次启动的 clean_orphans 收走，不把取消变成报错
                    with tempfile.TemporaryDirectory(prefix="wechat-workbench-", ignore_cleanup_errors=True) as directory:
                        temp = Path(directory)
                        (temp / "owner.json").write_text(json.dumps({"app": "wechat-workbench", "pid": os.getpid()}), encoding="utf-8")
                        source = temp / ("input" + Path(item["name"]).suffix.lower())
                        if "upload" in item:
                            source.write_bytes(item["content"])
                            with self.lock:
                                item.pop("content", None)
                        else:
                            # 分块复制并检查取消：网络路径卡死时取消不至于永远等在这一步
                            staging_cancelled = False
                            with open(item["path"], "rb") as reader, open(source, "wb") as writer:
                                while True:
                                    chunk = reader.read(4 * 1024 * 1024)
                                    if not chunk:
                                        break
                                    writer.write(chunk)
                                    with self.lock:
                                        if job.get("cancel"):
                                            staging_cancelled = True
                                            break
                            if staging_cancelled:
                                with self.lock:
                                    item.update(state="cancelled")
                                continue
                        output = temp / "result.html"
                        metadata_path = temp / "metadata.json"
                        log_path = temp / "worker.log"
                        prefix = [sys.executable, "--worker", "--log", str(log_path)] if getattr(sys, "frozen", False) else [python_exe(), "-B", str(HERE / "convert_worker.py")]
                        args = prefix + ["--root", str(self.engine_root),
                                "--source", str(source), "--output", str(output), "--metadata", str(metadata_path), "--modes", ",".join(job["modes"])]
                        if job["toc"]:
                            args.append("--toc")
                        if job["forecast_tables"]:
                            args.append("--forecast-tables")
                        marker = secrets.token_hex(8)
                        with log_path.open("w", encoding="utf-8") as log:
                            process = subprocess.Popen(args, stdout=log, stderr=subprocess.STDOUT,
                                                       env=dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1",
                                                                WECHAT_WORKBENCH_MARKER=marker),
                                                       cwd=str(HERE), creationflags=NO_WINDOW)
                            # 预估特征在 worker 启动之后读：与 worker 启动并行，不推迟转换；日志从头读，不会漏行
                            stats = docx_stats(source) if source.suffix.lower() == ".docx" else None
                            with self.lock:
                                item["stats"] = stats
                            stop_reason = self.follow_log(job, item, log_path, process, marker)
                        finished_at = time.time()
                        # 步内进度行（[SUB]）只供 ETA，不进结果页的日志与报错
                        log_text = "".join(line for line in log_path.read_text(encoding="utf-8", errors="replace").splitlines(True)
                                           if not ETA_SUB_LINE.match(line))
                        with self.lock:
                            item["exit_code"] = process.returncode
                        if stop_reason == "cancel":
                            with self.lock:
                                item.pop("content", None)
                                item.update(state="cancelled")
                            continue
                        if stop_reason == "unkillable":
                            with self.lock:
                                item.pop("content", None)
                                job["unkillable"] = True
                                item.update(state="error", error="转换进程无法终止，为避免与残留进程冲突，工作台已停止接受新任务。请关闭工作台并重新启动。")
                            continue
                        if stop_reason == "timeout":
                            raise RuntimeError(f"转换超过 {ITEM_TIMEOUT // 60} 分钟仍未完成，已自动停止。请检查报告内容后重试。")
                        if process.returncode or not output.is_file():
                            # worker 异常结束时 Word 可能没走到 Quit：按标记收割本实例
                            threading.Thread(target=reap_word_instance, args=(marker,), daemon=True).start()
                            # 日志已逐行进了结果页，这里只留退出码与最后一行（通常是异常类型与信息），不再重复整段日志
                            last = next((line.strip() for line in reversed(log_text.splitlines()) if line.strip()), "")
                            raise RuntimeError(f"转换进程退出码 {process.returncode}。" + (f"\n{last}" if last else "转换没有生成结果。"))
                        html = output.read_text(encoding="utf-8")
                        warnings = [line for line in log_text.splitlines() if any(w in line.lower() for w in ("warn", "skipped", "fallback", "跳过"))]
                        try:
                            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
                            if not isinstance(metadata, dict):
                                raise ValueError("not an object")
                        except (OSError, ValueError):
                            # 元数据只供标题与自检提示：读不出时正文照常交付，标题留空由用户填写
                            metadata = {}
                            if not any("标题等信息未能写出" in line for line in warnings):
                                warnings.append("[warn] 未能读取标题等信息：标题请手动填写，自检与文末三表的处理情况无法显示")
                    with self.lock:
                        main_title = metadata.get("main_title", "")
                        subtitle = metadata.get("subtitle", metadata.get("title", ""))
                        title_type = job["title_type"]
                        if title_type == AUTO_TYPE:
                            title_type = metadata.get("report_kind") or "深度"
                        options = job.get("title_options") or {}
                        title = compose_title(main_title, subtitle, title_type, job["title_prefix"],
                                              include_subtitle=options.get("include_subtitle", True),
                                              subtitle_sep=options.get("subtitle_sep", " | "),
                                              include_type=options.get("include_type", True),
                                              type_sep=options.get("type_sep", "："))
                        item.update(state="done", html=html, title=title, main_title=main_title, subtitle=subtitle, title_type=title_type,
                                    self_check=metadata.get("self_check", ""), self_check_hint=metadata.get("self_check_hint", ""),
                                    forecast_stripped=bool(metadata.get("forecast_stripped")),
                                    forecast_appended=bool(metadata.get("forecast_appended")),
                                    url="result/" + job["id"] + "/" + str(job["items"].index(item)))
                        if warnings:
                            item["warnings"] = warnings
                    self.learn_eta(job, item, finished_at)
                except Exception as exc:
                    with self.lock:
                        item.pop("content", None)
                        item.update(state="error", error=str(exc), hint=failure_hint(exc, "\n".join(item.get("log", []))))
        finally:
            with self.lock:
                for item in job["items"]:
                    if "upload" in item:
                        self.uploads.pop(item["upload"], None)
                unkillable = bool(job.get("unkillable"))
            # 终止失败时不释放互斥锁：残留 worker/Word 仍在运行，放行新任务会并发两个
            # Word 转换；工作台保持“忙”状态并已提示用户重启
            if not unkillable:
                self.operation.release()
            with self.lock:
                job["state"] = "done"
                job["expires"] = time.monotonic() + 7200

    def eta_memory(self):
        with self.lock:
            # 读失败也缓存空记忆：之后的轮询不再反复读文件
            if self.eta_store is None:
                try:
                    self.eta_store = load_eta_store()
                except Exception:
                    self.eta_store = []
            return self.eta_store

    def learn_eta(self, job, item, finished_at):
        """成功转换后按各步实际耗时算本次的本机性能系数，追加到记录（打包版写回 EXE 的备用数据流）。

        附带功能：此刻报告已标记完成，任何异常都不得把它改判为失败。"""
        try:
            store = self.eta_memory()
            with self.lock:
                self.eta_store = eta_learn(store, item, finished_at, job.get("modes"))
                snapshot = list(self.eta_store)
            save_eta_store(snapshot)
        except Exception:
            pass

    def record_lines(self, item, lines):
        with self.lock:
            log = item.setdefault("log", [])
            for raw in lines:
                line = raw.rstrip()
                if not line:
                    continue
                sub = ETA_SUB_LINE.match(line)
                if sub:
                    # 步内进度：不进日志，只记到 item 供 ETA（到达时刻同样取服务端时钟）
                    eta_record_sub(item, sub, time.time())
                    continue
                log.append(line)
                step = STEP_LINE.match(line)
                if step:
                    item["progress"] = {"step": int(step.group(1)), "total": int(step.group(2)), "text": step.group(3)}
                    # 到达时刻取服务端时钟：ETA 按已完成步骤的实际耗时修正，不依赖页面所在机器的时钟
                    item.setdefault("step_times", []).append((int(step.group(1)), int(step.group(2)), time.time()))
                    del item["step_times"][:-32]
                elif line.startswith("[modes] "):
                    item["detected_modes"] = line[len("[modes] "):].split(",")
            del log[:-200]

    def follow_log(self, job, item, log_path, process, marker):
        """转换进行中把 worker 输出和 [STEP] 进度同步给结果页，而不是等进程结束再读。

        同时监听取消与超时：返回 None（正常结束）、"cancel"、"timeout"，
        或 "unkillable"（终止失败、worker 仍在运行——调用方不得释放互斥锁）。"""
        pending = ""
        stop_reason = None
        deadline = time.monotonic() + ITEM_TIMEOUT
        kill_deadline = None
        with log_path.open("r", encoding="utf-8", errors="replace") as reader:
            while True:
                finished = process.poll() is not None
                if not finished and stop_reason is None:
                    with self.lock:
                        cancelled = bool(job.get("cancel"))
                    if cancelled or time.monotonic() > deadline:
                        stop_reason = "cancel" if cancelled else "timeout"
                        kill_deadline = time.monotonic() + 15
                        kill_process_tree(process.pid)
                        # Word 由 DCOM 拉起，不在 worker 进程树内：交给后台收割线程
                        # 按窗口标题定位单杀（见 reap_word_instance），不阻塞取消响应
                        threading.Thread(target=reap_word_instance, args=(marker,), daemon=True).start()
                chunk = reader.read()
                if chunk:
                    pending += chunk
                    lines = pending.split("\n")
                    pending = lines.pop()
                    self.record_lines(item, lines)
                elif finished:
                    break
                elif kill_deadline and time.monotonic() > kill_deadline:
                    if process.poll() is None:
                        stop_reason = "unkillable"
                    break
                else:
                    time.sleep(0.2)
        if pending.strip():
            self.record_lines(item, [pending])
        return stop_reason

    def maintain(self):
        while True:
            time.sleep(2)
            now = time.monotonic()
            with self.lock:
                for key, job in list(self.jobs.items()):
                    if job["state"] == "done" and now > job["expires"]:
                        del self.jobs[key]
                for key, upload in list(self.uploads.items()):
                    if now - upload["created"] > 7200 and not self.operation.locked():
                        del self.uploads[key]
                self.clients = {key: seen for key, seen in self.clients.items() if now - seen < 600}
                # viewers 与 clients 同口径过期：release 信标丢失时服务仍能按心跳超时退出，
                # 结果本身的保留期只由 expires 决定，不受此影响
                for job in self.jobs.values():
                    job["viewers"] = {key: seen for key, seen in job["viewers"].items() if now - seen < 600}
                alive_results = any(job["viewers"] and now < job["expires"] for job in self.jobs.values())
                # 宽限 90 秒：够 F5 刷新和 EXE 冷启动开浏览器；信标丢失时由上面的
                # 600 秒客户端心跳超时兜底（最坏退回旧行为）
                should_stop = not self.operation.locked() and not self.clients and not alive_results and now - self.last_activity > 90
            if should_stop and self.operation.acquire(blocking=False):
                with self.lock:
                    now = time.monotonic()
                    still_idle = (not self.clients
                                  and not any(job["viewers"] and now < job["expires"] for job in self.jobs.values())
                                  and now - self.last_activity > 90)
                    if still_idle:
                        # 与复查同一锁块置位：迟到的心跳要么先登记（本轮放弃关闭），
                        # 要么在置位后到达并收到明确的“正在关闭”错误
                        self.closing = True
                if still_idle:
                    self.server.shutdown()
                    return
                self.operation.release()


class Handler(BaseHTTPRequestHandler):
    server_version = "WechatWorkbench"

    @property
    def app(self):
        return self.server.app

    def log_message(self, *_):
        pass

    def send(self, code, content, mime="application/json; charset=utf-8"):
        if isinstance(content, (dict, list)):
            content = json.dumps(content, ensure_ascii=False)
        if isinstance(content, str):
            content = content.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Content-Security-Policy", "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; frame-src 'self'; frame-ancestors 'self'; object-src 'none'; base-uri 'self'; connect-src 'self'")
        self.end_headers()
        try:
            self.wfile.write(content)
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            pass

    def route(self):
        expected = "127.0.0.1:" + str(self.server.server_port)
        if self.headers.get("Host") != expected:
            raise PermissionError("无效主机。")
        origin = self.headers.get("Origin")
        if origin and origin != "http://" + expected:
            raise PermissionError("不允许跨站请求。")
        prefix = "/" + self.app.token + "/"
        path = urlsplit(self.path).path
        if not path.startswith(prefix):
            raise PermissionError("访问链接无效，请重新打开工作台。")
        return path[len(prefix):]

    def body(self, limit=65536):
        size = int(self.headers.get("Content-Length", "0"))
        if not 0 <= size <= limit:
            raise ValueError("请求或文件过大。")
        self.connection.settimeout(60)
        data = self.rfile.read(size)
        if len(data) != size:
            raise ValueError("文件未完整传输，请重新添加。")
        return data

    def do_GET(self):
        try:
            route = self.route()
            if route == "":
                page = (HERE / "web" / "index.html").read_text(encoding="utf-8")
                settings = load_settings()   # 一次快照：避免两键分读到并发保存的混合值
                options = load_title_options(settings)
                escape = lambda s: s.replace("&", "&amp;").replace("<", "&lt;").replace('"', "&quot;")
                page = page.replace("{{APP_VERSION}}", APP_VERSION)
                page = page.replace("{{FORECAST_TABLES}}", "checked" if load_forecast_tables(settings) else "")
                page = (page.replace("{{TITLE_INCLUDE_SUBTITLE}}", "checked" if options["include_subtitle"] else "")
                            .replace("{{TITLE_SUBTITLE_SEP}}", escape(options["subtitle_sep"]))
                            .replace("{{TITLE_INCLUDE_TYPE}}", "checked" if options["include_type"] else "")
                            .replace("{{TITLE_TYPE_SEP}}", escape(options["type_sep"])))
                # 用户前缀最后注入：其值若含花括号占位符文本也不会被后续替换二次改写
                page = page.replace("{{TITLE_PREFIX}}", escape(load_title_prefix(settings)))
                self.send(200, page.encode("utf-8"), "text/html; charset=utf-8")
            elif route == "app-icon.ico":
                self.send(200, (HERE / "assets" / "app-icon.ico").read_bytes(), "image/x-icon")
            elif route == "api/update":
                # 只回后台线程已经查到的结果，页面请求不触发联网
                self.send(200, self.app.update or {})
            elif route == "api/status":
                dependencies = self.app.check_dependencies()
                with self.app.lock:
                    last = self.app.jobs.get(self.app.last_job)
                    payload = {"dependencies": dependencies, "busy": self.app.operation.locked(), "install": dict(self.app.install),
                               "last_job": {"id": last["id"], "state": last["state"]} if last else None}
                self.send(200, payload)
            elif route.startswith("api/jobs/"):
                try:
                    store = self.app.eta_memory()
                except Exception:
                    store = []
                with self.app.lock:
                    job = self.app.jobs.get(route.split("/")[-1])
                    now = time.time()
                    payload = None if not job else {"id": job["id"], "state": job["state"], "now": now, "modes": list(job["modes"]), "items": [
                        {k: (list(v) if k == "log" else dict(v) if k == "progress" else v)
                         for k, v in item.items() if k not in ("html", "path", "upload", "content", "step_times")} for item in job["items"]]}
                    if payload:
                        for row, item in zip(payload["items"], job["items"]):
                            if item.get("state") == "running":
                                row["eta"], row["eta_progress"], row["eta_high"] = safe_item_eta(item, now, store, job["modes"])
                if payload is None:
                    self.send(404, {"error": "结果已清理，请重新转换。"})
                else:
                    self.send(200, payload)
            elif route.startswith("job/"):
                job_id = route.split("/")[-1]
                with self.app.lock:
                    if job_id not in self.app.jobs:
                        self.send(404, "结果已清理，请返回工作台重新转换。", "text/plain; charset=utf-8")
                        return
                page = (HERE / "web" / "result.html").read_text(encoding="utf-8")
                page = page.replace("__BASE__", "/" + self.app.token + "/").replace("__JOB__", job_id).replace("{{APP_VERSION}}", APP_VERSION)
                self.send(200, page, "text/html; charset=utf-8")
            elif route.startswith("report/"):
                _, job_id, index = route.split("/")
                with self.app.lock:
                    job = self.app.jobs.get(job_id)
                    if not job or not 0 <= int(index) < len(job["items"]):
                        self.send(404, "结果已清理。", "text/plain; charset=utf-8")
                        return
                    item = job["items"][int(index)]
                    if item.get("state") != "error":
                        self.send(404, {"error": "该报告没有错误信息。"})
                        return
                    job, item = dict(job), dict(item, log=list(item.get("log", [])))
                # 环境段要查进程列表，在锁外生成，不阻塞轮询与转换
                report = build_error_report(job, item)
                self.send(200, report, "text/plain; charset=utf-8")
            elif route.startswith("result/"):
                _, job_id, index = route.split("/")
                with self.app.lock:
                    job = self.app.jobs.get(job_id)
                    if not job or not 0 <= int(index) < len(job["items"]):
                        self.send(404, "结果已清理。", "text/plain; charset=utf-8")
                        return
                    html = job["items"][int(index)].get("html")
                if html is None:
                    self.send(404, {"error": "结果尚未生成。"})
                else:
                    self.send(200, html, "text/html; charset=utf-8")
            else:
                self.send(404, {"error": "页面不存在。"})
        except PermissionError as exc:
            self.send(403, {"error": str(exc)})
        except (ValueError, KeyError, IndexError) as exc:
            self.send(400, {"error": str(exc)})
        except Exception as exc:
            self.send(500, {"error": str(exc)})

    def do_POST(self):
        try:
            route = self.route()
            if route == "api/upload":
                name = unquote(self.headers.get("X-Filename", ""))
                if not name or Path(name).name != name or "/" in name or "\\" in name or Path(name).suffix.lower() not in (".doc", ".docx"):
                    raise ValueError("仅支持 .docx / .doc 文件。")
                content = self.body(MAX_FILE)
                if not content:
                    raise ValueError("文件为空。")
                with self.app.lock:
                    if sum(len(u["content"]) for u in self.app.uploads.values()) + len(content) > 500 * 1024 * 1024:
                        raise ValueError("暂存文件超过 500 MB，请退出工作台后重新启动。")
                    upload_id = secrets.token_hex(12)
                    self.app.uploads[upload_id] = {"name": name, "content": content, "created": time.monotonic()}
                    # 上传视为活跃：防止心跳恰好过期时空闲退出打断进行中的批次
                    self.app.last_activity = time.monotonic()
                self.send(200, {"id": upload_id, "name": name})
                return
            data = json.loads(self.body() or b"{}")
            if not isinstance(data, dict):
                raise ValueError("请求格式不正确。")
            if route == "api/jobs":
                job_id = self.app.create_job(data)
                self.send(200, {"id": job_id})
            elif route == "api/install":
                if not self.app.operation.acquire(blocking=False):
                    raise ValueError("当前正在转换或安装，请稍后重试。")
                with self.app.lock:
                    self.app.install = {"state": "running", "phase": "preparing", "source": "", "log": ""}
                try:
                    threading.Thread(target=self.app.install_missing, daemon=True).start()
                except Exception:
                    # 线程都起不来时回滚状态并释放锁，否则工作台永久报忙
                    with self.app.lock:
                        self.app.install.update(state="error", phase="error")
                    self.app.operation.release()
                    raise
                self.send(200, {})
            elif route == "api/discard":
                ids = data.get("uploads")
                # 任务在创建时已自持上传内容的引用，这里删除暂存不影响任何已建任务
                if isinstance(ids, list):
                    with self.app.lock:
                        for upload_id in ids[:60]:
                            if isinstance(upload_id, str):
                                self.app.uploads.pop(upload_id, None)
                self.send(200, {})
            elif route == "api/cancel":
                # 只置标志：实际终止由 follow_log 的轮询在 0.2 秒内执行，
                # 队列中尚未开始的报告在下一轮循环标记为已取消
                with self.app.lock:
                    job = self.app.jobs.get(data.get("job"))
                    if job and job["state"] != "done":
                        job["cancel"] = True
                self.send(200, {})
            elif route == "api/client-release":
                # 页面关闭时的注销信标（尽力而为）：让空闲退出等待从心跳超时缩短到宽限期
                with self.app.lock:
                    self.app.clients.pop(str(data.get("client", "main"))[:100], None)
                self.send(200, {})
            elif route == "api/heartbeat":
                with self.app.lock:
                    if self.app.closing:
                        raise ValueError("工作台正在关闭。")
                    self.app.clients[str(data.get("client", "main"))[:100]] = time.monotonic()
                    self.app.last_activity = time.monotonic()
                self.send(200, {})
            elif route in ("api/result-heartbeat", "api/release"):
                with self.app.lock:
                    if self.app.closing and route.endswith("result-heartbeat"):
                        raise ValueError("工作台正在关闭。")
                    job = self.app.jobs.get(data.get("job"))
                    if job:
                        viewer = str(data.get("viewer", "default"))[:100]
                        if route.endswith("release"):
                            # 离开结果页只注销查看者，不缩短保留期：结果按 expires（2 小时）
                            # 或服务退出清理，工作台的“重新打开上次结果”入口仍有效
                            job["viewers"].pop(viewer, None)
                        else:
                            job["viewers"][viewer] = time.monotonic()
                            job["expires"] = time.monotonic() + 7200
                            self.app.last_activity = time.monotonic()
                self.send(200, {})
            elif route == "api/shutdown":
                # 原子取得操作锁再关闭，避免检查与关闭之间有新任务启动（不再释放：服务即将退出）
                if not self.app.operation.acquire(blocking=False):
                    raise ValueError("转换或安装尚未结束，请等待完成后退出。")
                with self.app.lock:
                    self.app.closing = True
                # 先调度关闭再发响应：响应写出失败也不影响关闭，锁不会悬死
                try:
                    threading.Thread(target=self.server.shutdown, daemon=True).start()
                except Exception:
                    with self.app.lock:
                        self.app.closing = False
                    self.app.operation.release()
                    raise
                self.send(200, {})
            else:
                self.send(404, {"error": "接口不存在。"})
        except PermissionError as exc:
            self.send(403, {"error": str(exc)})
        except (ValueError, KeyError, TypeError, OSError) as exc:
            self.send(400, {"error": str(exc)})
        except Exception as exc:
            self.send(500, {"error": str(exc)})


def dropped_sources(paths):
    """拖到 EXE 上的文件参数：只收现存、不超限的 .doc/.docx，返回（有效路径, 提示信息）。

    超限文件在这里就剔除：留到 create_job 会因单份校验失败而使整批不成立，
    连同批的合法报告一起转不了。"""
    valid, rejected = [], []
    for raw in paths:
        path = Path(str(raw).strip().strip('"'))
        try:
            oversize = path.is_file() and path.stat().st_size > MAX_FILE
        except OSError:
            oversize = False
        if path.is_file() and path.suffix.lower() in (".doc", ".docx") and not path.name.startswith("~$") and not oversize:
            valid.append(path)
        else:
            rejected.append(path.name or str(path))
    notice = ""
    if rejected:
        notice = "以下文件未加入转换（仅支持现存的 .doc/.docx，单份不超过 100 MB）：\n" + "\n".join(rejected[:10])
    if len(valid) > 30:
        notice = (notice + "\n\n" if notice else "") + "一次最多转换 30 份，超出部分已忽略。"
        valid = valid[:30]
    return valid, notice


def main():
    exe = Path(sys.argv[0]).name
    parser = maintenance.ConsoleArgumentParser(
        description=f"Word-to-WeChat report workbench {APP_VERSION}. Converts Soochow-template Word reports "
                    "into HTML for the WeChat Official Account editor. Without arguments it starts the local "
                    "workbench and opens the browser.",
        epilog=f"Maintenance: {exe} --export-maintenance <empty dir> writes this version's source code and "
               "maintenance docs; start with MAINTAINING.md.")
    parser.add_argument("-V", "--version", action="version", version=f"WordToWeChat {APP_VERSION}")
    parser.add_argument("--export-maintenance", type=Path, metavar="DIR",
                        help="export source code, MAINTAINING.md and the WeChat layout checklist to DIR, then exit")
    parser.add_argument("--engine-root", type=Path, help="directory holding configs/ (default: built-in)")
    parser.add_argument("--no-browser", action="store_true", help="start the server without opening a browser")
    parser.add_argument("--port", type=int, default=0, help="local port (default: random)")
    parser.add_argument("--ready-file", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--no-update-check", action="store_true",
                        help="skip the one-time GitHub release check at startup")
    parser.add_argument("sources", nargs="*", type=Path,
                        help="Word reports to convert right away in auto mode (same as dropping them on the EXE)")
    args = parser.parse_args()
    if args.export_maintenance:
        written = maintenance.export_maintenance(args.export_maintenance)
        print(f"Exported {len(written)} files to {args.export_maintenance.resolve()}; start with MAINTAINING.md.")
        return
    root = args.engine_root or HERE
    remove_legacy_eta_file()
    clean_orphans()
    app = Workbench(root.resolve())
    server = ThreadingHTTPServer(("127.0.0.1", args.port), Handler)
    server.daemon_threads = True
    server.app = app
    app.server = server
    url = "http://127.0.0.1:" + str(server.server_port) + "/" + app.token + "/"
    if sys.stdout:
        print(url, flush=True)
    if args.ready_file:
        args.ready_file.write_text(url, encoding="utf-8")
    open_url = url
    if args.sources:
        valid, notice = dropped_sources(args.sources)
        if notice and os.name == "nt":
            ctypes.windll.user32.MessageBoxW(None, notice, "报告转公众号", 0x30)
        if valid:
            settings = load_settings()
            options = load_title_options(settings)
            try:
                job_id = app.create_job({
                    "inputs": [{"path": str(p)} for p in valid],
                    "modes": ["auto"], "toc": True, "report_type": AUTO_TYPE, "title_type": AUTO_TYPE,
                    "title_prefix": load_title_prefix(settings),
                    "forecast_tables": load_forecast_tables(settings),
                    "title_include_subtitle": options["include_subtitle"],
                    "title_subtitle_sep": options["subtitle_sep"],
                    "title_include_type": options["include_type"],
                    "title_type_sep": options["type_sep"]})
                open_url = url + "job/" + job_id
            except Exception as exc:
                if os.name == "nt":
                    ctypes.windll.user32.MessageBoxW(None, f"未能自动开始转换：{exc}", "报告转公众号", 0x30)
    if not args.no_update_check:
        def check_update():
            app.update = fetch_latest_release()

        threading.Thread(target=check_update, daemon=True).start()
    # 拖入提示弹窗可能阻塞很久：maintain 须在其后启动并从此刻重新计空闲，否则 90 秒宽限先关掉服务
    app.last_activity = time.monotonic()
    threading.Thread(target=app.maintain, daemon=True).start()
    if not args.no_browser:
        webbrowser.open_new(open_url)
    try:
        server.serve_forever(poll_interval=0.5)
    finally:
        server.server_close()
        app.jobs.clear()
        app.uploads.clear()
        if args.ready_file:
            args.ready_file.unlink(missing_ok=True)


if __name__ == "__main__":
    if "--worker" in sys.argv:
        sys.argv.remove("--worker")
        if "--log" in sys.argv:
            index = sys.argv.index("--log")
            log_path = sys.argv[index + 1]
            del sys.argv[index:index + 2]
            sys.stdout = sys.stderr = open(log_path, "w", encoding="utf-8", buffering=1)
        import convert_worker
        sys.exit(convert_worker.main())
    if maintenance.wants_console(sys.argv):
        maintenance.attach_console()
        sys.argv = [("--help" if arg == "/?" else arg) for arg in sys.argv]
    try:
        main()
    except Exception:
        message = traceback.format_exc()
        if sys.stderr:
            sys.stderr.write(message)
        if os.name == "nt":
            ctypes.windll.user32.MessageBoxW(None, message, "报告转公众号：启动失败", 0x10)
        sys.exit(1)
