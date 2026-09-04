from __future__ import annotations

import argparse
import ctypes
import importlib
import json
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


HERE = Path(__file__).resolve().parent
APP_VERSION = "v1.1.0"
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


# 路径内允许空格（如 Program Files）和中文段（如 工具软件），空格后紧跟中文视为
# 正文开始；宁可多脱敏掉路径后紧跟的英文词，也不能泄露含空格路径的后半段
_PATH_CHARS = r"(?:[^\s<>:\"|?*\r\n'，。;；]|[ ](?![一-鿿]))"
PATH_RE = re.compile(r"(?:\\\\\?\\)?[A-Za-z]:[\\/]" + _PATH_CHARS + r"*|\\\\" + _PATH_CHARS + "+")


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
        "",
        "## 进度与日志（脱敏后）",
    ]
    body = clean("\n".join(item.get("log", [])))
    error = ["", "## 错误信息（脱敏后）"] + clean(item.get("error", "")) if item.get("error") else []
    return "\n".join(head + body + error) + "\n"


MAX_ZIP_ENTRIES = 5000


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

    设有 zip 守卫（条目数、document.xml 解压尺寸）防异常压缩比拖垮内存；
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
        ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        return {"size": path.stat().st_size, "media_count": len(media), "media_bytes": media_bytes,
                "paragraphs": len(list(xml.iter(ns + "p"))), "tables": len(list(xml.iter(ns + "tbl"))),
                "chars": sum(len(t.text or "") for t in xml.iter(ns + "t"))}
    except Exception:
        return None


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
                    item.update(state="running", started=time.time(), log=[], exit_code=None,
                                progress={"step": 0, "total": 0, "text": "准备文件"})
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
                        stats = docx_stats(source) if source.suffix.lower() == ".docx" else None
                        with self.lock:
                            item["stats"] = stats
                        with log_path.open("w", encoding="utf-8") as log:
                            process = subprocess.Popen(args, stdout=log, stderr=subprocess.STDOUT,
                                                       env=dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1",
                                                                WECHAT_WORKBENCH_MARKER=marker),
                                                       cwd=str(HERE), creationflags=NO_WINDOW)
                            stop_reason = self.follow_log(job, item, log_path, process, marker)
                        log_text = log_path.read_text(encoding="utf-8", errors="replace")
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
                            raise RuntimeError(f"转换进程退出码 {process.returncode}。\n" + (log_text[-12000:] or "转换没有生成结果。"))
                        html = output.read_text(encoding="utf-8")
                        metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
                        warnings = [line for line in log_text.splitlines() if any(w in line.lower() for w in ("warn", "skipped", "fallback", "跳过"))]
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
                except Exception as exc:
                    with self.lock:
                        item.pop("content", None)
                        item.update(state="error", error=str(exc))
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

    def record_lines(self, item, lines):
        with self.lock:
            log = item.setdefault("log", [])
            for raw in lines:
                line = raw.rstrip()
                if not line:
                    continue
                log.append(line)
                step = STEP_LINE.match(line)
                if step:
                    item["progress"] = {"step": int(step.group(1)), "total": int(step.group(2)), "text": step.group(3)}
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
                with self.app.lock:
                    job = self.app.jobs.get(route.split("/")[-1])
                    payload = None if not job else {"id": job["id"], "state": job["state"], "now": time.time(), "modes": list(job["modes"]), "items": [
                        {k: (list(v) if k == "log" else dict(v) if k == "progress" else v)
                         for k, v in item.items() if k not in ("html", "path", "upload", "content")} for item in job["items"]]}
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
    parser = argparse.ArgumentParser()
    parser.add_argument("--engine-root", type=Path)
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--port", type=int, default=0)
    parser.add_argument("--ready-file", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--no-update-check", action="store_true",
                        help="不查 GitHub 上有没有新版本（默认启动后台查一次）")
    parser.add_argument("sources", nargs="*", type=Path,
                        help="拖到程序图标上的 Word 报告，启动后按“自动”模式直接转换")
    args = parser.parse_args()
    root = args.engine_root or HERE
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
    # 拖入提示弹窗可能阻塞很久：maintain 空闲计时须在其后才启动，否则 90 秒宽限先关掉服务
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
    try:
        main()
    except Exception:
        message = traceback.format_exc()
        if sys.stderr:
            sys.stderr.write(message)
        if os.name == "nt":
            ctypes.windll.user32.MessageBoxW(None, message, "报告转公众号：启动失败", 0x10)
        sys.exit(1)
