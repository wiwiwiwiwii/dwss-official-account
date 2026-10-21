"""命令行维护入口：让窗口版 EXE 的 --help/--version 与参数错误输出到终端，并导出随包的源码与维护文档。

维护包与工作区同布局（src/、docs/、tests/、MAINTAINING.md、LICENSE），打包时由
workbench.spec 收进 EXE 的 maintenance/ 目录；源码运行时直接取工作区文件。
"""
from __future__ import annotations

import argparse
import ctypes
import os
import shutil
import sys
from pathlib import Path

CONSOLE_FLAGS = {"-h", "--help", "/?", "-V", "--version", "--export-maintenance"}
WORKSPACE = Path(__file__).resolve().parent.parent
# 源码运行时的维护包清单（相对工作区）；打包版收录同一份清单，见 workbench.spec
BUNDLE_FILES = ("MAINTAINING.md", "LICENSE", "docs/wechat-mp-official/checklist.md",
                "docs/wechat-mp-official/verify_article_structure.md",
                "docs/wechat-mp-official/LICENSE-verify-article-structure-spec.txt",
                "tests/wechat_compliance_check.py")
SOURCE_SKIP = {"__pycache__"}
_console_attached = False


def wants_console(argv: list[str]) -> bool:
    return any(arg in CONSOLE_FLAGS or arg.startswith("--export-maintenance=") for arg in argv[1:])


def attach_console() -> None:
    """窗口版 EXE 没有标准输出：父进程传入的是管道或文件就直接写；是控制台句柄
    （窗口程序继承了句柄值却没挂上控制台，直接写会 Errno 22）或没有句柄时，挂到父进程控制台。"""
    global _console_attached
    if _console_attached or os.name != "nt" or not getattr(sys, "frozen", False):
        return
    _console_attached = True
    kernel32 = ctypes.windll.kernel32
    kernel32.GetStdHandle.restype = ctypes.c_void_p
    kernel32.GetFileType.argtypes = [ctypes.c_void_p]
    streams = {}
    for name, handle_id in (("stdout", -11), ("stderr", -12)):
        handle = kernel32.GetStdHandle(handle_id)
        # 1 = FILE_TYPE_DISK，3 = FILE_TYPE_PIPE；2 = 控制台字符设备，0 = 无效
        if handle not in (None, 0, ctypes.c_void_p(-1).value) and kernel32.GetFileType(handle) in (1, 3):
            try:
                import msvcrt
                fd = msvcrt.open_osfhandle(handle, os.O_WRONLY)
                streams[name] = open(fd, "w", encoding="utf-8", buffering=1, closefd=False)
            except OSError:
                pass
    if len(streams) < 2 and kernel32.AttachConsole(-1):
        for name in ("stdout", "stderr"):
            streams.setdefault(name, open("CONOUT$", "w", encoding="utf-8", buffering=1))
    for name, stream in streams.items():
        setattr(sys, name, stream)


class ConsoleArgumentParser(argparse.ArgumentParser):
    """参数错误也挂到终端：未知参数不在 CONSOLE_FLAGS 里，否则报错按系统代码页写进管道，中文程序名乱码。"""

    def error(self, message):
        attach_console()
        super().error(message)


def bundle_root() -> Path | None:
    """打包版返回 EXE 内的 maintenance/ 目录；源码运行返回 None（直接取工作区）。"""
    base = getattr(sys, "_MEIPASS", None)
    return Path(base) / "maintenance" if base else None


def export_maintenance(target: Path) -> list[str]:
    """把维护包写到 target（须为不存在或空的目录），返回写出的相对路径。"""
    target = target.resolve()
    if target.exists() and (not target.is_dir() or any(target.iterdir())):
        raise SystemExit(f"Target directory exists and is not empty: {target}")
    root = bundle_root()
    pairs: list[tuple[Path, str]] = []
    if root:
        for path in sorted(root.rglob("*")):
            if path.is_file():
                pairs.append((path, path.relative_to(root).as_posix()))
    else:
        src = WORKSPACE / "src"
        for path in sorted(src.rglob("*")):
            if path.is_file() and not SOURCE_SKIP.intersection(path.relative_to(src).parts):
                pairs.append((path, "src/" + path.relative_to(src).as_posix()))
        pairs += [(WORKSPACE / rel, rel) for rel in BUNDLE_FILES if (WORKSPACE / rel).is_file()]
    for source, rel in pairs:
        dest = target / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, dest)
    return [rel for _, rel in pairs]
