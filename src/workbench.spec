import os
import re
import sys
from pathlib import Path

from PyInstaller.utils.win32.versioninfo import (FixedFileInfo, StringFileInfo, StringStruct, StringTable,
                                                 VarFileInfo, VarStruct, VSVersionInfo)

root = Path(SPECPATH)
workspace = root.parent
sys.path.insert(0, str(root))
from maintenance import BUNDLE_FILES, SOURCE_SKIP  # noqa: E402

app_version = re.search(r'APP_VERSION = "([^"]+)"', (root / "app.py").read_text(encoding="utf-8")).group(1)
numbers = tuple(int(n) for n in re.match(r"v(\d+)\.(\d+)\.(\d+)", app_version).groups()) + (0,)
exe_name = os.environ.get("WORKBENCH_EXE_NAME", "报告转公众号")

# 维护包：与工作区同布局收进 maintenance/，供 --export-maintenance 导出
maintenance_datas = [(str(path), str(Path("maintenance/src") / path.relative_to(root).parent))
                     for path in sorted(root.rglob("*"))
                     if path.is_file() and not SOURCE_SKIP.intersection(path.relative_to(root).parts)]
maintenance_datas += [(str(workspace / rel), str(Path("maintenance") / Path(rel).parent)) for rel in BUNDLE_FILES]

version_info = VSVersionInfo(
    ffi=FixedFileInfo(filevers=numbers, prodvers=numbers),
    kids=[StringFileInfo([StringTable("080404b0", [
        StringStruct("ProductName", "报告转公众号"),
        StringStruct("FileDescription", "报告转公众号：Word 研报转公众号文章"),
        StringStruct("ProductVersion", app_version),
        StringStruct("FileVersion", app_version),
        StringStruct("OriginalFilename", exe_name + ".exe"),
        StringStruct("LegalCopyright", "Copyright (c) 2026 wiwiwiwiwii · MIT License"),
        StringStruct("Comments", "命令行：--help 查看用法；--export-maintenance <空目录> 导出源码与维护文档"),
    ])]), VarFileInfo([VarStruct("Translation", [0x0804, 1200])])],
)

a = Analysis(
    [str(root / "app.py")],
    pathex=[str(root)],
    binaries=[],
    datas=[(str(root / "web"), "web"), (str(root / "configs"), "configs"), (str(root / "assets"), "assets")]
          + maintenance_datas,
    hiddenimports=["convert_worker", "converter", "chart_relink", "table_image", "toc_image", "valuation_table_image", "maintenance", "win32com.client", "pythoncom", "pywintypes", "docx", "PIL.Image", "bs4"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "numpy", "pandas", "scipy", "IPython", "pytest", "tkinter", "PIL.ImageTk", "win32ui"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas,
    [], name=exe_name, debug=False, bootloader_ignore_signals=False,
    strip=False, upx=False, console=False, icon=str(root / "assets" / "app-icon.ico"),
    disable_windowed_traceback=False, version=version_info,
)
