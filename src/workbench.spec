from pathlib import Path

root = Path(SPECPATH)
a = Analysis(
    [str(root / "app.py")],
    pathex=[str(root)],
    binaries=[],
    datas=[(str(root / "web"), "web"), (str(root / "configs"), "configs"), (str(root / "assets"), "assets")],
    hiddenimports=["convert_worker", "converter", "table_image", "toc_image", "valuation_table_image", "win32com.client", "pythoncom", "pywintypes", "docx", "PIL.Image", "bs4"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "numpy", "pandas", "scipy", "IPython", "pytest", "tkinter", "PIL.ImageTk"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas,
    [], name="报告转公众号", debug=False, bootloader_ignore_signals=False,
    strip=False, upx=False, console=False, icon=str(root / "assets" / "app-icon.ico"),
    disable_windowed_traceback=False,
)
