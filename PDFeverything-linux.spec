# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec — Linux 单文件版
=================================
用法: pyinstaller PDFeverything-linux.spec --noconfirm --clean
      （或直接 python build_linux.py）
输出: dist/PDFeverything          单文件可执行
      dist/PDFeverything.desktop  桌面启动器（配合图标使用）

与 build_windows.spec 的差异：没有 win32com/pywin32 隐藏导入（Linux 上
Office 转换自动走 python-docx 等纯 Python 回退）；图标用 PNG（X11/Wayland
窗口管理器不认 .ico/.icns）。

同样排除 PyMuPDF 2.x 的 AI 分析栈（pymupdf4llm/pymupdf.layout/onnxruntime），
否则包体凭空膨胀几百 MB——core/fast_pymupdf 在运行时也拦一道，双保险。
"""

from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files

PROJECT_ROOT = Path(SPECPATH)

datas = []
for mod in ('fitz', 'PIL'):
    try:
        datas += collect_data_files(mod)
    except Exception:
        pass

# PNG for the window icon (X11/Wayland) and the .desktop launcher template
datas.append((str(PROJECT_ROOT / 'resources' / 'app_icon.png'), 'resources'))
datas.append((str(PROJECT_ROOT / 'resources' / 'PDFeverything.desktop'),
              'resources'))

hiddenimports = [
    "PyQt6.QtCore",
    "PyQt6.QtGui",
    "PyQt6.QtWidgets",
    "PyQt6.sip",
    "pypdf",
    "pypdf.generic",
    "pypdf.filters",
    "pikepdf",
    "pikepdf._core",
    "pikepdf.models",
    "cryptography",
    "cryptography.hazmat.primitives.ciphers",
    "cryptography.hazmat.backends.openssl",
    "fitz",
    "PIL",
    "PIL.Image",
    "PIL.ImageDraw",
    "PIL.PdfImagePlugin",
    "docx",
    "docx.document",
    "docx.table",
    "docx.text",
    "pptx",
    "pptx.slide",
    "pptx.shapes",
    "openpyxl",
    "openpyxl.worksheet",
    "openpyxl.reader",
    "mcp", "mcp.server",
    "core", "core.fast_pymupdf", "core.page_editor", "core.pdf_ops", "core.utils",
    "core.converters", "core.merger", "core.search",
    "gui", "gui.i18n", "gui.pdf_reader_widget", "gui.main_window",
    "gui.file_list_widget", "gui.workers", "gui.dialogs",
    "tempfile",
    "subprocess",
    "shutil",
    "uuid",
]

excludes = [
    "tkinter",
    "matplotlib",
    "numpy",
    "scipy",
    "pandas",
    "jedi",
    "IPython",
    "notebook",
    "sphinx",
    "pytest",
    "setuptools",
    "pip",
    "win32com", "win32api", "win32con", "pythoncom", "pywintypes",
    # PyMuPDF 2.x 的可选 AI 分析栈：本项目不用 to_markdown/layout，
    # 打进包里会连带 onnxruntime/numpy 让包体凭空膨胀几百 MB。
    "pymupdf4llm", "pymupdf.layout", "pymupdf_layout", "onnxruntime",
]

a = Analysis(
    [str(PROJECT_ROOT / "main.py")],
    pathex=[str(PROJECT_ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,         # onefile：内嵌所有 so/数据
    a.datas,
    [],
    name="PDFeverything",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True,   # 同 macOS spec：argv 透传给 CLI 模式
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
