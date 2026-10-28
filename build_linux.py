#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDFeverything — Linux 构建脚本
==============================
用法:
    python build_linux.py                # 装依赖 → 打包单文件二进制
    python build_linux.py --skip-deps    # 依赖已就绪时跳过 pip（CI 用）
    python build_linux.py --no-appimage  # 只出单文件二进制，不打 AppImage

产物:
    dist/PDFeverything                            单文件可执行（免安装）
    dist/PDFeverything_Setup_v<ver>_linux.AppImage  AppImage（装有 appimagetool 时）

Linux 上 Office 文件转换自动走 python-docx / python-pptx / openpyxl 的纯
Python 回退（不需要安装 Microsoft Office）；系统装有 LibreOffice 时也可以
用 libreoffice 命令行转换，二者都不影响构建。

AppImage 打包需要 appimagetool（自动下载到构建缓存），未联网时跳过该步。
"""

import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except Exception:
        pass

PROJECT_DIR = Path(__file__).parent.resolve()
DIST_DIR = PROJECT_DIR / "dist"
BUILD_DIR = PROJECT_DIR / "build"
OUTPUT_BIN = DIST_DIR / "PDFeverything"
SPEC_FILE = PROJECT_DIR / "PDFeverything-linux.spec"
DESKTOP_FILE = PROJECT_DIR / "resources" / "PDFeverything.desktop"
ICON_FILE = PROJECT_DIR / "resources" / "app_icon.png"
REQUIREMENTS = [
    "PyQt6", "PyMuPDF", "pypdf", "pikepdf", "pillow",
    "python-docx", "python-pptx", "openpyxl", "pyinstaller",
]

SKIP_DEPS = "--skip-deps" in sys.argv
NO_APPIMAGE = "--no-appimage" in sys.argv


def print_banner():
    print()
    print("=" * 60)
    print("    PDFeverything - Linux Build Script")
    print("=" * 60)
    print()


def step(msg):
    print(f"\n>>> {msg}")


def ok(msg=""):
    print(f"    [OK] {msg}")


def fail(msg):
    print(f"\n{'='*60}")
    print(f"  [FAILED] {msg}")
    print(f"{'='*60}")
    sys.exit(1)


def run(cmd, timeout=600, cwd=None):
    if cwd is None:
        cwd = str(PROJECT_DIR)
    print(f"    $ {' '.join(str(c) for c in cmd[:5])}{'...' if len(cmd) > 5 else ''}")
    try:
        proc = subprocess.Popen(
            [str(c) for c in cmd], stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", cwd=cwd,
        )
        last = time.time()
        for line in proc.stdout:
            line = line.rstrip()
            now = time.time()
            important = "error" in line.lower() or "warning" in line.lower()
            if important or now - last > 2:
                print(f"      {line if important else line[:120]}")
                last = now
        proc.wait(timeout=timeout)
        return proc.returncode
    except subprocess.TimeoutExpired:
        proc.kill()
        print(f"    Timeout ({timeout}s)")
        return -1
    except Exception as e:
        print(f"    Exception: {e}")
        return -1


def read_version():
    try:
        text = (PROJECT_DIR / "main.py").read_text(encoding="utf-8")
        import re
        m = re.search(r'^VERSION\s*=\s*"([^"]+)"', text, re.M)
        if m:
            return m.group(1)
    except Exception as e:
        print(f"    Warning: 读取版本号失败 ({e})")
    return "0.0.0"


def step1():
    step("Step 1/4 - Check environment")
    if sys.platform not in ("linux", "linux2"):
        fail("This script is for Linux. Windows: build_windows.py, macOS: PDFeverything.spec")
    v = sys.version_info
    print(f"    Python {v.major}.{v.minor}.{v.micro}")
    if v < (3, 10):
        fail("Python 3.10+ required")
    if shutil.which("pyinstaller") is None and run([sys.executable, "-c", "import PyInstaller"]) != 0:
        fail("PyInstaller missing — rerun without --skip-deps")
    ok()


def step2():
    step("Step 2/4 - Install dependencies")
    if SKIP_DEPS:
        print("    Skipped (--skip-deps)")
        return
    for pkg in REQUIREMENTS:
        print(f"    Installing {pkg}...")
        rc = run([sys.executable, "-m", "pip", "install", "--upgrade", pkg, "--quiet"])
        if rc != 0:
            print(f"    Retrying {pkg} with verbose output...")
            if run([sys.executable, "-m", "pip", "install", pkg]) != 0:
                fail(f"Failed to install {pkg}. Check network.")
    ok("All dependencies installed")


def step3():
    step("Step 3/4 - PyInstaller onefile build (3-10 min)")
    for d in (BUILD_DIR, OUTPUT_BIN):
        try:
            if d.is_dir():
                shutil.rmtree(d)
                print(f"    Deleted: {d}")
            elif d.is_file():
                d.unlink()
                print(f"    Deleted: {d}")
        except Exception as e:
            print(f"    Warning: {e}")
    rc = run([sys.executable, "-m", "PyInstaller",
              str(SPEC_FILE), "--noconfirm", "--clean", "--log-level", "WARN"],
             timeout=1200)
    if rc != 0 or not OUTPUT_BIN.exists():
        fail("PyInstaller failed. Check disk space and the log above.")
    size = OUTPUT_BIN.stat().st_size / 1024 / 1024
    print(f"    File size: {size:.1f} MB")
    # GUI smoke test: a missing system lib (e.g. libxcb-cursor.so.0) makes the
    # Qt xcb plugin fail at startup — PyInstaller only warns, the build "succeeds",
    # and the app dies the moment a user double-clicks it. Catch that here.
    gui_env = dict(os.environ, QT_QPA_PLATFORM="offscreen")
    try:
        probe = subprocess.run(
            [str(OUTPUT_BIN), "--version"],
            capture_output=True, text=True, timeout=120, env=gui_env)
        if probe.returncode != 0:
            fail(f"Smoke test failed (rc={probe.returncode}):\n{probe.stderr[:500]}")
        print(f"    Smoke test: {probe.stdout.strip() or 'ok'}")
    except subprocess.TimeoutExpired:
        fail("Smoke test timed out")
    ok("Build complete")


def _ensure_appimagetool(cache_dir: Path) -> Path | None:
    """Download the official appimagetool x86_64 AppImage once. None = unavailable."""
    arch = os.uname().machine
    if arch not in ("x86_64", "aarch64"):
        print(f"    appimagetool has no official build for {arch}")
        return None
    name = "appimagetool-x86_64.AppImage" if arch == "x86_64" else f"appimagetool-{arch}.AppImage"
    tool = cache_dir / name
    if tool.is_file() and os.access(tool, os.X_OK):
        return tool
    url = ("https://github.com/AppImage/appimagetool/releases/download/continuous/"
           + name)
    print(f"    Downloading {name} ...")
    try:
        cache_dir.mkdir(parents=True, exist_ok=True)
        urllib.request.urlretrieve(url, tool)
        tool.chmod(0o755)
        return tool
    except Exception as e:
        print(f"    Download failed ({e}); skipping AppImage")
        return None


def step4(version):
    step("Step 4/4 - Build AppImage")
    if NO_APPIMAGE:
        print("    Skipped (--no-appimage)")
        return
    if not DESKTOP_FILE.is_file() or not ICON_FILE.is_file():
        print("    Missing desktop file or icon; skipping AppImage")
        return

    cache = PROJECT_DIR / ".buildcache"
    tool = _ensure_appimagetool(cache)
    if tool is None:
        print("    Skipped AppImage (no appimagetool)")
        return

    with tempfile.TemporaryDirectory(prefix="pdfe_appimage_") as tmp:
        appdir = Path(tmp) / "PDFeverything.AppDir"
        (appdir / "usr" / "bin").mkdir(parents=True)
        shutil.copy2(OUTPUT_BIN, appdir / "usr" / "bin" / "PDFeverything")
        shutil.copy2(DESKTOP_FILE, appdir / "PDFeverything.desktop")
        shutil.copy2(ICON_FILE, appdir / "PDFeverything.png")
        # AppRun: run the bundled binary. #!/bin/bash is required because
        # `exec -a` (sets WM_CLASS) is a bash builtin, not available in dash.
        apprun = appdir / "AppRun"
        apprun.write_text(
            "#!/bin/bash\n"
            'exec -a PDFeverything "$(dirname "$0")/usr/bin/PDFeverything" "$@"\n',
            encoding="utf-8")
        apprun.chmod(0o755)
        # .DirIcon for appimaged / file managers
        shutil.copy2(ICON_FILE, appdir / ".DirIcon")

        out = DIST_DIR / f"PDFeverything_Setup_v{version}_linux.AppImage"
        # Write to a temp name first: a RUNNING instance of the previous
        # AppImage makes mksquashfs fail on the final path (its squashfs
        # mount holds the file), while os.replace over it is always allowed.
        stage = DIST_DIR / f".{out.name}.new"
        rc = run([str(tool), str(appdir), str(stage)], timeout=900)
        if rc != 0 or not stage.exists():
            # Show WHY mksquashfs refused (disk full, bad runtime, …).
            diag = subprocess.run(
                [str(tool), "--runtime-file", str(runtime),
                 str(appdir), str(stage)],
                capture_output=True, text=True, timeout=900)
            tail = (diag.stderr or diag.stdout or "").strip().splitlines()[-5:]
            for line in tail:
                print(f"      {line}")
            stage.unlink(missing_ok=True)
            print("    AppImage build failed; the plain binary is still valid")
            return
        os.replace(stage, out)
        out.chmod(0o755)
        print(f"    {out.name} - {out.stat().st_size / 1024 / 1024:.1f} MB")
    ok("AppImage complete")


def main():
    print_banner()
    version = read_version()
    print(f"    Version: {version}")
    step1()
    step2()
    step3()
    step4(version)
    print()
    print("=" * 60)
    print("    Build finished! Artifacts in dist/:")
    print(f"      PDFeverything                                  (single-file binary)")
    out_ai = DIST_DIR / f"PDFeverything_Setup_v{version}_linux.AppImage"
    if out_ai.exists():
        print(f"      {out_ai.name}          (AppImage)")
    print("=" * 60)


if __name__ == "__main__":
    main()
