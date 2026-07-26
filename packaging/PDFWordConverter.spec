# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
from PyInstaller.utils.hooks import collect_all


project_root = Path(SPECPATH).parent
datas = [(str(project_root / "LICENSE"), ".")]
binaries = []
hiddenimports = ["app_environment", "batch_logic"]

for package in ("pymupdf", "docx"):
    package_datas, package_binaries, package_hiddenimports = collect_all(package)
    datas += package_datas
    binaries += package_binaries
    hiddenimports += package_hiddenimports

comtypes_datas, comtypes_binaries, comtypes_hiddenimports = collect_all(
    "comtypes", filter_submodules=lambda name: not name.startswith("comtypes.test")
)
datas += comtypes_datas
binaries += comtypes_binaries
hiddenimports += comtypes_hiddenimports

a = Analysis(
    [str(project_root / "launcher.py")],
    pathex=[str(project_root)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

portable = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="PDFWordConverter-v0.3.0-Portable-x64",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch="x86_64",
    version=str(project_root / "packaging" / "version_info.txt"),
    manifest=str(project_root / "packaging" / "app.manifest"),
)

directory_exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PDFWordConverter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch="x86_64",
    version=str(project_root / "packaging" / "version_info.txt"),
    manifest=str(project_root / "packaging" / "app.manifest"),
)

COLLECT(
    directory_exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="PDFWordConverter-v0.3.0",
)
