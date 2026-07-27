# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
from PyInstaller.utils.hooks import collect_all


project_root = Path(SPECPATH).parent
datas = [
    (str(project_root / "LICENSE"), "."),
    (str(project_root / "THIRD_PARTY_NOTICES.md"), "."),
    (str(project_root / "packaging" / "licenses"), "licenses"),
]
binaries = []
hiddenimports = [
    "app_environment",
    "batch_logic",
    "conversion_specs",
    "drop_logic",
    "engine_models",
    "macos_office",
    "platform_services",
]

for package in ("pymupdf", "docx", "pptx", "PIL"):
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
    name="PDF-Word-PPT批量转换工具-v0.5.0-便携版-x64",
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
