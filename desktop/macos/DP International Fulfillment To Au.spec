# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import collect_all


project_root = Path(SPECPATH).resolve().parents[1]
datas = []
binaries = []
hiddenimports = []
for package in ("openpyxl", "xlrd", "pypdf"):
    package_data, package_binaries, package_hiddenimports = collect_all(package)
    datas += package_data
    binaries += package_binaries
    hiddenimports += package_hiddenimports

a = Analysis(
    ["app.py"],
    pathex=[str(project_root), str(project_root / "shared" / "fulfillment")],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="DP International Fulfillment To Au",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    target_arch="arm64",
)
collection = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="DP International Fulfillment To Au",
)
app = BUNDLE(
    collection,
    name="DP International Fulfillment To Au.app",
    icon=None,
    bundle_identifier="com.draper.dp-international-fulfillment",
    info_plist={"NSHighResolutionCapable": True},
)
