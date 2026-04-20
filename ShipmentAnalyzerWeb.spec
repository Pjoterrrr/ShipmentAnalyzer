# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules, copy_metadata

project_dir = Path.cwd()

datas = [
    (str(project_dir / "app.py"), "."),
    (str(project_dir / "assets"), "assets"),
    (str(project_dir / ".streamlit"), ".streamlit"),
    (str(project_dir / "config"), "config"),
] + copy_metadata("streamlit") + copy_metadata("altair") + copy_metadata("pandas")

hiddenimports = (
    collect_submodules("streamlit")
    + collect_submodules("altair")
)

a = Analysis(
    ["launcher.py"],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ShipmentAnalyzerWeb",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=str(project_dir / "assets" / "icon.ico"),
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="ShipmentAnalyzerWeb",
)
