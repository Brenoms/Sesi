# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import sys

from PyInstaller.utils.hooks import collect_submodules

project_dir = Path(SPECPATH)
python_base = Path(sys.base_prefix)
tcl_root = python_base / "tcl"

datas = [
    (str(project_dir / "selectors_bookings_sesi.json"), "."),
    (str(project_dir / "bookings_jobs_exemplo.json"), "."),
]

if (tcl_root / "tcl8.6").exists():
    datas.append((str(tcl_root / "tcl8.6"), "_tcl_data"))
if (tcl_root / "tk8.6").exists():
    datas.append((str(tcl_root / "tk8.6"), "_tk_data"))
if (python_base / "Lib" / "tkinter").exists():
    datas.append((str(python_base / "Lib" / "tkinter"), "tkinter"))

hiddenimports = ["_tkinter"]

block_cipher = None

a = Analysis(
    ["bookings_site_automator.py"],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    name="BookingsSESI",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="BookingsSESI",
)
