# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import sys

project_dir = Path(SPECPATH)
python_base = Path(sys.base_prefix)
tcl_root = python_base / "tcl"

datas = [
    (str(project_dir / "bookings_sesi.ico"), "."),
    (str(project_dir / "imagens" / "sesi_logo_app.png"), "."),
    (str(project_dir / "sesi_logo_recortada.png"), "."),
    (str(project_dir / "imagens" / "sesi_logo_vermelha.png"), "."),
    (str(Path.home() / "AppData" / "Local" / "Temp" / "BookingsSESI_InstallerBuild" / "payload.zip"), "."),
]

if (python_base / "Lib" / "tkinter").exists():
    datas.append((str(python_base / "Lib" / "tkinter"), "tkinter"))
if (tcl_root / "tcl8.6").exists():
    datas.append((str(tcl_root / "tcl8.6"), "_tcl_data"))
if (tcl_root / "tk8.6").exists():
    datas.append((str(tcl_root / "tk8.6"), "_tk_data"))

a = Analysis(
    ["installer_gui.py"],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=["_tkinter"],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="BookingsSESI-Instalador",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=str(project_dir / "bookings_sesi.ico"),
)
