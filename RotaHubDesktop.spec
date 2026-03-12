# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('C:\\Users\\fisca\\AppData\\Local\\Programs\\Python\\Python314\\Lib\\tkinter', 'tkinter'), ('assets', 'assets'), ('certificados', 'certificados'), ('config', 'config'), ('C:\\Users\\fisca\\AppData\\Local\\Programs\\Python\\Python314\\tcl\\tcl8.6', 'tcl\\tcl8.6'), ('C:\\Users\\fisca\\AppData\\Local\\Programs\\Python\\Python314\\tcl\\tk8.6', 'tcl\\tk8.6')]
binaries = []
hiddenimports = ['pandas', 'openpyxl', 'xlrd', 'tkinter', '_tkinter']
tmp_ret = collect_all('pandas')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('xlrd')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['scripts\\pyi_rth_tkinter.py'],
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
    name='RotaHubDesktop',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['assets\\app_icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='RotaHubDesktop',
)
