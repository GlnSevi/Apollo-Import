# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_data_files

release_version = os.environ.get('APOLLO_RELEASE_VERSION', 'v0.1.6')
exe_name = f'Apollo-Import-GUI-{release_version}-onefile'

datas = []
datas += collect_data_files('playwright')
datas += collect_data_files('pymupdf')
datas += [
    ('assets\\apollo_import_logo.png', 'assets'),
    ('assets\\apollo_import_logo.ico', 'assets'),
]


a = Analysis(
    ['apollo_import_gui.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['anthropic'],
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
    a.binaries,
    a.datas,
    [],
    name=exe_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='assets\\apollo_import_logo.ico',
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
