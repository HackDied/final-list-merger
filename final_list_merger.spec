# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec dosyası - HIZLI BAŞLATMA (onedir mod)
# Kullanım: pyinstaller final_list_merger.spec

import os
import customtkinter

ctk_path = os.path.dirname(customtkinter.__file__)

# Gereksiz modülleri dışla (boyut ve hız için)
EXCLUDES = [
    'pytest', 'py', 'pygments', 'lxml',
    'setuptools', 'sqlite3', 'unittest',
    'xmlrpc', 'pydoc', 'doctest',
    'matplotlib', 'scipy', 'IPython',
]

a = Analysis(
    ['final_list_merger.py'],
    pathex=[],
    binaries=[],
    datas=[
        (ctk_path, 'customtkinter'),
    ],
    hiddenimports=[
        'openpyxl',
        'pandas',
        'customtkinter',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludedimports=EXCLUDES,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Final List Merger Pro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
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
    name='Final List Merger Pro',
)
