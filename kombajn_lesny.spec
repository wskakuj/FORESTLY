# -*- mode: python ; coding: utf-8 -*-
import sys
import os
from PyInstaller.utils.hooks import collect_submodules, collect_data_files, collect_dynamic_libs

block_cipher = None

hidden_imports = collect_submodules('app')
hidden_imports += collect_submodules('customtkinter')
hidden_imports += collect_submodules('CTkToolTip')
hidden_imports += collect_submodules('pandas')
hidden_imports += collect_submodules('numpy')
hidden_imports += collect_submodules('fitz')
hidden_imports += collect_submodules('openpyxl')
hidden_imports += collect_submodules('docx')
hidden_imports += collect_submodules('PIL')
hidden_imports += collect_submodules('pypdf')
hidden_imports += collect_submodules('pyodbc')
hidden_imports += [
    'win32com.client',
    'pythoncom',
    'pyautogui',
    'pandas._libs',
    'pandas._libs.tslibs',
    'pandas._libs.tslibs.timedeltas',
    'pandas._libs.tslibs.nattype',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.strptime',
    'numpy.core',
    'numpy.core._multiarray_umath',
    'PIL.Image',
    'PIL.ImageDraw',
    'fitz',
]

datas = [
    ('STR_TYT.docx', '.'),
    ('STR_TYT_TYLKO-ISL-2.docx', '.'),
    ('Skroty.docx', '.'),
    ('BIAŁYNIN KRASÓWKA.xlsx', '.'),
    ('config', 'config'),
    ('pusty', 'pusty'),
]
datas += collect_data_files('customtkinter')
datas += collect_data_files('CTkToolTip')

binaries = []
binaries += collect_dynamic_libs('pandas')
binaries += collect_dynamic_libs('numpy')
binaries += collect_dynamic_libs('fitz')

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hidden_imports,
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='KombajnLesnyPRO',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='kombajn.ico' if os.path.exists('kombajn.ico') else None,
)
