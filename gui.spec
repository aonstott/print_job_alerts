# -*- mode: python ; coding: utf-8 -*-

import os

downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
a = Analysis(
    ['gui.py'],
    pathex=['C:\\Users\aonstott\Downloads\print_job_alerts'],
    binaries=[],
    datas=[('main.py', '.'), ('Emailer.py', '.'), ('address_reader.py', '.'), ('email_addresses.cfg', '.'), ('group_leaders.cfg', '.'), ('groups.cfg', '.'), ('JobGroup.py', '.'), ('excel/*', 'excel')],
    hiddenimports=[],
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
    name='gui',
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
    icon='./favicon.ico'
)