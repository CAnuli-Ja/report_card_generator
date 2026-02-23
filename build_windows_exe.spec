# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for Windows executable
# Run this on Windows with: pyinstaller build_windows_exe.spec

a = Analysis(
    ['report_card_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pandas', 'docx', 'tkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludedimports=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Report_Card_Generator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
