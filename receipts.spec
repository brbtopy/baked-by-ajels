# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['receipts.py'],
    pathex=['C:/Users/samas/OneDrive/Desktop/proj/ajels/baked-by-ajels-receipt'],
    binaries=[],
    datas=[],
    hiddenimports=['reportlab', 'openpyxl', 'num2words'],
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
    name='AJELS RECEIPT',
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
    icon='C:/Users/samas/OneDrive/Desktop/proj/ajels/baked-by-ajels-receipt/pictures/ico_logo.ico'  # Specify the path to your custom icon file
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AJELS RECEIPT',
)