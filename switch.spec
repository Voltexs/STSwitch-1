# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['switch.py'],
    pathex=[],
    binaries=[],
    datas=[('app.ico', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Remove unnecessary Qt modules
a.binaries = [x for x in a.binaries if not x[0].startswith("Qt6.Qt")]
a.binaries = [x for x in a.binaries if not x[0].startswith("Qt6.QtNetwork")]
a.binaries = [x for x in a.binaries if not x[0].startswith("Qt6.QtPositioning")]
a.binaries = [x for x in a.binaries if not x[0].startswith("Qt6.QtQml")]

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PatchSwitcher',
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
    icon='app.ico',
    uac_admin=True,  # This requests admin privileges
) 