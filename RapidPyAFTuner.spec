# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['RapidPy\\af_tuner\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('RapidPy/af_tuner/assets', 'assets')],
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
    name='RapidPyAFTuner',
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
    icon=['RapidPy\\af_tuner\\assets\\af_tuner_icon.ico'],
)
