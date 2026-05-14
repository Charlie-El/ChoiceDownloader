# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['scripts\\choice_automation_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['openpyxl'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'IPython', 'jedi', 'pygments', 'zmq', 'notebook', 'nbformat', 'pytest', 'yapf', 'gevent', 'cloudpickle', 'PIL', 'numpy', 'scipy', 'pandas', 'cryptography', 'sqlite3'],
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
    name='ChoiceDownloader',
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
)
