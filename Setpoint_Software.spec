# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['D:\\Setpoint_Project\\Setpoint_Software\\Setpoint_Project_GUI.py', 'D:\\Setpoint_Project\\Setpoint_Software\\hook-gcloud.py', 'D:\\Setpoint_Project\\Setpoint_Software\\set-point-attender-firebase.json'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=['D:\\Setpoint_Project\\Setpoint_Software'],
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
    name='Setpoint_Software',
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
    icon=['D:\\Setpoint_Project\\Setpoint_Software\\assets\\setpoint_logo_icon.ico'],
)
