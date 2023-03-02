# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['Report2Excel_0220.py'],
    pathex=['D:\\R2E\\venv\\'],
    binaries=[],
    datas=[('D:\\R2E\\venv\\lib\\site-packages\\sklearn', 'sklearn'),('D:\\R2E\\venv\\lib\\site-packages\\matplotlib', 'matplotlib')],
    hiddenimports=['matplotlib.pyplot','seaborn',],
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
    name='Report2Excel_0220',
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
)
