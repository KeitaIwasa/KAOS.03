# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(
    ['KAOS.3.7.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('Automation.py', '.'), 
        ('setup/KAOS_icon.ico', 'setup'),
        ('setup/sheet_icon.png', 'setup'),
        ('setup/setting_icon.png', 'setup'),
        ('setup/KAOS_Support_QR_resized.png', 'setup')
    ],
    hiddenimports=[
        'chardet',
        'charset_normalizer',
        'charset_normalizer.md__mypyc', # Include specific module causing error
        'requests',
        'idna',
        'certifi',
        'urllib3',
    ],
    hookspath=[],
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
    [],
    exclude_binaries=True,
    name='KAOS',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='setup/KAOS_icon.ico',
    version='app.version'
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='KAOS',
)
