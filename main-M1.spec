# main.spec
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['/Users/javi/Code/XML_to_DOCX_APP'],
    binaries=[],
    datas=[
        ('styles_config.json', '.'),
        ('utils.py', '.'),
        ('resources/icon.icns', 'resources')
    ],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
)
coll = BUNDLE(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='main.app',
    icon=None,
    bundle_identifier=None,
    info_plist=None,
    codesign_identity=None,
    entitlements_file=None,
)
