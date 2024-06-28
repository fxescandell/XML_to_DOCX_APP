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
    hiddenimports=['docx', 'docx.oxml', 'docx.oxml.ns', 'docx.oxml.text', 'docx.enum.style', 'docx.parts.document', 'docx.parts.numbering', 'docx.parts.styles'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure)

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
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch='x86_64',
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main'
)

app = BUNDLE(
    coll,
    name='main.app',
    icon='resources/icon.icns',
    bundle_identifier='com.yourdomain.main',
    target_arch='x86_64',
    info_plist={
        'CFBundleName': 'MainApp',
        'CFBundleDisplayName': 'MainApp',
        'CFBundleGetInfoString': 'MainApp',
        'CFBundleIdentifier': 'com.yourdomain.main',
        'CFBundleVersion': '0.1.0',
        'CFBundleShortVersionString': '0.1.0',
        'NSHighResolutionCapable': 'True'
    }
)
