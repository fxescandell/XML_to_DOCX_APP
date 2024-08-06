# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

from PyInstaller.utils.hooks import collect_submodules

# Colecta todos los submódulos del paquete `docx`
hiddenimports = collect_submodules('docx')

a = Analysis(
    ['main.py'],  # Archivo principal
    pathex=['Z:\\Users\\imac\\Documents\\Code\\XML_to_DOCX_APP-main'],  # Ruta del proyecto
    binaries=[],  # Archivos binarios adicionales
    datas=[
        ('styles_config.json', '.'),  # Archivo de configuración JSON
        ('utils.py', '.'),  # Archivo adicional de utilidades
    ],
    hiddenimports=hiddenimports,  # Inclusión de todos los submódulos de `docx`
    hookspath=[],  # Scripts hook personalizados
    runtime_hooks=[],  # Scripts de gancho en tiempo de ejecución
    excludes=[],  # Módulos específicos a excluir
    win_no_prefer_redirects=False,  # Preferencias específicas de Windows
    win_private_assemblies=False,  # Uso de ensamblados privados en Windows
    cipher=block_cipher,  # Cifrado para los scripts
    noarchive=False,  # Incluir scripts en el archivo ZIP
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,  # Excluir binarios adicionales
    name='main.exe',  # Nombre del ejecutable
    debug=False,  # Modo depuración
    bootloader_ignore_signals=False,  # Ignorar señales del bootloader
    strip=False,  # No eliminar símbolos de depuración
    upx=True,  # Comprimir con UPX
    console=True,  # Mostrar consola
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,  # No eliminar símbolos de depuración
    upx=True,  # Comprimir con UPX
    upx_exclude=[],  # Excluir archivos de UPX
    name='main'
)
