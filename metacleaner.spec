# -*- mode: python ; coding: utf-8 -*-
# ╔══════════════════════════════════════════════════════════════════╗
# ║  metacleaner.spec — Archivo de configuración de PyInstaller     ║
# ║                                                                  ║
# ║  Para compilar:                                                  ║
# ║      build.bat                                                   ║
# ║  O manualmente:                                                  ║
# ║      pyinstaller metacleaner.spec --clean                        ║
# ╚══════════════════════════════════════════════════════════════════╝

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# ── Ruta raíz del proyecto ────────────────────────────────────────────────
BASE = os.path.abspath(os.path.dirname(SPEC))

# ══════════════════════════════════════════════════════════════════════════
#  DATOS — archivos de recursos que cada paquete necesita en runtime
# ══════════════════════════════════════════════════════════════════════════
datas = []

# customtkinter: temas JSON e imágenes internas (obligatorio)
datas += collect_data_files('customtkinter')

# tkinterdnd2: librería nativa TkDND (.dll en Windows, .so en Linux/Mac)
datas += collect_data_files('tkinterdnd2')

# openpyxl: plantillas XML para crear libros en blanco
datas += collect_data_files('openpyxl')

# python-docx: plantilla default.docx y esquemas XML
datas += collect_data_files('docx')

# pypdf: no tiene datos externos, pero la incluimos por si acaso
datas += collect_data_files('pypdf')

# ══════════════════════════════════════════════════════════════════════════
#  IMPORTS OCULTOS — módulos que PyInstaller no detecta por análisis estático
# ══════════════════════════════════════════════════════════════════════════
hiddenimports = [
    # GUI
    'customtkinter',
    'tkinterdnd2',
    'darkdetect',
    'packaging',
    'packaging.version',
    'packaging.specifiers',

    # Imágenes
    'PIL',
    'PIL.Image',
    'PIL.ExifTags',
    'PIL.ImageDraw',
    'PIL.ImageFont',
    'PIL._tkinter_finder',
    'PIL.JpegImagePlugin',
    'PIL.PngImagePlugin',
    'PIL.TiffImagePlugin',

    # PDF
    'pypdf',
    'pypdf._reader',
    'pypdf._writer',
    'pypdf._page',
    'pypdf.filters',
    'pypdf.generic',
    'pypdf.constants',

    # Word / Excel
    'docx',
    'docx.oxml',
    'docx.oxml.ns',
    'docx.oxml.shared',
    'openpyxl',
    'openpyxl.workbook',
    'openpyxl.worksheet',
    'openpyxl.styles',
    'openpyxl.utils',
    'openpyxl.reader.excel',
    'openpyxl.writer.excel',

    # XML (requerido por docx y openpyxl)
    'lxml',
    'lxml.etree',
    'lxml._elementpath',
    'lxml.html',

    # Estándar
    'email.mime.text',
    'email.mime.multipart',
]

# ── Submodules completos ──────────────────────────────────────────────────
hiddenimports += collect_submodules('docx')
hiddenimports += collect_submodules('openpyxl')

# ── Icono ─────────────────────────────────────────────────────────────────
_ico = os.path.join(BASE, 'assets', 'icon.ico')
_ico = _ico if os.path.isfile(_ico) else None

# ══════════════════════════════════════════════════════════════════════════
#  ANÁLISIS
# ══════════════════════════════════════════════════════════════════════════
a = Analysis(
    [os.path.join(BASE, 'app.py')],
    pathex=[BASE],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[os.path.join(BASE, 'hooks')],   # hooks locales opcionales
    hooksconfig={},
    runtime_hooks=[os.path.join(BASE, 'runtime_hook.py')],
    excludes=[
        # Excluir paquetes científicos que no se usan (reducen el tamaño)
        'matplotlib', 'numpy', 'pandas', 'scipy',
        'PyQt5', 'PyQt6', 'wx',
        'IPython', 'jupyter',
        'tkinter.test',
        'unittest',
        'test',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

# ══════════════════════════════════════════════════════════════════════════
#  EJECUTABLE — carpeta (portable, inicio más rápido que onefile)
# ══════════════════════════════════════════════════════════════════════════
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,     # Los binarios van en la carpeta COLLECT
    name='MetaCleaner',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,                  # Comprimir con UPX si está disponible
    console=False,             # Sin ventana de consola (app GUI)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=_ico,
    version_file=None,
)

# ══════════════════════════════════════════════════════════════════════════
#  COLECCIÓN — carpeta dist/MetaCleaner lista para distribuir
# ══════════════════════════════════════════════════════════════════════════
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[
        'vcruntime140.dll',
        'python3*.dll',
        'tk*.dll',
        'tcl*.dll',
    ],
    name='MetaCleaner',
)
