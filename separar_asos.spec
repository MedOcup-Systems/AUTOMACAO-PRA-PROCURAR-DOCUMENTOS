# -*- mode: python ; coding: utf-8 -*-
# Arquivo de configuracao do PyInstaller para o Separador de ASOs
# Execute: pyinstaller separar_asos.spec

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Coleta todos os dados necessarios das bibliotecas
datas = []
datas += collect_data_files('pdfplumber')
datas += collect_data_files('pdfminer')

# Coleta submodulos que o PyInstaller pode nao detectar automaticamente
hiddenimports = []
hiddenimports += collect_submodules('pdfplumber')
hiddenimports += collect_submodules('pdfminer')
hiddenimports += collect_submodules('PIL')
hiddenimports += [
    'PIL._tkinter_finder',
    'openpyxl',
    'openpyxl.cell',
    'openpyxl.styles',
    'openpyxl.utils',
    'pypdf',
    'pytesseract',
    'pdf2image',
    'unicodedata',
    'queue',
    'threading',
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
]

a = Analysis(
    ['separar_asos_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy',
        'IPython', 'jupyter', 'notebook',
        'PyQt5', 'PyQt6', 'wx',
        'test', 'unittest',
    ],
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
    name='Separador_ASOs',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,        # False = sem janela de terminal (so a GUI)
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='icone.ico',   # descomente e coloque um .ico se quiser icone proprio
)
