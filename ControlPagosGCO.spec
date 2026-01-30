# -*- mode: python ; coding: utf-8 -*-

"""
ARCHIVO DE CONFIGURACIÓN PYINSTALLER - VERSIÓN CARPETA
Este archivo crea una carpeta con el ejecutable y dependencias.
Ventajas: Inicia más rápido, mejor para debugging
Uso: pyinstaller ControlPagos_Dir.spec
"""

block_cipher = None

a = Analysis(
    ['control_pagos_1_1.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Agregar archivos adicionales si es necesario
        # Ejemplo: ('icon.ico', '.'),
    ],
    hiddenimports=[
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'win32com.gen_py',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.utils.dataframe',
        'pandas',
        'pandas.core',
        'tkinter',
        'tkcalendar',
        'locale',
        'datetime',
        'pathlib',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'numpy.random._examples',
        'PIL',
        'PyQt5',
        'PySide2',
        'IPython',
        'jupyter',
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
    [],
    exclude_binaries=True,  # Diferencia principal: dependencias separadas
    name='ControlPagosGCO',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ControlPagosGCO',
)