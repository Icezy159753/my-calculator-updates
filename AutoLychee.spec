# -*- mode: python ; coding: utf-8 -*-
# Separate PySide6 onedir bundle: avoid extracting Qt on every launch.
from PyInstaller.utils.hooks import collect_submodules

source = "All_Programs/158_AutoLychee_OneFile.py"
hiddenimports = [
    'PySide6.QtCore', 'PySide6.QtGui', 'PySide6.QtWidgets',
    'pythoncom', 'pywintypes', 'win32com.client', 'win32com.client.dynamic',
    'win32job', 'win32gui', 'win32process', 'win32api', 'win32con',
    'openpyxl', 'lxml.etree', 'comtypes.client', 'PIL.ImageGrab',
]
hiddenimports += collect_submodules('pywinauto')
hiddenimports += collect_submodules('comtypes')

a = Analysis(
    [source],
    pathex=[],
    binaries=[],
    datas=[(source, '.')],  # Loader reads the embedded sections from this source.
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'PyQt5', 'PyQt6', 'PySide2',
        # These modules are served by the one-file section importer at runtime.
        'onefile', 'core', 'fast_styles', 'history', 'chrome', 'lyche_windows', 'driver', 'sounds',
        'post_total_na', 'post_del_sig', 'post_cut_percent', 'worker', 'app', 'onefile_assets',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name='AutoLychee',
    debug=False,
    strip=False,
    upx=False,
    console=False,
    icon='Icon/Autolychee.png',
)
coll = COLLECT(
    exe, a.binaries, a.datas,
    strip=False,
    upx=False,
    name='AutoLychee',
    contents_directory='_internal',
)
