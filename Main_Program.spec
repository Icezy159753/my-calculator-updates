# -*- mode: python ; coding: utf-8 -*-
import os
import sysconfig
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_dynamic_libs
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all

site_packages = sysconfig.get_paths()["purelib"]
datas = [
    (os.path.join(site_packages, "savReaderWriter"), "savReaderWriter"),
    (os.path.join(site_packages, "sv_ttk"), "sv_ttk"),
    ("Icon", "Icon"),
    ("All_Programs", "All_Programs"),
    ("openrouter.json", "."),
    ("template.xlsx", "."),
    ("Test3.json", "."),
    ("Itemdef - Format.xlsx", "."),
    ("savReaderWriter", "savReaderWriter"),
]

try:
    import matplotlib
    mpl_data = matplotlib.get_data_path()
    datas.append((mpl_data, "matplotlib/mpl-data"))
except Exception:
    pass

datas = [item for item in datas if os.path.exists(item[0])]
binaries = []
hiddenimports = ['statsmodels.api', 'ttkbootstrap.tableview', 'xlsxwriter', 'scipy._cyutility', 'ttkbootstrap.scrolled', 'ttkbootstrap', 'openpyxl.utils.dataframe', 'pyspssio', 'natsort', 'docx', 'openpyxl', 'pyperclip', 'tkinter.scrolledtext', 'xlswriter', 'pandas', 'seaborn', 'matplotlib.backends.backend_tkagg', 'seaborn', 'pandas', 'numpy', 'xlsxwriter', 'pyreadstat', 'pyreadstat._readstat_writer', 'worker', 'PyQt6', 'PyQt6.QtCore', 'PyQt6.QtGui', 'PyQt6.QtWidgets', 'PyQt6.QtPrintSupport', 'sip', 'gspread', 'google.auth', 'google.auth.transport.requests', 'google.oauth2.service_account', 'google_auth_oauthlib', 'googleapiclient.discovery']
datas += collect_data_files('pyspssio')
datas += collect_data_files('googleapiclient')
datas += collect_data_files('certifi')
binaries += collect_dynamic_libs('savReaderWriter')
binaries += collect_dynamic_libs('savReaderWriter.spssio')
datas += collect_data_files('savReaderWriter.spssio')
binaries += collect_dynamic_libs('pyspssio')
hiddenimports += collect_submodules('savReaderWriter')
hiddenimports += collect_submodules('ttkbootstrap')
hiddenimports += collect_submodules('factor_analyzer')
hiddenimports += collect_submodules('scipy')
hiddenimports += collect_submodules('sklearn')
hiddenimports += collect_submodules('ttkbootstrap')
hiddenimports += collect_submodules('customtkinter')
hiddenimports += collect_submodules('pyspssio')
hiddenimports += collect_submodules('pyreadstat')
hiddenimports += collect_submodules('pyperclip')
tmp_ret = collect_all('statsmodels')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('sklearn')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('gspread')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('PyQt6')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['Main_Program.py'],
    pathex=[os.path.abspath("All_Programs")],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Main_Program',
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
    icon=['I_Main.ico','Icon\\I_Main.ico'],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Main_Program',
    contents_directory='_internal',
)
