# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import copy_metadata

datas = [('duty_app.py', '.'), ('logic.py', '.')]
hiddenimports = ['streamlit.web.cli', 'streamlit.runtime.scriptrunner.magic_funcs', 'numpy._core._exceptions', 'openpyxl', 'xlsxwriter', 'pdfplumber', 'tkinter', 'tkinter.filedialog']
datas += collect_data_files('streamlit')
datas += collect_data_files('altair')
datas += copy_metadata('streamlit')
datas += copy_metadata('altair')
datas += copy_metadata('pandas')
datas += copy_metadata('numpy')
datas += copy_metadata('pyarrow')
datas += copy_metadata('pdfplumber')
datas += copy_metadata('openpyxl')
datas += copy_metadata('xlsxwriter')
datas += copy_metadata('watchdog')
hiddenimports += collect_submodules('streamlit')


a = Analysis(
    ['run_main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'tensorflow', 'scipy', 'sklearn', 'statsmodels', 'matplotlib', 'pytest', 'IPython', 'jupyter', 'notebook', 'numba', 'plotly', 'bokeh', 'sqlalchemy', 'duckdb', 'polars', 'dask', 'pyspark', 'pandas.tests', 'pyarrow.tests', 'numpy.tests'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='DutyScheduler',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='DutyScheduler',
)
