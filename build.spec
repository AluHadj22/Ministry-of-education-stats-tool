# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['your_script.py'],  # замените на имя вашего файла
    pathex=[],
    binaries=[],
    datas=[
        ('C:/Python310/Lib/site-packages/pandas/lib/*', 'pandas/lib/'),
        ('C:/Python310/Lib/site-packages/pandas/_libs/*', 'pandas/_libs/')
    ],
    hiddenimports=[
        'pandas',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.np_datetime',
        'tkinter',
        'matplotlib.backends.backend_tkagg',
        'matplotlib.pyplot',
        'openpyxl',
        'openpyxl.worksheet',
        'openpyxl.chart'
    ],
    hookspath=[],
    hooksconfig={},
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='DistrictAnalyzer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Измените на True если нужна консоль для отладки
    icon=None,  # Можете добавить иконку .ico файл
)
