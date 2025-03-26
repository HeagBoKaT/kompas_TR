# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('icons/icon.ico', 'icons'),  # Добавляем иконку как ресурс
        # Если используются другие файлы (например, templates.json, settings.json), добавьте их:
        # ('templates.json', '.'),
        # ('settings.json', '.'),
    ],
    hiddenimports=[
        # PyQt6 модули
        'PyQt6',
        'PyQt6.QtWidgets',
        'PyQt6.QtGui',
        'PyQt6.QtCore',
        # COM-объекты для работы с KOMPAS-3D
        'win32com',
        'win32com.client',
        'pythoncom',
        # Дополнительные модули из win32com
        'win32com.client.Dispatch',
        'win32com.client.gencache',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='KOMPAS-TR',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # Убедитесь, что UPX установлен, иначе установите False
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Отключаем консоль для GUI-приложения
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icons\\icon.ico'],  # Иконка для .exe
)