# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec para ReminderMailAgua v2
# Uso: pyinstaller reminderagua.spec --clean
# Salida: dist/reminderagua.exe

block_cipher = None

a = Analysis(
    ['main.py'],                    # Punto de entrada
    pathex=[],
    binaries=[],
    datas=[
        # Bundlear la carpeta de traducciones junto al ejecutable
        ('locales', 'locales'),
        # Incluir config.json con valores por defecto
        ('config.json', '.'),
        # Incluir el ícono
        ('reminderagua.ico', '.'),
    ],
    hiddenimports=[
        # win32com.client no es detectado automáticamente por PyInstaller
        'win32com.client',
        'win32api',
        'win32con',
        'pywintypes',
        'pythoncom',
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
    name='reminderagua',            # Nombre del .exe generado
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,                       # Comprimir con UPX para reducir tamaño
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                  # Sin ventana de consola (aplicación de escritorio)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='reminderagua.ico',        # Ícono de la aplicación
)
