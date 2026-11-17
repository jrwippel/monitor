# -*- mode: python ; coding: utf-8 -*-
# Build do Monitor de Produtividade (versão integrada ao TimeSheet).
# Gera um único .exe, sem console, com ícone do relógio.
#   Compilar:  pyinstaller app_timesheet_integrado.spec

a = Analysis(
    ['app_timesheet_integrado.py'],
    pathex=[],
    binaries=[],
    datas=[('assets', 'assets')],
    # win32crypt (login guardado), win32event/winerror (instância única): importados dentro de funções
    hiddenimports=['PIL._tkinter_finder', 'win32crypt', 'win32event', 'winerror'],
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
    name='MonitorTimeSheet',
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
    icon='timesheet.ico',
)
