# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=[('snippets.json', '.'), ('app.ico', '.'), ('D:\\Python\\Lib\\site-packages\\customtkinter', 'customtkinter'), ('templates', 'templates'), ('assets', 'assets'), ('ui/themes', 'ui/themes')],
    hiddenimports=['PIL._tkinter_finder', 'babel.numbers', 'win32timezone', 'docx', 'customtkinter', 'openai'],
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
    name='MarkdownToWord',
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
    icon=['app.ico'],
)
