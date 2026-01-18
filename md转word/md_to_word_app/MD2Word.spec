# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['G:\\Learn\\work_notes\\md转word\\md_to_word_app\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('G:\\Learn\\work_notes\\md转word\\md_to_word_app\\assets', 'assets'), ('S:/Tools/Miniconda/envs/word-tools/Library/plugins', 'PyQt5/Qt5/plugins')],
    hiddenimports=['PyQt5.sip', 'docx', 'docx.oxml', 'docx.oxml.ns', 'lxml', 'lxml.etree', 'lxml._elementpath'],
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
    name='MD2Word',
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
)
