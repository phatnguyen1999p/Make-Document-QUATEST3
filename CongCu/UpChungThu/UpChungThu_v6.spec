# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('GIAY TIEU DE.pdf', '.'),
        ('GIAY TIEU DE - E.pdf', '.'),
        ('settings_source.json', '.')
    ],
    hiddenimports=[
        'core',
        'core.config_manager',
        'core.utils',
        'core.pdf_service',
        'core.word_service',
        'core.excel_service',
        'core.mapping_service',
        'gui',
        'gui.theme',
        'gui.app',
        'gui.dialogs',
        'gui.dialogs.signer_dialog',
        'gui.dialogs.time_rules_dialog',
        'gui.dialogs.calendar_dialog',
        'gui.dialogs.bulk_config_dialog',
        'gui.dialogs.status_dialog',
        'gui.components',
        'gui.components.base_processor',
        'gui.views',
        'gui.views.main_menu',
        'gui.views.word_view',
        'gui.views.excel_view',
        'win32com.client',
        'pythoncom',
        'PyPDF2'
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
    [],
    exclude_binaries=True,
    name='UpChungThu_v6',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='UpChungThu_v6',
)
