# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['src\\main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
    ('resources\\icono.ico', 'resources'),
    ('resources\\Logo_CHC_new.png', 'resources'),
    ('resources\\gear_icon_b.png', 'resources'),
    ('.env', '.'),
    ('template\\plantilla_Excel.xlsx', 'template'),
    ('template\\plantilla_OLO.xlsx', 'template'),
    ('template\\plantilla_PLASTIMEX.xlsx', 'template'),
    ('template\\plantilla_ALPEMUSA.xlsx', 'template'),
    ('template\\plantilla_SAN_NICOLAS.xlsx', 'template'),
    ],
    hiddenimports=[
        'tkinter',
        'msal',
        'requests',
        'PIL.Image',
        'PIL.ImageTk',
        'tkhtmlview'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter.test',
        'tkinter.ttk.test',
        'unittest',
    ],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='CHC_Reportes',
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
    icon='resources\\icono.ico',
    manifest="""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
        <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
            <security>
                <requestedPrivileges>
                    <requestedExecutionLevel level="requireAdministrator" uiAccess="false"/>
                </requestedPrivileges>
            </security>
        </trustInfo>
    </assembly>"""
)
