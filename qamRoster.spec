# -*- mode: python ; coding: utf-8 -*-

import json
from pathlib import Path

from PyInstaller.utils.win32.versioninfo import (
    VSVersionInfo,
    FixedFileInfo,
    StringFileInfo,
    StringTable,
    StringStruct,
    VarFileInfo,
    VarStruct,
)

project_root = Path(globals().get("SPECPATH", ".")).resolve()
project_metadata = json.loads((project_root / 'project.json').read_text(encoding='utf-8'))
version = project_metadata.get('version', '0.0.0').strip() or '0.0.0'
version_parts = [int(part) for part in version.split('.')]
while len(version_parts) < 4:
    version_parts.append(0)
version_tuple = tuple(version_parts[:4])

version_info = VSVersionInfo(
    ffi=FixedFileInfo(
        filevers=version_tuple,
        prodvers=version_tuple,
        mask=0x3F,
        flags=0x0,
        OS=0x40004,
        fileType=0x1,
        subtype=0x0,
        date=(0, 0),
    ),
    kids=[
        StringFileInfo(
            [
                StringTable(
                    '040904B0',
                    [
                        StringStruct('CompanyName', project_metadata.get('author', '')),
                        StringStruct('FileDescription', project_metadata.get('purpose', '')),
                        StringStruct('FileVersion', version),
                        StringStruct('InternalName', 'qamRoster'),
                        StringStruct('OriginalFilename', 'qamRoster.exe'),
                        StringStruct('ProductName', project_metadata.get('name', 'qamRoster')),
                        StringStruct('ProductVersion', version),
                    ],
                )
            ]
        ),
        VarFileInfo([VarStruct('Translation', [1033, 1200])]),
    ],
)


a = Analysis(
    ['main.py'],
    pathex=[str(project_root)],
    binaries=[],
    datas=[
        (str(project_root / 'modules' / 'credentials.json'), 'modules'),
        (str(project_root / 'project.json'), '.'),
        (str(project_root / 'QAM-Logo-1-2048x1310whiteBGRND.png'), '.'),
    ],
    hiddenimports=[],
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
    name='qamRoster',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon=str(project_root / 'QAM-Logo-1-2048x1310whiteBGRND.ico'),
    version=version_info,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
