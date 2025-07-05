# -*- mode: python ; coding: utf-8 -*-
block_cipher = None

a = Analysis(
    ['gui_tk.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('categories.json', '.'),
        ('accounts_to.json', '.'),
        ('bank2zen.py', '.')
    ],
    hiddenimports=['openpyxl'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name="bank2zen",
    console=False,
    icon=None,
)
