import os

block_cipher = None

a = Analysis(
    ['bot_grammer.py'],
    pathex=[os.getcwd()],  # Set the current directory as the path
    binaries=[],
    datas=[
        ('needyamin.ico', '.'), 
        ('logo_display.png', '.'), 
        ('logo_display_computer.png', '.'), 
        ('computer.ico', '.'), 
        ('load.gif', '.'), 
        ('data.sqlite', '.'), 
        ('error.log', '.')
    ],
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name='SpeechAssistant',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Set to True if you want a console window
    icon='needyamin.ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SpeechAssistant',
)