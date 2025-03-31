# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('config.ini', '.'), ('README.md', '.'), ('email.png', '.'), ('C:\\Users\\ND\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\PyQt5\\Qt5\\bin', 'PyQt5/Qt5/bin'), ('C:\\Users\\ND\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\PyQt5\\Qt5\\plugins', 'PyQt5/Qt5/plugins'), ('C:\\Users\\ND\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\qt_material\\resources', 'qt_material/resources'), ('C:\\Users\\ND\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\docx', 'docx')]
binaries = []
hiddenimports = ['PyQt5.sip', 'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets', 'lxml._elementpath', 'lxml.etree']
tmp_ret = collect_all('lxml')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt6', 'PySide6', 'PySide2'],
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
    name='邮件群发工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt',
    uac_admin=True,
    icon=['email.png'],
)
