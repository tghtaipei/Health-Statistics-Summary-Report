# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 配置檔
用於打包衛生統計報表 PDF 轉換工具
"""

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('excel_to_pdf_with_bookmarks.py', '.'),
        ('toc_generator.py', '.'),
        ('cover.png', '.'),
        ('additionalinfo.png', '.'),
    ],
    hiddenimports=[
        'win32com',
        'win32com.client',
        'pypdf',
        'reportlab',
        'reportlab.pdfgen',
        'reportlab.lib',
        'reportlab.pdfbase',
        'reportlab.pdfbase.ttfonts',
        'reportlab.lib.pagesizes',
        'reportlab.lib.utils',
        'reportlab.lib.colors',
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
    name='衛生統計報表PDF轉換工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # 關閉 UPX 壓縮以避免防毒軟體誤報
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不顯示命令列視窗
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可選：加入 icon='app.ico'
)
