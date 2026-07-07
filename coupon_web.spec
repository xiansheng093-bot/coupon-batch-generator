# -*- mode: python ; coding: utf-8 -*-
"""
优惠券批量生成器 - PyInstaller 打包配置
用法: pyinstaller coupon_web.spec
"""

import os
import sys
import importlib

block_cipher = None

# reportlab 的字体数据目录（需要打包进去）
import reportlab as _rl
reportlab_dir = os.path.join(os.path.dirname(_rl.__file__), 'fonts')

a = Analysis(
    ['coupon_web.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # reportlab 内置字体（Helvetica 等 PDF 标准字体需要）
        (reportlab_dir, 'reportlab/fonts'),
    ],
    hiddenimports=[
        # 显式声明所有依赖，避免动态导入遗漏
        'qrcode',
        'qrcode.image.pil',
        'qrcode.image.pure',
        'qrcode.constants',
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFont',
        'reportlab',
        'reportlab.pdfgen',
        'reportlab.pdfgen.canvas',
        'reportlab.lib.utils',
        'reportlab.lib.colors',
        'reportlab.pdfbase',
        'reportlab.pdfbase.pdfmetrics',
        'reportlab.lib.pagesizes',
        'openpyxl',
        'openpyxl.workbook',
        'openpyxl.reader.excel',
        'xlrd',
        # HTTP 服务器标准库
        'http.server',
        'socketserver',
        'webbrowser',
        'csv',
        'json',
        'io',
        'threading',
        'tempfile',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 减少包体积，排除不需要的大型库
        'matplotlib',
        'numpy',
        'scipy',
        'tkinter',
        'unittest',
        'test',
        'setuptools',
        'pip',
        'asyncio',
        'xml.etree',
        'pydoc',
        'doctest',
        'distutils',
    ],
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
    name='优惠券批量生成器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,       # 保持控制台窗口，显示服务地址和 Ctrl+C 提示
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='优惠券批量生成器',
)
