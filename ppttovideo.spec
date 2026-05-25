# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置
支持 macOS (.app) 和 Windows (.exe)
"""
import sys
import os
from PyInstaller.utils.hooks import collect_data_files, copy_metadata

block_cipher = None

# 收集关键包的元数据和数据文件
extra_datas = [('templates', 'templates')]
extra_datas += copy_metadata('imageio')
extra_datas += copy_metadata('imageio_ffmpeg')
extra_datas += copy_metadata('moviepy')
extra_datas += copy_metadata('numpy')
extra_datas += copy_metadata('pillow')
extra_datas += copy_metadata('tqdm')
extra_datas += copy_metadata('proglog')
extra_datas += copy_metadata('pymupdf')
extra_datas += copy_metadata('edge_tts')
extra_datas += collect_data_files('imageio_ffmpeg')
extra_datas += collect_data_files('certifi')

# 获取当前目录
BASE_DIR = os.path.dirname(os.path.abspath(SPEC))

a = Analysis(
    ['main.py'],
    pathex=[BASE_DIR],
    binaries=[],
    datas=extra_datas,
    hiddenimports=[
        'app',
        'core',
        'flask',
        'jinja2',
        'werkzeug',
        'edge_tts',
        'moviepy',
        'moviepy.video',
        'moviepy.video.io',
        'moviepy.video.fx',
        'moviepy.audio',
        'moviepy.audio.fx',
        'PIL',
        'PIL.Image',
        'fitz',
        'pymupdf',
        'pptx',
        'webview',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'numpy',
        'imageio',
        'imageio_ffmpeg',
        'proglog',
        'tqdm',
        'certifi',
        'aiohttp',
        'aiosignal',
        'frozenlist',
        'multidict',
        'yarl',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'unittest',
        'test',
        'tests',
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
    name='PPTtoVideo',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # 不显示终端窗口
    icon=None,  # 可替换为自定义图标 icon='icon.icns' (macOS) / 'icon.ico' (Windows)
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PPTtoVideo',
)

# macOS 专用：生成 .app 包
if sys.platform == 'darwin':
    app = BUNDLE(
        coll,
        name='PPTtoVideo.app',
        icon=None,  # 可替换为 'icon.icns'
        bundle_identifier='com.ppttools.ppttovideo',
        info_plist={
            'CFBundleName': 'PPT转视频',
            'CFBundleDisplayName': 'PPT转视频',
            'CFBundleShortVersionString': '1.0.0',
            'CFBundleVersion': '1.0.0',
            'NSHighResolutionCapable': True,
            'LSMinimumSystemVersion': '10.15',
        },
    )
