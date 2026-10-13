# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all, collect_submodules

# 收集这些库的全部内容（包括数据文件、隐藏导入）
datas = []
binaries = []
hiddenimports = []

for pkg in [
    'selenium',
    'pandas',
    'openpyxl',
    'xlrd',
    'docx',
    'pdfplumber',
    'PyPDF2',
    'tkinter',
    'win32com',
    'pythoncom',
    'pywintypes',
]:
    try:
        d, b, h = collect_all(pkg)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception:
        pass

# pdfplumber / pdfminer 的一些子模块需要显式声明
hiddenimports += collect_submodules('pdfminer')
hiddenimports += collect_submodules('docx')
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('pandas')

# 排除不必要的大模块，减小体积
excludes = [
    'matplotlib', 'numpy.f2py', 'scipy', 'IPython', 'jupyter',
    'notebook', 'PIL.ImageQt', 'PyQt5', 'PySide2', 'wx',
    'test', 'unittest', 'pydoc_data',
]

a = Analysis(
    ['hp_test.py'],           # ← 改成你的脚本名
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='题库匹配助手',          # ← 生成的 exe 名字
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                   # 若装了 UPX 可设 True 减小体积
    console=False,               # 不显示黑色控制台窗口
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,                   # ← 有图标的话填 'icon.ico'
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='题库匹配助手',
)