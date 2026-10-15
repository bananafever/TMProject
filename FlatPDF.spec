# -*- mode: python ; coding: utf-8 -*-
# flat_pdf.py 단독 실행 파일 빌드: python -m PyInstaller FlatPDF.spec --clean --noconfirm

# 쓰지 않는 무거운 패키지 (다른 프로젝트용으로 설치돼 있어 같이 끌려오는 것들)
EXCLUDES = [
    'tkinter', 'numpy', 'lxml', 'PIL.AvifImagePlugin', 'PIL._avif',
    'PySide6.QtNetwork', 'PySide6.QtQml', 'PySide6.QtQuick', 'PySide6.QtPdf', 'PySide6.QtOpenGL',
]
# PySide6 훅이 플러그인과 함께 넣는 불필요한 Qt DLL (위젯 앱에는 필요 없음)
DROP_BINARIES = ('opengl32sw', 'Qt6Quick', 'Qt6Qml', 'Qt6Pdf', 'Qt6Network', 'Qt6OpenGL', 'Qt6VirtualKeyboard', 'qpdf')

a = Analysis(
    ['flat_pdf.py'],
    pathex=[],
    binaries=[],
    datas=[('app_icon.ico', '.'), ('assets/fonts', 'assets/fonts')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=0,
)
a.binaries = [b for b in a.binaries if not any(k.lower() in b[0].lower() for k in DROP_BINARIES)]
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FlatPDF',
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
    icon=['app_icon.ico'],
)
