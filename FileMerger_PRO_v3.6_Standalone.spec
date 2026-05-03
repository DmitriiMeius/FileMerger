# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


def collect_tree(source, prefix):
    source = Path(source)
    datas = []

    for path in source.rglob("*"):
        if path.is_file():
            target = Path(prefix) / path.relative_to(source).parent
            datas.append((str(path), str(target)))

    return datas


datas = [("FileMerger_PRO.ico", ".")]
datas += collect_tree(
    Path("Tesseract-OCR"),
    "Tesseract-OCR",
)

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
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
    name='FileMerger_PRO_v3.6_Standalone',
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
    version='version_info.txt',
    icon=['FileMerger_PRO.ico'],
)
