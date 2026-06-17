# -*- mode: python ; coding: utf-8 -*-
"""
PV-DOC PyInstaller yapılandırması.

Tek dosyalık (onefile) Windows EXE üretir. GitHub Actions windows-latest
üzerinde çalıştırılır (bkz. .github/workflows/build.yml).

Yerel derleme (Windows'ta):
    pyinstaller PVDOC.spec
Çıktı:
    dist/PV-DOC.exe
"""

block_cipher = None

# python-docx, Document() çağrısında templates/default.docx dosyasını okur.
# PyInstaller bunu otomatik toplamayabilir; açıkça topluyoruz.
from PyInstaller.utils.hooks import collect_data_files
_docx_datas = collect_data_files('docx')


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    # Not: cikti/sablonlar ve kaynaklar/spek_kartlari klasörleri uygulama
    # ilk çalıştığında kod tarafından oluşturulur (bkz. core/yollar.py).
    datas=[
        ('kaynaklar_ikon/ikon.ico', 'kaynaklar_ikon'),
    ] + _docx_datas,
    hiddenimports=['docx', 'lxml._elementpath', 'lxml.etree'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # İhtiyaç duyulmayan ağır PyQt6 modüllerini hariç tutarak boyutu küçült.
    excludes=[
        'PyQt6.QtQml',
        'PyQt6.QtQuick',
        'PyQt6.Qt3DCore',
        'PyQt6.QtBluetooth',
        'PyQt6.QtPositioning',
        'PyQt6.QtMultimedia',
        'PyQt6.QtWebEngineCore',
        'PyQt6.QtWebEngineWidgets',
        'tkinter',
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PV-DOC',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,        # GUI uygulaması: konsol penceresi açma
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='kaynaklar_ikon/ikon.ico',
)
