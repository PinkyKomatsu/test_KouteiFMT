# -*- mode: python ; coding: utf-8 -*-
# PyInstaller の設定（onedir 形式・ウィンドウアプリ）。build.bat から使う。
#   dist\MedAcctReport\MedAcctReport.exe
#   dist\MedAcctReport\models\ja-en\   翻訳モデル（build.bat がコピー）
#   dist\MedAcctReport\data\glossary.csv   初期用語集（build.bat がコピー）
from PyInstaller.utils.hooks import collect_dynamic_libs

block_cipher = None

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=collect_dynamic_libs("ctranslate2"),
    datas=[],
    hiddenimports=["lxml._elementpath", "sentencepiece", "ctranslate2"],
    excludes=[
        # 実行時に使わないもの（モデル変換用・大きなライブラリ）
        "torch", "transformers", "tokenizers", "huggingface_hub", "PIL", "tkinter",
        "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtQml", "PySide6.QtQuick",
        "PySide6.Qt3DCore", "PySide6.QtMultimedia", "PySide6.QtCharts", "PySide6.QtDataVisualization",
        "PySide6.QtPdf", "PySide6.QtNetwork",
    ],
    noarchive=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="MedAcctReport",
    console=False,
    disable_windowed_traceback=False,
    upx=False,
)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="MedAcctReport")
