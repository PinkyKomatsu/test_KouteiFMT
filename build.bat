@echo off
chcp 65001 > nul
rem ============================================================
rem  Excel報告書作成アシスタント  ビルド
rem    1. .venv を作成
rem    2. requirements.txt をインストール
rem    3. PyInstaller で単体 exe（dist\ReportAssistant.exe）を作成
rem  ※ このファイルは UTF-8・CRLF 改行で保存すること
rem ============================================================
setlocal
cd /d "%~dp0"

set PY=py -3
%PY% --version > nul 2>&1 || set PY=python

if not exist .venv (
    echo [1/3] 仮想環境を作成しています...
    %PY% -m venv .venv || goto :error
)
call .venv\Scripts\activate.bat || goto :error

echo [2/3] ライブラリをインストールしています...
python -m pip install --upgrade pip > nul
python -m pip install -r requirements.txt || goto :error

echo [3/3] exe を作成しています...
pyinstaller --noconfirm --clean --onefile --windowed ^
    --name ReportAssistant ^
    --hidden-import lxml._elementpath ^
    --exclude-module PIL ^
    main.py || goto :error

if not exist dist\data mkdir dist\data
echo.
echo 完了しました: dist\ReportAssistant.exe
echo 配布するときは dist フォルダ（ReportAssistant.exe と data フォルダ）をまとめてコピーしてください。
exit /b 0

:error
echo.
echo ビルドに失敗しました。上のメッセージを確認してください。
exit /b 1
