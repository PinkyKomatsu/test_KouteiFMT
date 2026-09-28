@echo off
chcp 65001 > nul
rem ============================================================
rem  医事会計システム更新 工程報告書作成アプリ  ビルド
rem    1. .venv（Python 3.11）を作成
rem    2. requirements.txt をインストール
rem    3. pytest を実行（失敗したら中止）
rem    4. PyInstaller（onedir・windowed）で dist\MedAcctReport\ を作成
rem    5. models\ と data\glossary.csv を dist\MedAcctReport\ にコピー
rem  ※ このファイルは UTF-8・CRLF 改行で保存すること
rem ============================================================
setlocal
cd /d "%~dp0"

if not exist .venv (
    echo [1/5] 仮想環境（Python 3.11）を作成しています...
    where uv > nul 2>&1
    if not errorlevel 1 (
        uv venv --python 3.11 .venv || goto :error
    ) else (
        py -3.11 -m venv .venv || goto :error
    )
)
set PY=.venv\Scripts\python.exe

echo [2/5] ライブラリをインストールしています...
where uv > nul 2>&1
if not errorlevel 1 (
    uv pip install --python %PY% -r requirements.txt || goto :error
) else (
    %PY% -m pip install --upgrade pip > nul
    %PY% -m pip install -r requirements.txt || goto :error
)

echo [3/5] テストを実行しています...
%PY% -m pytest tests -q -p no:cacheprovider || goto :error

echo [4/5] exe を作成しています...
%PY% -m PyInstaller --noconfirm --clean MedAcctReport.spec || goto :error

echo [5/5] 翻訳モデルと用語集をコピーしています...
set OUT=dist\MedAcctReport
if not exist %OUT%\data mkdir %OUT%\data
copy /y data\glossary.csv %OUT%\data\glossary.csv > nul || goto :error
if exist models\ja-en\model.bin (
    robocopy models %OUT%\models /E /NFL /NDL /NJH /NJS /NP > nul
    if errorlevel 8 goto :error
) else (
    echo   注意: models\ja-en がありません。英訳は簡易訳になります。
    echo         tools\convert_model.py で翻訳モデルを準備してから、もう一度ビルドしてください。
)

rem 作業用フォルダ build\ にも exe ができるが、_internal がなく起動できないため削除する
if exist build rmdir /s /q build

echo.
echo 完了しました: %OUT%\MedAcctReport.exe
echo 配布するときは %OUT% フォルダごとコピーしてください（exe・_internal・models・data）。
exit /b 0

:error
echo.
echo ビルドに失敗しました。上のメッセージを確認してください。
exit /b 1
