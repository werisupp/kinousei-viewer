@echo off
chcp 65001 > nul
echo ====================================
echo  機能性表示食品 ビューワー 起動中
echo ====================================
echo.

REM スクリプトと同じフォルダに移動
cd /d "%~dp0"

REM Python で build_viewer.py を実行
python build_viewer.py

if %ERRORLEVEL% neq 0 (
    echo.
    echo エラーが発生しました。
    echo Python がインストールされているか確認してください。
    echo   https://www.python.org/
    echo.
    echo 初回のみ以下を実行してください:
    echo   pip install pandas openpyxl
    echo.
    pause
)
