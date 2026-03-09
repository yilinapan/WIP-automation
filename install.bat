@echo off
chcp 65001 >nul
echo ========================================
echo   WIP 轉貼紙 - 環境安裝
echo ========================================
echo.

cd /d "%~dp0"

if not exist "venv" (
    echo [1/2] 建立虛擬環境...
    python -m venv venv
    if errorlevel 1 (
        echo 錯誤：請確認已安裝 Python，並將 Python 加入系統 PATH
        pause
        exit /b 1
    )
    echo 虛擬環境建立完成。
) else (
    echo 虛擬環境已存在，略過建立。
)

echo.
echo [2/2] 安裝套件...
call venv\Scripts\activate.bat
pip install -r requirements.txt
if errorlevel 1 (
    echo 錯誤：套件安裝失敗
    pause
    exit /b 1
)

echo.
echo ========================================
echo   安裝完成！
echo   請雙擊 run.bat 啟動 Streamlit 面版
echo ========================================
pause
