@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist "venv\Scripts\activate.bat" (
    echo 尚未安裝環境，請先雙擊 install.bat
    pause
    exit /b 1
)

call venv\Scripts\activate.bat
echo 啟動 Streamlit 面版...
echo 瀏覽器將自動開啟，若未開啟請訪問 http://localhost:8501
echo 按 Ctrl+C 可關閉
echo.
streamlit run app.py
