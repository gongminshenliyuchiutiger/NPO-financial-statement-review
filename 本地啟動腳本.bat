@echo off
setlocal
echo ==========================================
echo   NPO 財務報表自動檢核系統 - 本地啟動腳本
echo ==========================================
echo.
echo [1] 正在檢查 Python 環境...
where python >nul 2>nul
if %ERRORLEVEL% == 0 (
    echo [OK] 找到 Python，正在啟動臨時伺服器...
    echo 請在瀏覽器中查看: http://localhost:8000
    echo (啟動後請勿關閉此視窗)
    start "" "http://localhost:8000"
    python -m http.server 8000
    goto end
)

echo [2] 正在檢查 Node.js (npm) 環境...
where npm >nul 2>nul
if %ERRORLEVEL% == 0 (
    echo [OK] 找到 npm，正在使用 npx 啟動臨時伺服器...
    echo 請在瀏覽器中查看: http://localhost:3000
    echo (啟動後請勿關閉此視窗)
    npx -y serve -l 3000 .
    goto end
)

echo [!] 錯誤：電腦中找不到 Python 或 Node.js 執行環境。
echo.
echo 因為瀏覽器的安全機制 (CORS)，直接點擊 index.html 無法執行 AI 核心程式。
echo.
echo 建議解決方案：
echo 1. 使用 VS Code 並安裝 "Live Server" 擴充套件，點擊右下角 "Go Live" 執行。
echo 2. 將此專案上傳至 GitHub 並啟用 GitHub Pages (推薦)。
echo 3. 安裝 Python (https://www.python.org/) 或是 Node.js 後，重新點擊此腳本。
echo.
pause

:end
