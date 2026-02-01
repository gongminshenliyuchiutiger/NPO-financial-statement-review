# NPO Financial Statement Review (GitHub Pages Version)

這是一個協助非營利組織（NPO）檢核財務報表的自動化工具。已經從 Python 版本完全改寫為 **純前端 JavaScript 版本**，因此可以直接在瀏覽器中執行，無需安裝任何軟體，並可透過 GitHub Pages 免費部署。

## 功能特色
- **完全離線執行**: 所有資料處理都在您的瀏覽器中完成，Excel 檔案不會上傳到任何伺服器，確保資安。
- **Gemini AI 整合**: 支援輸入您的 Gemini API Key 來辨識 PDF 或圖片格式的財報。
- **資深會計師級檢核**: 內建嚴謹審核邏輯（如餘絀上限、負債比、基金動支偵測）。
- **動態範本生成**: 可直接下載最新的標準 4 表範本（收支、資負、基金、財產目錄），避免格式錯誤。
- **互動式小幫手**: 介面包含互動式吉祥物，提升使用體驗。

## 如何使用

### 本機執行
1. 下載此專案。
2. 直接雙擊打開 `index.html` 即可使用。

### 線上部署 (GitHub Pages)
1. 將此專案上傳到 GitHub Repository。
2. 到 Repository 的 **Settings** -> **Pages**。
3. 在 **Build and deployment** 下的 **Branch** 選擇 `main` (或 `master`)，資料夾選擇 `/ (root)`。
4. 點擊 **Save**。
5. 等待幾分鐘後，GitHub 會提供一個網址 (例如 `https://yourname.github.io/repo-name/`)，即可線上使用。

## 檔案結構
- `index.html`: 主程式頁面。
- `js/`: 核心邏輯 (包含標準化、檢核規則、Gemini 串接)。
- `css/`: 樣式表。
- `legacy_python/`: (封存) 舊版的 Python 程式碼，僅供參考。

## 注意事項
- **API Key**: 若要使用圖片/PDF 辨識功能，需要自備 Google Gemini API Key。由於是純前端應用，API Key 不會被儲存在伺服器，但每次重新整理頁面都需要重新輸入。
