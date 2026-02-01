# 部署指南 (Deployment Guide)

由於這是一個 Python Flask 後端應用程式，無法使用 GitHub Pages（僅支援靜態網頁）。推薦使用 **Render** 進行免費部署。

## 部署步驟 (使用 Render)

### 1. 準備工作 (已完成)
我們已經為專案新增了以下必要檔案：
- `requirements.txt`: 包含 `gunicorn` (生產環境伺服器)。
- `Procfile`: 告訴 Render 如何啟動程式 (`web: gunicorn app:app`)。
- `.gitignore`: 避免上傳垃圾檔案。

### 2. 上傳到 GitHub
請確保你已經將這些檔案 Commit 並 Push 到你的 GitHub Repository。

### 3. 在 Render 建立服務
1. 註冊並登入 [Render Dashboard](https://dashboard.render.com/)。
2. 點擊 **"New +"** 按鈕，選擇 **"Web Service"**。
3. 選擇 **"Build and deploy from a Git repository"**。
4. 連結你的 GitHub 帳號，並選擇這個 Repository (NPO-financial-statement-review)。

### 4. 設定參數
在設定頁面確認以下資訊：
- **Name**: 給你的服務取個名字 (例如 `npo-review-app`)。
- **Region**: 選擇離你最近的 (例如 Singapore 或 Oregon)。
- **Runtime**: 選擇 **Python 3**。
- **Build Command**: `pip install -r requirements.txt` (預設應該就是這個)。
- **Start Command**: `gunicorn app:app` (Render 會自動讀取 Procfile，所以這裡應該會自動填入，如果沒有請手動輸入)。
- **Environment Variables** (環境變數):
    - 點擊 "Add Environment Variable"。
    - Key: `GOOGLE_API_KEY`
    - Value: (填入你的 Google Gemini API Key)
    - *如果不設定，程式中若未提供 API Key 將無法使用 AI 功能。*

### 5. 部署
點擊 **"Create Web Service"**。Render 會開始自動下載套件並部署。等待約幾分鐘，出現 "Live" 狀態後，你就會獲得一個網址 (例如 `https://npo-review-app.onrender.com`)，這就是你的線上系統了！

---

## 其他選項
如果不使用 Render，也可以考慮：
- **PythonAnywhere**: 專門針對 Python 的託管服務，有免費方案，但設定方式略有不同。
- **Heroku**: 老牌服務，但目前已無免費方案。
- **GCP Cloud Run**: Google 的雲端服務，適合容器化部署 (需要 Dockerfile)。
