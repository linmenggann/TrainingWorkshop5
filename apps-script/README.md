# 📊 Apps Script 報名資料串接說明

將 `index.html` 報名表單的資料寫入 Google Sheets。

- 試算表：<https://docs.google.com/spreadsheets/d/1idrIOTSCd1Ev2CHZ0I306N-LyqJViEpxXyQAML-aw6I/edit>
- 分頁名稱：`工作坊報名資料`

## 📝 Google Sheets 表頭（第 1 列）

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| 報名時間 | 姓名 | 機構名 | 職稱 | 負責的職類 | 是否擔任教學訓練計畫主持人 | 參與方式 | Email | 聯繫電話 |

> 首次部署後，於 Apps Script 編輯器內手動執行 `setupSheet` 函式即可自動建立分頁並寫入上述表頭、樣式與凍結首列。

## 🚀 部署步驟

1. 開啟試算表 → 選單 **擴充功能 (Extensions) → Apps Script**。
2. 將 [`Code.gs`](./Code.gs) 內容整段貼入，並儲存專案。
3. 於編輯器執行一次 `setupSheet`（第一次需授權 Google 帳號權限）。
4. 點右上角 **部署 Deploy → 新增部署作業**：
   - 類型：**網頁應用程式 (Web app)**
   - 執行身分 Execute as：**我 (Me)**
   - 存取權限 Who has access：**任何人 (Anyone)**
5. 部署成功後複製 **網頁應用程式網址**，格式類似：
   ```
   https://script.google.com/macros/s/AKfycbx.../exec
   ```
6. 開啟 `index.html`，將檔案頂端 `<script>` 內的 `APPS_SCRIPT_URL` 換成上一步取得的網址。

## 🔁 更新程式碼

修改 `Code.gs` 後，請至「**管理部署作業** → 編輯 ✏️ → 版本選 **新版本** → 部署」，舊的 URL 才會指向最新程式碼。

## 🎟 報名人數上限

`Code.gs` 頂端的 `MAX_REGISTRATIONS` 常數控制報名人數上限（預設 **4 名**）。

- `doPost` / `action=register`：寫入前檢查人數，超過即回傳 `status: "full"` 並拒絕寫入。
- `action=capacity`：輕量查詢端點，回傳 `{ total, max, remaining, full }`，供 `index.html` 載入時檢查額滿狀態。
- `action=data`：回傳資料同時附上 `capacity` 物件，dashboard 顯示「剩 X 席／額滿」。

調整上限後請記得 **重新部署新版本** 才會生效。

## 📊 儀表板（dashboard.html）

`dashboard.html` 會透過 `?action=data&callback=xxx` 以 JSONP 方式讀取試算表資料並顯示：

- KPI：總報名數、實體／線上人數、現任主持人數
- 圖表：職類分佈（doughnut）、參與方式（bar）、是否擔任主持人（pie）、每日報名趨勢（line）
- 最新 20 筆報名名單

> Email 與電話在後端已做隱碼處理（例：`ab***@gmail.com`、`091****678`），降低外洩風險。

## 🧪 本機測試

1. 在瀏覽器直接打開 Apps Script Web app URL，應回傳 `{"status":"ok",...}`。
2. 於 `index.html` 填寫表單送出，切回試算表 `工作坊報名資料` 分頁確認新增一列資料。

## 🔐 隱私提醒

Apps Script Web app 若設為 **Anyone**，任何知道 URL 的人都能呼叫。若需限制，可改為驗證 token 或改用表單 triggers；本範例以最簡流程為主。
