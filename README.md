# 臨床教學中的拋與接-教學溝通的心理安全感與回饋技巧工作坊

奇美醫院教學部｜工作坊視覺化介紹與線上報名網頁。

## 檔案

| 檔案 | 說明 |
|------|------|
| `index.html` | 響應式報名網頁（主視覺、活動資訊、議程、線上報名表單） |
| `dashboard.html` | 報名現況儀表板（直接讀取 Google Sheets GViz 端點） |
| `apps-script/Code.gs` | Google Apps Script 後端，將報名資料寫入 Google Sheets |

## 串接 Google Sheets 報名（重要）

報名資料會寫入下列試算表的 **「工作坊報名資料」** 分頁：
<https://docs.google.com/spreadsheets/d/1oghXn_uNIrESKl-i7WZu8XyAyskizCGgtLnzCVdQNTw/edit>

### 1. 建立試算表表頭

在「工作坊報名資料」分頁的**第一列**依序填入下列欄位（或在 Apps Script 中執行 `setupHeaders` 函式自動建立）：

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| 報名時間 | 院區 | 單位 | 姓名 | 人事號 | 職稱 | 連絡電話/分機 | E-mail |

### 2. 部署 Apps Script

1. 開啟試算表 →「擴充功能 (Extensions)」→「Apps Script」。
2. 將 `apps-script/Code.gs` 內容貼入，存檔。
3. 「部署」→「新增部署作業」→ 類型選「網頁應用程式 (Web app)」。
   - 執行身分：**我**
   - 具有存取權的使用者：**所有人 (Anyone)**
4. 第一次部署需授權，請依指示允許。
5. 複製產生的 `/exec` 網址。

### 3. 設定網頁

打開 `index.html`，找到 `SCRIPT_URL` 常數，將其值換成步驟 2 取得的 `/exec` 網址：

```js
const SCRIPT_URL = 'https://script.google.com/macros/s/XXXXX.../exec';
```

完成後，網頁表單送出的報名資料即會即時寫入 Google Sheets。

> 註：表單以 `mode: 'no-cors'` 送出，瀏覽器無法讀取回應內容，因此網頁在送出完成後即顯示「報名成功」。實際寫入結果請至試算表確認。

## 報名現況儀表板（dashboard.html）

儀表板**不經 Apps Script**，直接讀取 Google Sheets 的 GViz 端點
（`.../gviz/tq?tqx=out:json&sheet=工作坊報名資料`），由 Google Docs 高可用服務提供，較即時呼叫 Apps Script 穩定快速。

**設定（重要）**：需將試算表的共用權限設為 **「知道連結的任何人 → 檢視者 (Anyone with the link → Viewer)」**，儀表板才能讀取資料。

- 資料讀取為唯讀；寫入（報名）仍透過 Apps Script。
- 若日後調整欄位順序，儀表板會依**標題名稱**自動對應，無需改程式。
