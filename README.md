# 臨床教學中的拋與接-教學溝通的心理安全感與回饋技巧工作坊

奇美醫院教學部｜工作坊視覺化介紹與線上報名網頁。

## 檔案

| 檔案 | 說明 |
|------|------|
| `index.html` | 響應式報名網頁（主視覺、活動資訊、議程、線上報名表單） |
| `dashboard.html` | 報名現況儀表板（直接讀取 Google Sheets GViz 端點） |
| `checkin.html` | 當天現場報到頁（輸入人事號即完成報到） |
| `checkin-dashboard.html` | 報到現況儀表板（報到率、各院區報到、已/未報到名單） |
| `apps-script/Code.gs` | Google Apps Script 後端，處理報名與報到寫入 |

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

## 當天現場報到（checkin.html + checkin-dashboard.html）

報到資料寫入試算表的 **「工作坊報到資料」** 分頁（首次報到時由 Apps Script 自動建立，或先在 Apps Script 執行 `setupCheckinSheet`）。表頭：

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| 報到時間 | 人事號 | 院區 | 單位 | 姓名 | 職稱 |

### 報到頁 `checkin.html`

- 報到者只需輸入 **人事號** → 按「報到」。
- 後端會以人事號比對「工作坊報名資料」：
  - **找到且未報到** → 寫入報到資料，回傳姓名並顯示「報到成功」。
  - **已報到** → 顯示「已完成報到」，不重複寫入。
  - **查無報名** → 顯示提醒，不寫入。
- 沿用同一組 Apps Script `/exec` 網址（以 `action:'checkin'` 分流），無需另外部署新網址；但更新 `Code.gs` 後需**重新部署新版本**。

### 報到儀表板 `checkin-dashboard.html`

- 同時讀取「工作坊報名資料」與「工作坊報到資料」兩個 GViz 端點。
- 顯示：報名人數、已報到、報到率、未報到；報到狀況與各院區報到圓餅圖、報到時段分布，以及**已報到／未報到名單**（可搜尋）。
- 同樣需要試算表共用設為「知道連結的任何人 → 檢視者」。
