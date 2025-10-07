# 同人場行事曆＋攤位收藏 LINE Bot（Google Apps Script）

此專案提供一個可直接部署於 Google Apps Script（V8）的 LINE Bot，支援同人活動查詢、攤位搜尋、收藏與每日提醒，並附有簡易後台統計頁面。

## 專案結構

- `app.gs`：主程式（Webhook、資料讀寫、指令解析、排程任務、工具函式）。
- `admin.html`：後台頁面版型（顯示統計數據與試算表連結）。
- `admin_styles.html`：後台頁面樣式。
- `README.md`：部署與操作指引。

## 主要功能與對應函式

| 功能 | 函式 |
| --- | --- |
| 建立試算表結構 | `initSheets()` |
| 匯入示範資料 | `seedSample()` |
| LINE Webhook 入口 | `doPost(e)` |
| 後台與健康檢查入口 | `doGet(e)` |
| LINE 事件處理 | `handleLineEvent(event)`、`onTextMessage(event)` |
| 每日提醒排程 | `dailyTomorrowFavorites()` |
| 後台渲染 | `renderAdminDashboard()` |

## 部署步驟

1. **建立試算表與 Apps Script 專案**
   - 建立一個新的 Google 試算表，開啟後選擇「擴充功能 → Apps Script」。
   - 刪除範例程式碼後，將 `app.gs` 內容貼入「程式碼.gs」。
   - 新增兩個 HTML 檔案：`admin.html`、`admin_styles.html`，分別貼上對應內容。

2. **初始化資料表**
   - 在 Apps Script 編輯器中，於左側的函式選單選擇 `initSheets`，點擊「執行」，授權後會建立 `Users`、`Events`、`Booths`、`Favorites`、`Config` 分頁與欄位。

3. **填入 Config**
   - 回到試算表的 `Config` 分頁，新增以下 KEY/VALUE：
     - `CHANNEL_TOKEN`：LINE Messaging API Channel access token。
     - `CHANNEL_SECRET`：LINE Channel secret（僅存放於表中，不在程式中輸出）。
     - `BASE_URL`：部署後的 Web App URL（例如：https://script.google.com/macros/s/XXXX/exec）。
     - `ADMIN_SECRET`：自訂後台密碼，用於 `GET BASE_URL?a=admin&secret=...`。

4. **匯入示範資料（可選）**
   - 在 Apps Script 編輯器中選擇 `seedSample` 函式並執行，會建立兩場示範活動與 5 個攤位資料。

5. **部署 Web App**
   - 於 Apps Script 點擊「部署 → 新部署」。
   - 選擇「Web 應用程式」，輸入描述後設定：
     - **執行身份**：自己（擁有者）。
     - **存取權限**：任何擁有連結的人。
   - 部署後取得 URL，更新至 `Config` 表的 `BASE_URL`。

6. **設定 LINE Webhook**
   - 前往 LINE Developers Console → Messaging API → Webhook 設定。
   - 將 Webhook URL 設為：`BASE_URL?a=callback`。
   - 確認啟用 Webhook，並於「回覆設定」保持「使用 webhook」。

7. **建立每日觸發器**
   - 在 Apps Script 編輯器中，於左側「觸發條件」新增觸發器：
     - 函式：`dailyTomorrowFavorites`
     - 事件來源：時間驅動
     - 類型：日曆排程 → 每天 → 時間：上午 9:00（系統會以 Asia/Taipei 執行）。

## 操作說明

- **Webhook 測試**：將 LINE Bot 加入好友後，輸入「場次」即可收到近期活動列表。
- **後台**：瀏覽 `GET BASE_URL?a=admin&secret=你的密碼` 可查看統計與開啟試算表。
- **健康檢查**：`GET BASE_URL?a=ok` 會回傳 `OK`。

## 測試案例建議

1. 首次互動：加入好友後輸入「場次」，應回傳 5 筆內近期活動。
2. 設定上下文：輸入「攤位 FF」，應列出 FF 場前 10 攤並設定 lastEvent。
3. 收藏：輸入「收藏 A12」，Favorites 表會新增一筆並回覆成功訊息。
4. 提醒：輸入「提醒 A12 提前=15」，應更新提醒分鐘並回覆成功。
5. 我的收藏：輸入「我的收藏」，列出收藏清單及所屬場次。
6. 搜攤：輸入「搜攤 Blue」，顯示跨場次搜尋結果。
7. 每日推播：在 Apps Script 執行 `dailyTomorrowFavorites()`，針對明日活動推播收藏提醒。
8. 後台：瀏覽 `BASE_URL?a=admin&secret=...`，顯示四項統計與試算表連結。
9. 錯誤處理：輸入「收藏 Z99」時，應提示找不到並提供指引。

## 常見問題

- **沒有資料**：請確認已執行 `initSheets()` 與 `seedSample()`，或手動填入 Events、Booths。
- **推播失敗**：確保 `CHANNEL_TOKEN` 正確且未過期，必要時重新部署或刷新 Token。
- **後台顯示 Auth Error**：確認網址參數 `secret` 與 `Config` 表的 `ADMIN_SECRET` 相符。

