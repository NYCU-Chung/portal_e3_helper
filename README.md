# NYCU E3 Helper

繁體中文 · [English](README.en.md)

將陽明交通大學 E3 的作業、課程、公告、信件與教材整合到同一個側欄。

[Chrome 線上應用程式商店](https://chromewebstore.google.com/detail/nycu-e3-helper/cmagonljljocpkfojkabhiedjafamoef) · [E3](https://e3p.nycu.edu.tw/) · [更新日誌](CHANGELOG.md)

本指南對應此 repository 的 2.3.0 原始碼；商店版本可能尚未包含最新變更。

## 功能

- 可調整寬度的側欄與可拖曳浮動按鈕；預設淺色外觀，可切換深色或跟隨系統。
- 作業倒數、繳交狀態、自訂作業與編輯／刪除。
- 課程列表、成員統計與異動、成績查詢。
- 公告與信件整合，支援篩選、已讀狀態與完整內容預覽。
- Google Translate 中英雙向翻譯；可選 Gemini 或 OpenAI 產生摘要與附來源連結的今日總覽。
- 教材掃描、個別下載與 ZIP 批次下載。
- 側欄通知、未讀徽章、可設定的桌面通知與截止前提醒。
- 繁體中文／英文介面與背景同步。

## 安裝

可從 [Chrome 線上應用程式商店](https://chromewebstore.google.com/detail/nycu-e3-helper/cmagonljljocpkfojkabhiedjafamoef) 安裝，或在 Chrome、Edge、Brave 載入原始碼。最低 Chrome 版本為 93；建議使用最新穩定版瀏覽器。

1. 下載原始碼，或執行：

   ```bash
   git clone https://github.com/NYCU-Chung/portal_e3_helper.git
   cd portal_e3_helper
   ```

2. 開啟 `chrome://extensions/`、`edge://extensions/` 或 `brave://extensions/`。
3. 啟用「開發人員模式」，選擇「載入未封裝項目」，指定含有 `manifest.json` 的資料夾。
4. 開啟一般網頁，點擊浮動的「E3小助手」按鈕。

更新原始碼後，請重新載入擴充功能並重新整理已開啟的網頁。瀏覽器內部頁面及部分受限制頁面無法顯示側欄。

## 首次使用、語言與外觀

1. 登入 [E3](https://e3p.nycu.edu.tw/)。
2. 開啟側欄，點擊「同步」載入作業與課程。
3. 使用「作業」、「課程」、「下載」、「公告」、「通知」與「說明」六個分頁。窄側欄會將分頁換行，保留完整標籤。

介面初始語言跟隨瀏覽器：中文使用繁體中文，其他語言使用英文。可在「更多（⋯）→ 設定 → 介面語言」選擇語言，再按「儲存設定」。更換語言會重新整理目前頁面，其他已開啟頁面需手動重新整理。

語言偏好儲存在本機。課程、作業名稱與原始公告／信件保留來源文字；新產生的 AI 摘要與總覽使用所選介面語言。既有通知與更新紀錄可能保留建立時的語言。

外觀預設為淺色。在「更多（⋯）→ 設定 → 外觀 → 主題」選擇深色或跟隨系統，儲存後立即套用。深色主題中的公告／信件預覽仍使用淺色背景，保留原文配色的可讀性。

## 日常操作

### 作業

點擊作業卡片前往繳交頁面；按「標記為已繳交」變更狀態，再按「已繳交」可改回待處理。使用新增、編輯與刪除管理作業。

截止時間使用本機時區。未繳交作業剩餘不到 1 小時顯示紅色、不到 24 小時顯示黃色，其他為一般狀態，逾期為灰色。已繳交且逾期的作業會從列表隱藏。刪除 E3 作業後，再次同步可能重新出現。

### 公告與信件

開啟「公告」載入資料，或按「重新載入」。可依類型與已讀狀態篩選，紅點表示未讀。「查看內容」提供完整預覽、已讀切換與原始頁面連結。

預覽中的「中→英」、「英→中」使用 Google Translate，不需要 AI 金鑰；「AI摘要」使用設定的供應商。「顯示原文」可還原內容。

### 今日總覽

先載入公告與信件，再按公告分頁上方的「產生總覽」。使用前須在設定啟用 AI 摘要，選擇 Gemini 或 OpenAI 並儲存對應金鑰與模型。

總覽依本機日期整理今天最多 40 則最新項目，不受篩選條件影響。輸入為標題、課程、寄件者與時間，不會逐篇讀取完整內文；詳細要求與截止時間請開啟來源連結確認。成功結果存於本機快取，重新產生失敗時保留前次結果。

### 課程、成績與教材

同步後，在「課程」選擇課程並開啟統計或成績；載入資料需要有效的 E3 登入狀態。

在「下載」可於 E3 頁面按「掃描此頁」，或用「選擇課程」掃描指定課程。勾選檔案後，選擇個別下載或下載 ZIP。支援文件、影片與壓縮檔；大型檔案建議分批或從原始頁面下載。

### 通知與提醒

「通知」分頁集中顯示作業、評分、公告、成員異動與版本通知；24 小時內到期作業也會列入提醒。

開啟「更多 → 設定 → 設定通知與提醒時間」，或從瀏覽器擴充功能管理頁開啟選項：

- 開啟或關閉桌面通知，以及新公告／信件通知。
- 選擇即時通知或每日彙整，設定每日通知的本機時間。
- 新增多個截止前提醒時間；移除全部時間即可關閉截止提醒。
- 按「儲存設定」套用變更；「測試通知」檢查桌面通知是否可用。

首次公告／信件快照只建立基準。每日模式在指定時間分別送出待處理通知；已繳交、刪除、變更截止時間或已截止的作業提醒會略過。瀏覽器休眠後若已錯過多個門檻，只保留最近的有效提醒。

通知引擎在瀏覽器開啟時每分鐘處理佇列，休眠可能延後送出；新資料仍依賴 E3 同步與登入。請在瀏覽器與作業系統允許通知。關閉或無法使用桌面通知時，側欄提醒仍可使用。詳見[通知設定與截止提醒](docs/features/notification-preferences.md)。

### 同步

作業、課程與成員檢查每 60 分鐘背景同步；公告與信件每 30 分鐘同步。側欄開啟時也會檢查作業／課程資料是否過舊。「同步」手動更新作業與課程，公告與信件使用公告分頁的「重新載入」。

一般網站也可查看本機資料與背景載入 E3 資料。公告／信件定時同步只使用已開啟的 E3 分頁，沒有分頁時跳過。手動載入可能暫時建立背景 E3 分頁，完成或逾時後關閉。掃描目前頁面的教材需先開啟對應 E3 頁面。

## AI 設定與資料處理

在「更多 → 設定」啟用 AI 摘要，選擇 Gemini 或 OpenAI，填入該供應商的 API Key、選擇模型並儲存設定。切換供應商會保留兩組設定。Gemini 啟用時可從 API 載入模型清單，也可重新載入或手動輸入模型 ID；產生摘要前需明確選擇模型。「測試連接」檢查目前供應商。額度、費用與模型存取權依供應商帳戶而定。

課程資料、已讀狀態、設定與 API Key 儲存在本機瀏覽器。翻譯會傳送待翻譯文字到 Google Translate；單篇摘要會傳送內文到所選 Gemini／OpenAI，總覽則傳送今天的標題、課程、寄件者與時間。OpenAI 請求使用 Responses API 並設定 `store: false`。擴充功能使用既有 E3 登入狀態，不要求輸入 E3 密碼。

來源內容可能包含姓名或其他個人資料，使用翻譯／AI 前請確認適合傳送至第三方。第三方的資料處理依各服務政策；可停用 AI，並繼續使用其他功能。詳見[摘要與今日總覽](docs/features/openai-digest.md)。

## 常見問題

- **同步或載入失敗：** 重新登入 E3、確認網路後重試。
- **AI 摘要失敗：** 確認已啟用摘要、儲存金鑰並選擇模型，再測試連接。
- **翻譯失敗：** 確認網路後重試；翻譯不需要 AI 金鑰。
- **更新後仍是舊介面：** 重新載入擴充功能，再重新整理已開啟網頁。
- **ZIP 下載失敗：** 減少檔案數量或改用個別下載。
- **桌面通知未出現：** 確認設定已儲存、瀏覽器與系統通知權限，並使用測試通知。
- **回報問題：** 使用「更多 → 查看日誌」或 [GitHub Issues](https://github.com/NYCU-Chung/portal_e3_helper/issues)；分享日誌前請檢查個人資料。

## 開發與驗證

`content.js` 實作側欄與 E3 互動；`background.js` 處理下載、同步與 API 代理；`i18n.js` 與 `_locales/` 提供介面及擴充功能說明翻譯。`notification-engine.js` 處理桌面通知排程，`notification-settings.html` 為獨立設定頁。

介面翻譯在插入動態資料前套用，保留課程名稱、URL 與原始內容。請將 E3 擷取選擇器／解析規則與介面翻譯分開。下載與內容請求只允許 HTTPS E3 主機，包含最終重新導向驗證；預覽 HTML 會清理危險標籤、事件屬性與 URL。

執行單元測試與語法檢查：

```bash
node --test tests/*.test.cjs
node --check i18n.js
node --check content.js
node --check background.js
node --check notification-engine.js
```

可選的瀏覽器 fixture（需 Playwright 與瀏覽器）：

```bash
node tests/ui-smoke.cjs
node tests/language-ui-smoke.cjs
node tests/translation-html-smoke.cjs
node tests/notification-ui-smoke.cjs
node tests/notification-extension-smoke.cjs
```

可透過 `CHROME_PATH` 指定 Chrome，或 `PLAYWRIGHT_PATH` 指定外部 Playwright 安裝。一般 UI fixtures 使用本機資料，不連接 E3／AI；真實擴充功能通知 fixture 需 Playwright Chromium，並使用隔離 profile 與阻擋外部 DNS。

功能實作說明：[外觀](docs/features/ui-appearance.md)、[AI 摘要](docs/features/openai-digest.md)、[英文介面](docs/features/english-interface.md)、[通知](docs/features/notification-preferences.md)。

歡迎透過 [Issues](https://github.com/NYCU-Chung/portal_e3_helper/issues) 與 Pull Request 貢獻。歷史版本詳見 [CHANGELOG.md](CHANGELOG.md)。

## 授權

[MIT License](LICENSE)
