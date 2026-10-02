# App Store 送審準備

## 發行流程

1. 先依 README 完成兩個 target 的 Team 與自動簽署。
2. 在 Safari 實際驗證：啟用／網站授權、E3 登入、手動同步、背景同步、六個分頁、單檔／批次／ZIP 下載、語言切換、翻譯、摘要。
3. 在 App Store Connect 建立 macOS App，Bundle ID 必須與 Xcode 的 App target 相同。語言可選繁體中文，SKU 使用自己帳號內唯一的代碼。
4. App 與 Extension 的版本號在本分支皆保留上游版本號 2.2.0，Build 為 1；每次重新上傳同一版本時增加兩個 target 的 Build。
5. Xcode 選 Product → Archive，Organizer 選 Validate App，再 Distribute App → App Store Connect 上傳。
6. 填寫商店內容、App Privacy、內容分級及出口合規問項，加入截圖、隱私政策與支援網址，選擇 build 送審。原生通知需在 Safari 實測授權、彈窗及點擊行為後再確認商店功能描述。

## 仍需準備

- 正式公開的隱私政策 URL：依 `USER_NOTICE.md`、實際資料流及使用的第三方服務撰寫。需涵蓋本機資料／API Key、Google Translate、OpenAI，以及使用者如何刪除資料。
- 支援 URL：可使用專案 GitHub Issues，但需確認可公開使用。
- App 與 Safari 功能截圖：避免呈現真實學生姓名、信件、成績、API Key 或 Token。
- App Privacy：依實際資料處理填寫；啟用翻譯和摘要 時會傳送內容給第三方，不能只因沒有自架伺服器便一律宣稱沒有資料傳輸。
- 審查存取方式：E3 需要學校帳號，應提供獲授權的審查用帳號或可供審查的展示方式；不要交出自己的校務帳號。現有專案沒有離線展示模式。
- 最終圖示品質：converter 已由現有 128px 圖示產生 macOS 尺寸，512／1024px 為放大版本。上架前建議換成清晰的原始高解析度圖示。

## 商店描述草稿

NYCU E3 Helper 將陽明交通大學 E3 的作業、課程、公告、信件與教材整合到 Safari 側欄，支援截止倒數、閱讀狀態、課程成績與教材下載。

支援繁體中文及英文。翻譯使用 Google Translate，AI 摘要需自行提供 OpenAI API Key。Safari 版提供側欄提醒與需系統授權的 macOS 原生桌面通知。資料同步需要保持已登入的 E3 分頁開啟，背景工作可能因 Safari 關閉或電腦休眠延後。

本工具為非官方輔助工具，與學校沒有官方隸屬關係。請以 E3 原始資訊為準。

## 審查說明草稿

這是 macOS Safari Web Extension，含一個提供啟用引導的 macOS App。請先執行 App，點 Open Safari Extension Settings，啟用 NYCU E3 Helper 並允許網站存取，再開啟 E3。

需要有效 E3 帳號才能載入課程資料。[提交前填入已授權的審查帳號或展示流程。]

AI 摘要是選用功能，需要使用者的 OpenAI API Key；未設定時不會執行。Safari 使用 nativeMessaging 與 macOS 原生通知橋接，桌面通知需要使用者授權。

以上是草稿；帳號／展示流程等未完成項目填妥後才能送審。
