# NYCU E3 Helper - 安裝說明

## 快速安裝（Chrome / Edge / Brave）

### 步驟 1：下載專案

**方法 A：使用 Git**
```bash
git clone https://github.com/YOUR_USERNAME/e3_helper.git
cd e3_helper
```

**方法 B：直接下載**
1. 下載專案壓縮檔
2. 解壓縮到任意位置

### 步驟 2：安裝到瀏覽器

#### Chrome

1. 開啟 Chrome 瀏覽器
2. 在網址列輸入 `chrome://extensions/` 並按 Enter
3. 開啟右上角的「開發人員模式」
4. 點擊「載入未封裝項目」
5. 選擇 `e3_helper` 資料夾
6. 完成！

#### Microsoft Edge

1. 開啟 Edge 瀏覽器
2. 在網址列輸入 `edge://extensions/` 並按 Enter
3. 開啟左側的「開發人員模式」
4. 點擊「載入解壓縮的擴充功能」
5. 選擇 `e3_helper` 資料夾
6. 完成！

#### Brave

1. 開啟 Brave 瀏覽器
2. 在網址列輸入 `brave://extensions/` 並按 Enter
3. 開啟右上角的「開發人員模式」
4. 點擊「載入未封裝項目」
5. 選擇 `e3_helper` 資料夾
6. 完成！

### 步驟 3：驗證安裝

1. 訪問 https://e3p.nycu.edu.tw/
2. 登入你的帳號
3. 按 F12 開啟開發者工具
4. 切換到 Console 分頁
5. 應該會看到：
   ```
   NYCU E3 Helper 已載入
   E3 Helper: 已修改 X 個作業連結
   E3 Helper: DOM 監聽已啟動
   ```

6. 找到任何作業連結，將滑鼠移上去
7. 應該會看到連結下方有綠色虛線
8. Tooltip 會顯示「[直接進入繳交區]」

### 步驟 4：測試功能

1. 點擊任一作業連結
2. 應該會直接進入「新增繳交內容」頁面
3. 不需要先看作業說明再點「新增繳交內容」

## 更新擴充功能

當有新版本發布時：

### 方法 A：Git 更新

```bash
cd e3_helper
git pull
```

然後在瀏覽器的擴充功能頁面點擊「重新載入」按鈕。

### 方法 B：手動更新

1. 下載新版本
2. 解壓縮覆蓋舊檔案
3. 在瀏覽器的擴充功能頁面點擊「重新載入」按鈕

## 解除安裝

1. 開啟瀏覽器的擴充功能管理頁面
   - Chrome: `chrome://extensions/`
   - Edge: `edge://extensions/`
   - Brave: `brave://extensions/`

2. 找到「NYCU E3 Helper」
3. 點擊「移除」
4. 確認移除

## 疑難排解

### 問題：擴充功能無法載入

**解決方法**：
1. 確認已開啟「開發人員模式」
2. 確認選擇的資料夾是 `e3_helper`，不是外層資料夾
3. 確認資料夾內有 `manifest.json` 檔案

### 問題：沒有看到修改作業連結的訊息

**解決方法**：
1. 重新載入 E3 頁面（Ctrl+R 或 F5）
2. 清除瀏覽器快取
3. 在擴充功能頁面點擊「重新載入」按鈕
4. 確認擴充功能已啟用

### 問題：作業連結沒有被修改

**解決方法**：
1. 檢查 Console 是否有錯誤訊息
2. 確認連結確實是指向 `/mod/assign/view.php`
3. 在 Console 執行 `window.e3HelperRefresh()` 手動刷新
4. 回報 Issue 並附上連結範例

### 問題：點擊後沒有直接進入繳交區

**可能原因**：
1. 該作業可能還未開放繳交
2. 該作業可能已超過截止時間
3. 該作業可能已經繳交過（需要「編輯繳交內容」）

這些情況都是正常的，是 E3 系統的限制，不是擴充功能的問題。

## 需要幫助？

- 查看 [README.md](README.md) 了解更多功能說明
- 在 GitHub 提交 Issue
- 聯絡開發者

## 系統需求

- Chrome 93+ / Edge 93+ / Brave 或其他 Chromium 核心瀏覽器
- NYCU E3 帳號
- 網際網路連線
