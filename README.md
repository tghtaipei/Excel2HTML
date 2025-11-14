# EXCEL轉HTML表格系統

## 📋 目錄
- [系統簡介](#系統簡介)
- [功能特色](#功能特色)
- [系統架構](#系統架構)
- [安裝步驟](#安裝步驟)
- [使用說明](#使用說明)
- [進階功能](#進階功能)
- [開發指南](#開發指南)
- [常見問題](#常見問題)

---

## 系統簡介

這是一個基於 Google Apps Script 開發的自動化工具，可將 Excel 檔案（.xlsx / .xls）轉換為互動式 HTML 表格，並自動儲存到 Google Drive。生成的 HTML 支援搜尋、篩選、多分頁切換等功能。

### 適用場景
- 政府機關公開資料發布
- 企業內部資料共享
- 醫療院所名冊管理
- 服務據點資訊展示
- 任何需要將 Excel 資料轉為網頁展示的場合

---

## 功能特色

### ✨ 核心功能

1. **一鍵匯入**
   - 支援 .xlsx 和 .xls 格式
   - 自動解析所有分頁
   - 自動移除空白資料列

2. **智慧標題識別**
   - 上傳後詢問標題列位置
   - 提供前5列預覽協助判斷
   - 自動檢測合併儲存格並提醒

3. **多分頁整合**
   - 單一 Excel 的所有分頁整合為一個 HTML 檔案
   - 分頁切換按鈕快速導覽
   - 節省檔案管理成本

4. **進階搜尋功能**
   - 關鍵字即時搜尋
   - 搜尋結果自動標記高亮
   - 顯示符合筆數統計

5. **智慧篩選系統**
   - **行政區篩選**：自動偵測台北市12個行政區（中正、中山、萬華、信義、大安、文山、內湖、南港、北投、士林、大同）
   - **勾選欄位篩選**：自動識別勾選型欄位（✓、V、○等符號），提供單一勾選框快速篩選
   - 多條件組合篩選
   - 即時更新顯示筆數

6. **自動儲存與記錄**
   - HTML 檔案自動上傳至 Google Drive
   - 產生可分享的公開連結
   - Google Sheets 自動記錄匯入資訊（檔名、筆數、URL、處理時間）

7. **Froala Editor 相容**
   - 無 `<html>`、`<head>`、`<body>` 等外層標籤
   - 所有 CSS class 加上 `excel-` 前綴避免衝突
   - 所有 JavaScript 函數加上 `excel` 前綴避免衝突
   - 可直接插入任何 HTML 編輯器中

### 🎨 介面設計

- **淡藍色系配色**：清新專業的視覺風格
- **響應式設計**：支援各種螢幕尺寸
- **直覺操作**：拖拉式按鈕、即時反饋
- **友善提示**：粉紅色輸入框提醒使用者操作

---

## 系統架構

### 檔案結構

```
📦 Google Apps Script 專案
├── 📄 Code.gs                    # 主控制器
│   ├── showSidebar()            # 開啟側邊欄
│   ├── onOpen()                 # 建立選單
│   ├── handleImport()           # 第一階段：上傳解析
│   ├── handleImportWithHeaders() # 第二階段：處理資料
│   ├── handleClear()            # 清除資料
│   ├── parseExcelFile()         # Excel 解析
│   ├── getSheetPreview()        # 分頁預覽
│   └── checkMergedCells()       # 檢查合併儲存格
│
├── 📄 ParserFactory.gs           # 解析器工廠
│   ├── ParserFactory            # 解析器管理
│   ├── BaseParser               # 基礎解析器類別
│   ├── DefaultParser            # 預設解析器
│   ├── DisciplineParser         # 規則格式解析器
│   └── ContractUnitParser       # 特約單位解析器
│
├── 📄 HTMLGenerator.gs           # HTML 生成器
│   ├── generateMultiSheet()     # 生成多分頁 HTML
│   ├── generateSheetContent()   # 生成單一分頁內容
│   ├── detectDistricts()        # 偵測行政區
│   ├── detectCheckboxColumns()  # 偵測勾選欄位
│   ├── getStyles()              # CSS 樣式
│   ├── getScripts()             # JavaScript 功能
│   └── escapeHtml()             # HTML 跳脫
│
├── 📄 DriveManager.gs            # Drive 管理器
│   ├── saveHTML()               # 儲存 HTML 到 Drive
│   └── getOrCreateFolder()      # 取得或建立資料夾
│
├── 📄 SheetsLogger.gs            # Sheets 記錄器
│   └── logResults()             # 記錄匯入結果
│
└── 📄 Sidebar.html               # 使用者介面
    ├── 檔案選擇區
    ├── 標題列輸入介面
    └── 狀態顯示區
```

### 資料流程

```
┌─────────────────┐
│ 使用者上傳 EXCEL │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  parseExcelFile  │ ──► 轉換為 Google Sheets 暫存
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  顯示預覽介面    │ ──► 詢問標題列位置
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  ParserFactory   │ ──► 選擇適當解析器
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ HTMLGenerator    │ ──► 生成互動式 HTML
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  DriveManager    │ ──► 儲存到 Google Drive
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  SheetsLogger    │ ──► 記錄到試算表
└─────────────────┘
```

---

## 安裝步驟

### 前置需求
- Google 帳號
- Google Sheets 試算表
- 基本的 Apps Script 操作知識

### 步驟 1：建立 Google Sheets

1. 開啟 [Google Sheets](https://sheets.google.com)
2. 建立新的試算表
3. 命名為「EXCEL轉HTML系統」（或任意名稱）

### 步驟 2：開啟 Apps Script 編輯器

1. 點擊試算表選單：**擴充功能** > **Apps Script**
2. 刪除預設的 `Code.gs` 內容

### 步驟 3：建立專案檔案

依序建立以下檔案並貼上對應程式碼：

#### 3.1 Code.gs
```javascript
// 從本專案的 Code.gs 複製完整程式碼
```

#### 3.2 ParserFactory.gs
```javascript
// 從本專案的 ParserFactory.gs 複製完整程式碼
```

#### 3.3 HTMLGenerator.gs
```javascript
// 從本專案的 HTMLGenerator.gs 複製完整程式碼
```

#### 3.4 DriveManager.gs
```javascript
// 從本專案的 DriveManager.gs 複製完整程式碼
```

#### 3.5 SheetsLogger.gs
```javascript
// 從本專案的 SheetsLogger.gs 複製完整程式碼
```

#### 3.6 Sidebar.html
```html
<!-- 從本專案的 Sidebar.html 複製完整程式碼 -->
```

### 步驟 4：啟用 Drive API

1. 在 Apps Script 編輯器左側，點擊 **「服務」**（Services）圖示（齒輪旁邊的 + 號）
2. 找到 **「Drive API」**
3. 選擇版本 **v2** 或 **v3**
4. 點擊 **「新增」**

### 步驟 5：授權與部署

1. 點擊 **「儲存」** 圖示（磁碟片）
2. 點擊 **「執行」** > 選擇 `onOpen` 函數
3. 第一次執行會要求授權：
   - 點擊 **「審查權限」**
   - 選擇您的 Google 帳號
   - 點擊 **「進階」** > **「前往 EXCEL轉HTML系統（不安全）」**
   - 點擊 **「允許」**

### 步驟 6：測試安裝

1. 回到 Google Sheets
2. 重新整理頁面（F5）
3. 頂部選單應該會出現 **「資料匯入」** 選項
4. 點擊 **「資料匯入」** > **「開啟匯入介面」**
5. 右側應該會顯示側邊欄介面

---

## 使用說明

### 基本操作流程

#### Step 1：開啟操作介面

1. 開啟 Google Sheets 試算表
2. 點擊選單：**資料匯入** > **開啟匯入介面**
3. 右側會出現 600px 寬的側邊欄

#### Step 2：選擇 Excel 檔案

1. 點擊 **「📊 匯入EXCEL表格」** 按鈕
2. 選擇本機的 .xlsx 或 .xls 檔案
3. 等待上傳（會顯示「上傳並解析檔案中...」）

#### Step 3：指定標題列位置

上傳完成後，系統會顯示每個分頁的預覽：

1. **查看預覽表格**
   - 每個分頁顯示前 5 列資料
   - 列號標示在左側
   - 最多顯示前 10 欄

2. **輸入標題列位置**
   - 在粉紅色輸入框中輸入數字（例如：1、2、3）
   - 標題列通常是欄位名稱所在的列
   - 每個分頁可以設定不同的標題列位置

3. **注意事項**
   - 如果系統偵測到標題列有合併儲存格，會提示錯誤
   - 請回到 Excel 移除合併儲存格後重新上傳

4. **確認送出**
   - 點擊 **「✓ 確認並繼續匯入」** 按鈕
   - 等待處理（會顯示「處理資料中...」）

#### Step 4：取得結果

處理完成後：

1. **成功訊息**
   - 顯示綠色提示：「成功匯入 X 個分頁，共 X 筆資料」

2. **查看記錄**
   - 回到 Google Sheets
   - 試算表中會新增一列記錄，包含：
     - 匯入時間
     - 檔案名稱
     - 分頁資訊
     - 資料筆數
     - HTML 檔案 URL
     - 處理時間

3. **開啟 HTML 檔案**
   - 點擊記錄中的 URL 連結
   - 在新視窗開啟 HTML 檔案

### HTML 表格操作

生成的 HTML 檔案提供以下功能：

#### 1. 分頁切換
- 頂部顯示所有分頁名稱按鈕
- 點擊按鈕切換不同分頁
- 當前分頁會以深藍色高亮顯示

#### 2. 關鍵字搜尋
- 在搜尋框輸入關鍵字
- 點擊「搜尋」按鈕或按 Enter
- 符合的內容會以黃色標記
- 顯示找到的項目數量
- 點擊「清除」取消搜尋

#### 3. 行政區篩選
如果資料包含台北市行政區資訊：
- 在「行政區：」區塊勾選要篩選的區域
- 可複選多個區域
- 即時顯示符合的資料

#### 4. 勾選欄位篩選
如果資料包含勾選欄位（✓、V、○等）：
- 在「其他過濾條件：」區塊勾選欄位名稱
- 勾選後只顯示該欄位「已勾選」的資料
- 可組合多個條件篩選

#### 5. 資料統計
- 頁面底部顯示「總計 X 筆資料」
- 篩選或搜尋後數字會即時更新

### 清除資料

如需清除試算表中的匯入記錄：

1. 點擊側邊欄的 **「🗑️ 清除資料」** 按鈕
2. 確認刪除對話框
3. 所有記錄（第2列以後）會被清除
4. 標題列（第1列）保留

---

## 進階功能

### 自訂解析器

如果您的 Excel 有特殊格式，可以建立自訂解析器：

```javascript
/**
 * 自訂解析器範例
 */
function CustomParser() {
  BaseParser.call(this);
  this.typeName = 'CustomType';
}

CustomParser.prototype = Object.create(BaseParser.prototype);
CustomParser.prototype.constructor = CustomParser;

// 判斷是否使用此解析器
CustomParser.prototype.canHandle = function(sheetName, worksheet) {
  // 例如：分頁名稱包含「特殊」關鍵字
  return sheetName.indexOf('特殊') !== -1;
};

// 自訂解析邏輯
CustomParser.prototype.parse = function(worksheet, sheetName) {
  const data = this.worksheetToArray(worksheet);
  
  // 自訂處理邏輯
  // ...
  
  return {
    title: sheetName,
    headers: [...],
    data: [...]
  };
};
```

在 `ParserFactory.getParser()` 中加入：

```javascript
const customParser = new CustomParser();
if (customParser.canHandle(sheetName, worksheet)) {
  return customParser;
}
```

### 修改配色主題

在 `HTMLGenerator.gs` 的 `getStyles()` 函數中修改 CSS：

```javascript
// 修改主題色
css += '.tab-btn { background: linear-gradient(135deg, #YOUR_COLOR1 0%, #YOUR_COLOR2 100%); }\n';
```

常用色系：
- **淡藍色**：`#4fc3f7` → `#0288d1`
- **綠色**：`#66bb6a` → `#43a047`
- **橘色**：`#ff9800` → `#f57c00`
- **紫色**：`#9c27b0` → `#7b1fa2`

### 新增行政區

在 `HTMLGenerator.gs` 的 `detectDistricts()` 函數中修改：

```javascript
const districts = ['中正', '中山', '萬華', '信義', '大安', '文山', '內湖', '南港', '北投', '士林', '大同', '新增區域'];
```

### 調整勾選符號識別

在 `HTMLGenerator.gs` 的 `isCheckboxChecked()` 函數中修改：

```javascript
const checkedValues = ['v', 'V', '✓', '✔', '√', 'true', 'yes', '是', 'o', 'O', '●', '⊙', '新符號'];
```

### 修改儲存位置

在 `DriveManager.gs` 中修改資料夾名稱：

```javascript
const folder = this.getOrCreateFolder('自訂資料夾名稱');
```

---

## 開發指南

### 開發環境

- **編輯器**：Google Apps Script 編輯器
- **語言**：JavaScript (ES5)
- **測試**：直接在 Google Sheets 中測試

### 程式碼規範

1. **命名規則**
   - 函數：camelCase（例如：`handleImport`）
   - 類別：PascalCase（例如：`BaseParser`）
   - 常數：UPPER_CASE（例如：`MAX_ROWS`）

2. **註解規範**
   ```javascript
   /**
    * 函數說明
    * @param {string} param1 - 參數說明
    * @return {object} 回傳值說明
    */
   ```

3. **錯誤處理**
   - 使用 try-catch 捕捉例外
   - 記錄到 Logger
   - 回傳友善的錯誤訊息

### 除錯技巧

#### 1. 查看執行記錄

```javascript
Logger.log('除錯訊息: ' + variable);
```

在 Apps Script 編輯器：**檢視** > **記錄**

#### 2. 測試單一函數

```javascript
function testParsing() {
  // 測試程式碼
}
```

選擇函數後點擊「執行」

#### 3. 使用中斷點

在 Apps Script 編輯器無法使用中斷點，建議：
- 使用 `Logger.log()` 追蹤變數
- 分段測試小功能
- 使用瀏覽器開發者工具測試 JavaScript

### 效能優化

1. **減少 API 呼叫**
   - 批次處理資料
   - 快取常用資料

2. **優化資料結構**
   - 避免巢狀迴圈
   - 使用適當的資料結構

3. **限制檔案大小**
   - 建議單一 Excel 檔案 < 5MB
   - 單一分頁資料 < 10,000 列

---

## 常見問題

### Q1：上傳後沒有反應？

**可能原因：**
- 檔案太大（超過 10MB）
- 網路連線不穩定
- Excel 檔案損毀

**解決方法：**
1. 檢查檔案大小，嘗試分割大檔案
2. 重新整理頁面後再試
3. 確認 Excel 檔案可正常開啟

### Q2：顯示「合併儲存格」錯誤？

**原因：**
標題列包含合併的儲存格

**解決方法：**
1. 在 Excel 中選取標題列
2. 點擊「常用」>「合併與置中」>「取消合併儲存格」
3. 儲存後重新上傳

### Q3：HTML 表格在 Froala Editor 中顯示異常？

**可能原因：**
- Froala 的 CSS 與生成的 CSS 衝突
- JavaScript 函數名稱衝突

**解決方法：**
系統已使用 `excel-` 前綴避免衝突，如仍有問題：
1. 檢查 Froala 的自訂 CSS
2. 在瀏覽器開發者工具查看衝突的樣式
3. 修改 `HTMLGenerator.gs` 中的 CSS class 名稱

### Q4：行政區篩選沒有出現？

**原因：**
資料中沒有包含台北市行政區關鍵字

**檢查方法：**
資料是否包含以下任一關鍵字：
- 中正、中山、萬華、信義、大安、文山
- 內湖、南港、北投、士林、大同

如需新增其他行政區，請參考[新增行政區](#新增行政區)

### Q5：勾選欄位篩選無法正常運作？

**可能原因：**
- 勾選符號不在識別清單中
- 資料格式不一致

**解決方法：**
1. 確認使用的勾選符號：✓、V、v、○、●
2. 統一資料格式
3. 或參考[調整勾選符號識別](#調整勾選符號識別)新增符號

### Q6：生成的 HTML 檔案連結無法開啟？

**可能原因：**
- 檔案權限設定問題
- Drive 儲存空間已滿

**解決方法：**
1. 檢查 Google Drive 儲存空間
2. 確認「匯入資料」資料夾存在
3. 手動設定檔案分享權限

### Q7：處理時間過長？

**原因：**
- 資料量過大
- 複雜的格式解析

**優化建議：**
1. 分割大型 Excel 檔案
2. 移除不必要的分頁
3. 簡化 Excel 格式（移除公式、合併儲存格等）

### Q8：如何備份系統？

**方法：**
1. **備份程式碼**
   - Apps Script 編輯器 > 檔案 > 建立版本
   - 或複製整個試算表

2. **匯出程式碼**
   - 逐一複製各 .gs 和 .html 檔案內容
   - 儲存為本機檔案

3. **定期備份**
   - 建議每次重大修改後建立版本

### Q9：能否批次處理多個 Excel 檔案？

**目前限制：**
系統設計為一次處理一個 Excel 檔案

**替代方案：**
1. 逐一上傳各檔案
2. 或修改程式碼實現批次處理功能

### Q10：生成的 HTML 可以離線使用嗎？

**答案：**
可以，HTML 檔案包含所有必要的 CSS 和 JavaScript

**使用方法：**
1. 從 Drive 下載 HTML 檔案
2. 用瀏覽器開啟即可
3. 所有功能（搜尋、篩選）都可正常運作

---

## 版本資訊

### v1.0.0 (2024-11-14)

**功能：**
- ✅ Excel 檔案匯入與解析
- ✅ 智慧標題列識別
- ✅ 多分頁整合
- ✅ 關鍵字搜尋
- ✅ 行政區篩選
- ✅ 勾選欄位篩選
- ✅ Google Drive 自動儲存
- ✅ Froala Editor 相容
- ✅ 淡藍色主題設計

---

## 授權聲明

本專案採用 MIT License

```
MIT License

Copyright (c) 2024

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 聯絡資訊

如有問題或建議，歡迎透過以下方式聯繫：

- **Email**：tghtaipei@gmail.com

---

## 致謝

感謝以下專案與資源：

- Google Apps Script 文件
- Google Drive API
- 所有使用者的回饋與建議

---

**最後更新：2024-11-14**
