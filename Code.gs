/**
 * 臺北市政府衛生局長照2.0特約單位匯入系統
 * 主程式檔案 - Code.gs
 * 
 * 功能：處理Excel檔案上傳、資料解析、HTML生成
 */

// ==================== 全域設定 ====================

const CONFIG = {
  SHEET_NAME: '長照特約單位資料',
  META_ROW: 1,
  HTML_ROW: 2,
  MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
  
  // 欄位對應
  COMMON_COLUMNS: {
    序號: 'A',
    機構名稱: 'B',
    服務區別: 'C',
    郵遞區號: 'D',
    機構地址: 'E',
    聯絡電話: 'F',
    聯絡窗口: 'G'
  }
};

// ==================== UI 相關函數 ====================

/**
 * 在選單中加入自訂功能
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📋 長照特約單位系統')
      .addItem('🔄 開啟匯入介面', 'showSidebar')
      .addItem('🗑️ 清除所有資料', 'clearData')
      .addToUi();
  } catch (e) {
    // 在某些執行環境下無法使用UI,忽略錯誤
    Logger.log('無法創建選單: ' + e.toString());
  }
}

/**
 * 顯示側邊欄
 */
function showSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('UI')
      .setTitle('長照特約單位匯入系統')
      .setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    Logger.log('無法顯示側邊欄: ' + e.toString());
    throw new Error('無法顯示側邊欄,請確認在Google Sheets環境中執行');
  }
}

// ==================== 主要處理函數 ====================

/**
 * 處理上傳的Excel檔案
 * @param {Object} fileData - Base64編碼的檔案資料
 * @returns {Object} 處理結果
 */
function processExcelFile(fileData) {
  try {
    const startTime = new Date();
    
    // 驗證檔案
    if (!fileData || !fileData.data) {
      throw new Error('檔案資料無效');
    }
    
    // 解碼Base64
    const bytes = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(bytes, fileData.mimeType, fileData.name);
    
    // 驗證檔案大小
    if (blob.getBytes().length > CONFIG.MAX_FILE_SIZE) {
      throw new Error('檔案大小超過10MB限制');
    }
    
    // 解析Excel檔案
    Logger.log('開始解析Excel檔案: ' + fileData.name);
    const workbookData = parseExcelFile(blob);
    
    if (!workbookData || workbookData.sheets.length === 0) {
      throw new Error('無法解析Excel檔案或檔案為空');
    }
    
    Logger.log('成功解析 ' + workbookData.sheets.length + ' 個分頁');
    
    // 生成整合HTML
    const integratedHTML = generateIntegratedHTML(workbookData.sheets);
    
    // 計算統計資訊
    const endTime = new Date();
    const processingTime = ((endTime - startTime) / 1000).toFixed(2);
    const totalInstitutions = workbookData.sheets.reduce(function(sum, sheet) { return sum + sheet.dataCount; }, 0);
    
    // 寫入Google Sheet
    writeToSheet(workbookData, integratedHTML, processingTime, totalInstitutions);
    
    return {
      success: true,
      message: '匯入成功！',
      details: {
        分頁數: workbookData.sheets.length,
        總機構數: totalInstitutions,
        處理時間: processingTime + '秒',
        分頁列表: workbookData.sheets.map(function(s) { return s.name + ' (' + s.dataCount + '筆)'; })
      }
    };
    
  } catch (error) {
    Logger.log('處理錯誤: ' + error.toString());
    return {
      success: false,
      message: '處理失敗',
      error: error.toString()
    };
  }
}

/**
 * 解析Excel檔案
 * @param {Blob} blob - Excel檔案Blob
 * @returns {Object} 工作簿資料
 */
function parseExcelFile(blob) {
  try {
    // 將Blob轉換為臨時檔案ID
    const tempFile = DriveApp.createFile(blob);
    const fileId = tempFile.getId();
    
    // 轉換為Google Sheets
    const resource = {
      title: 'temp_conversion',
      mimeType: MimeType.GOOGLE_SHEETS
    };
    
    const sheet = Drive.Files.copy(resource, fileId);
    const spreadsheet = SpreadsheetApp.openById(sheet.id);
    
    // 解析所有分頁
    const sheets = spreadsheet.getSheets();
    const parsedSheets = [];
    
    sheets.forEach(function(sheet, index) {
      try {
        Logger.log('處理分頁 ' + (index + 1) + ': ' + sheet.getName());
        const sheetData = parseSheet(sheet);
        if (sheetData) {
          parsedSheets.push(sheetData);
        }
      } catch (e) {
        Logger.log('分頁 ' + sheet.getName() + ' 解析失敗: ' + e.toString());
      }
    });
    
    // 清理臨時檔案
    DriveApp.getFileById(fileId).setTrashed(true);
    DriveApp.getFileById(sheet.id).setTrashed(true);
    
    return {
      sheets: parsedSheets,
      totalSheets: sheets.length
    };
    
  } catch (error) {
    Logger.log('parseExcelFile錯誤: ' + error.toString());
    throw new Error('Excel檔案解析失敗: ' + error.message);
  }
}

/**
 * 解析單個分頁
 * @param {Sheet} sheet - Google Sheets分頁物件
 * @returns {Object} 分頁資料
 */
function parseSheet(sheet) {
  const sheetName = sheet.getName();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 3) {
    Logger.log('分頁 ' + sheetName + ' 資料不足，跳過');
    return null;
  }
  
  // 讀取所有資料
  const dataRange = sheet.getRange(1, 1, lastRow, lastCol);
  const values = dataRange.getValues();
  
  // 識別格式類型
  const formatType = identifySheetFormat(values);
  
  Logger.log('分頁 ' + sheetName + ' 識別為格式: ' + formatType.mode);
  
  // 根據格式類型解析
  let parsedData;
  if (formatType.mode === 'A') {
    parsedData = parseModeA(values, formatType);
  } else if (formatType.mode === 'B') {
    parsedData = parseModeB(values, formatType);
  } else {
    Logger.log('無法識別分頁格式: ' + sheetName);
    return null;
  }
  
  parsedData.sheetName = sheetName;
  parsedData.formatType = formatType.mode;
  
  return parsedData;
}

/**
 * 識別分頁格式類型
 * @param {Array} values - 分頁資料陣列
 * @returns {Object} 格式類型資訊
 */
function identifySheetFormat(values) {
  // 檢查第一列是否包含標題關鍵字
  const row1 = values[0].join('');
  if (!row1.includes('臺北市政府衛生局長照2.0')) {
    return { mode: 'UNKNOWN' };
  }
  
  // 尋找「序號」和「機構名稱」欄位來確定資料起始列
  let headerRow = -1;
  let dataStartRow = -1;
  
  for (let i = 1; i < Math.min(5, values.length); i++) {
    const rowStr = values[i].join('');
    if (rowStr.includes('序號') && rowStr.includes('機構名稱')) {
      headerRow = i;
      dataStartRow = i + 1;
      break;
    }
  }
  
  if (headerRow === -1) {
    return { mode: 'UNKNOWN' };
  }
  
  // 判斷是模式A還是模式B
  // 模式A: headerRow = 1 (第2列), 8欄左右
  // 模式B: headerRow = 2 (第3列), 13-15欄
  
  const numCols = values[headerRow].filter(function(cell) { return cell !== ''; }).length;
  
  if (headerRow === 1 && numCols <= 10) {
    return {
      mode: 'A',
      titleRow: 0,
      headerRow: 1,
      dataStartRow: 2,
      numCols: numCols
    };
  } else if (headerRow === 2 && numCols >= 10) {
    return {
      mode: 'B',
      titleRow: 0,
      headerRow: 2,
      headerRow2: 1,
      dataStartRow: 3,
      numCols: numCols
    };
  } else {
    return { mode: 'UNKNOWN' };
  }
}

/**
 * 寫入資料到Google Sheet
 * @param {Object} workbookData - 工作簿資料
 * @param {String} htmlCode - HTML原始碼
 * @param {String} processingTime - 處理時間
 * @param {Number} totalInstitutions - 總機構數
 */
function writeToSheet(workbookData, htmlCode, processingTime, totalInstitutions) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  // 如果不存在，創建新分頁
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  
  // 清除現有資料
  sheet.clear();
  
  // 設定欄寬
  sheet.setColumnWidth(1, 150); // 處理時間
  sheet.setColumnWidth(2, 100); // 總分頁數
  sheet.setColumnWidth(3, 100); // 成功數
  sheet.setColumnWidth(4, 100); // 失敗數
  sheet.setColumnWidth(5, 100); // 總機構數
  sheet.setColumnWidth(6, 200); // 處理日期
  
  // 寫入第一列：轉換資訊
  const metaData = [
    [
      '處理時間(秒)',
      '總分頁數',
      '成功數',
      '失敗數',
      '總機構數',
      '處理日期'
    ],
    [
      processingTime,
      workbookData.totalSheets,
      workbookData.sheets.length,
      workbookData.totalSheets - workbookData.sheets.length,
      totalInstitutions,
      new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })
    ]
  ];
  
  sheet.getRange(1, 1, 2, 6).setValues(metaData);
  
  // 設定第一列格式
  sheet.getRange(1, 1, 1, 6)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.getRange(2, 1, 1, 6)
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');
  
  // 檢查HTML大小並決定如何儲存
  const htmlLength = htmlCode.length;
  const MAX_CELL_SIZE = 45000; // 保守設定為45000,確保安全邊際
  
  Logger.log('HTML大小: ' + htmlLength + ' 字元');
  
  // 如果HTML過大,先警告
  if (htmlLength > 200000) {
    Logger.log('⚠️ 警告: HTML大小超過20萬字元,可能需要較長處理時間');
  }
  
  if (htmlLength <= MAX_CELL_SIZE) {
    // HTML不大，直接寫入單一儲存格
    sheet.getRange(3, 1)
      .setValue('整合HTML原始碼（可直接複製使用）')
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    sheet.getRange(4, 1)
      .setValue(htmlCode)
      .setWrap(true);
    
    sheet.setColumnWidth(1, 800);
    
  } else {
    // HTML太大,需要分割
    Logger.log('HTML超過儲存格限制,進行分割...');
    
    // 計算需要的儲存格數量
    const numChunks = Math.ceil(htmlLength / MAX_CELL_SIZE);
    Logger.log('分割為 ' + numChunks + ' 個儲存格');
    
    // 檢查是否超過合理範圍
    if (numChunks > 50) {
      Logger.log('⚠️ 警告: 需要超過50個儲存格,建議優化HTML大小');
    }
    
    // 寫入說明
    sheet.getRange(3, 1)
      .setValue('整合HTML原始碼（共' + numChunks + '個儲存格，請從上到下依序複製合併）')
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    // 分割並寫入HTML
    try {
      for (let i = 0; i < numChunks; i++) {
        const start = i * MAX_CELL_SIZE;
        const end = Math.min(start + MAX_CELL_SIZE, htmlLength);
        const chunk = htmlCode.substring(start, end);
        
        // 驗證chunk大小
        if (chunk.length > 50000) {
          throw new Error('分割後的片段(' + chunk.length + ')仍超過50000字元限制');
        }
        
        const row = 4 + i;
        sheet.getRange(row, 1)
          .setValue(chunk)
          .setWrap(false)
          .setBackground('#f8f9fa');
        
        // 每10個chunk記錄一次進度
        if ((i + 1) % 10 === 0 || i === numChunks - 1) {
          Logger.log('進度: ' + (i + 1) + '/' + numChunks + ' (已寫入 ' + end + '/' + htmlLength + ' 字元)');
        }
      }
    } catch (e) {
      Logger.log('❌ 分割寫入失敗: ' + e.toString());
      throw e;
    }
    
    sheet.setColumnWidth(1, 1000);
    
    // 在最後加上合併說明
    const instructionRow = 4 + numChunks;
    sheet.getRange(instructionRow, 1)
      .setValue('📝 使用說明：\n' +
                '1. 從第4列開始，依序複製每個儲存格的內容\n' +
                '2. 全部貼到同一個文字檔案中（記事本或VSCode）\n' +
                '3. 確保沒有遺漏任何部分\n' +
                '4. 儲存為 .html 檔案（編碼選UTF-8）\n' +
                '5. 用瀏覽器開啟即可使用\n\n' +
                '💡 提示：可以全選第4-' + (3 + numChunks) + '列，一次複製所有內容\n' +
                '⚠️ 注意：共' + numChunks + '個片段，總大小約' + (htmlLength/1024).toFixed(1) + 'KB')
      .setBackground('#fff3cd')
      .setWrap(true)
      .setVerticalAlignment('top');
  }
  
  // 凍結前三列
  sheet.setFrozenRows(3);
  
  Logger.log('資料已成功寫入Sheet');
}

/**
 * 清除所有資料
 */
function clearData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (sheet) {
    try {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        '確認清除',
        '確定要清除所有資料嗎?此操作無法復原。',
        ui.ButtonSet.YES_NO
      );
      
      if (response == ui.Button.YES) {
        sheet.clear();
        ui.alert('資料已清除');
      }
    } catch (e) {
      // 無UI環境,直接清除
      Logger.log('無UI環境,直接清除資料');
      sheet.clear();
      Logger.log('資料已清除');
    }
  } else {
    try {
      SpreadsheetApp.getUi().alert('找不到資料分頁');
    } catch (e) {
      Logger.log('找不到資料分頁');
    }
  }
}

/**
 * 取得已匯入的HTML（供外部呼叫）
 * @returns {String} HTML原始碼
 */
function getImportedHTML() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    return null;
  }
  
  const htmlCode = sheet.getRange(4, 1).getValue();
  return htmlCode || null;
}