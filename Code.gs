/**
 * 開啟側邊欄
 */
function showSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('EXCEL 轉 HTML 表格工具')
      .setWidth(600); // 增加寬度到600px (約視窗一半)
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    SpreadsheetApp.getUi().alert('無法開啟側邊欄: ' + error.toString());
  }
}

/**
 * 建立選單
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('資料匯入')
    .addItem('開啟匯入介面', 'showSidebar')
    .addToUi();
}

/**
 * 處理檔案匯入 - 第一階段：上傳並解析
 */
function handleImport(fileData, fileName) {
  try {
    Logger.log('開始處理檔案: ' + fileName);
    
    // 1. 解析Excel檔案
    const workbook = parseExcelFile(fileData, fileName);
    Logger.log('解析完成，共 ' + workbook.SheetNames.length + ' 個分頁');
    
    // 2. 暫存到 PropertiesService
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('tempWorkbook', JSON.stringify({
      fileName: fileName,
      sheetNames: workbook.SheetNames,
      sheets: workbook.Sheets
    }));
    
    // 3. 返回分頁資訊，請求使用者指定標題列
    return {
      success: true,
      needHeaderInput: true,
      fileName: fileName,
      sheets: workbook.SheetNames.map(function(name, index) {
        const worksheet = workbook.Sheets[name];
        const preview = getSheetPreview(worksheet, 5);
        const suggestedHeaderRow = detectHeaderRow(worksheet);
        return {
          name: name,
          index: index,
          preview: preview,
          suggestedHeaderRow: suggestedHeaderRow
        };
      })
    };
    
  } catch (error) {
    Logger.log('匯入錯誤: ' + error.toString());
    Logger.log('錯誤堆疊: ' + error.stack);
    return { 
      success: false, 
      message: '匯入失敗: ' + error.toString() 
    };
  }
}

/**
 * 處理檔案匯入 - 第二階段：使用者指定標題列後處理
 */
function handleImportWithHeaders(headerRows) {
  try {
    const startTime = new Date();
    
    // 1. 從暫存取回工作簿
    const scriptProperties = PropertiesService.getScriptProperties();
    const tempData = JSON.parse(scriptProperties.getProperty('tempWorkbook'));
    const fileName = tempData.fileName;
    const workbook = {
      SheetNames: tempData.sheetNames,
      Sheets: tempData.sheets
    };
    
    Logger.log('繼續處理檔案: ' + fileName);
    Logger.log('使用者指定標題列: ' + JSON.stringify(headerRows));
    
    // 2. 檢查合併儲存格
    const mergeWarnings = [];
    for (let i = 0; i < workbook.SheetNames.length; i++) {
      const sheetName = workbook.SheetNames[i];
      const headerRow = headerRows[i];
      const worksheet = workbook.Sheets[sheetName];
      
      const hasMerge = checkMergedCells(worksheet, headerRow);
      if (hasMerge) {
        mergeWarnings.push(sheetName);
      }
    }
    
    if (mergeWarnings.length > 0) {
      return {
        success: false,
        message: '以下分頁的標題列存在合併儲存格，請移除合併後重新上傳：\n' + mergeWarnings.join('、')
      };
    }
    
    // 3. 處理所有分頁
    const allSheets = [];
    
    for (let i = 0; i < workbook.SheetNames.length; i++) {
      const sheetName = workbook.SheetNames[i];
      const headerRow = headerRows[i];
      Logger.log('處理分頁: ' + sheetName + '，標題列: ' + headerRow);
      
      const worksheet = workbook.Sheets[sheetName];
      
      // 4. 使用指定的標題列解析
      const parser = ParserFactory.getParser(sheetName, worksheet);
      const parsedData = parser.parseWithHeaderRow(worksheet, sheetName, headerRow);
      Logger.log('解析完成，資料筆數: ' + parsedData.data.length);
      
      allSheets.push({
        name: sheetName,
        type: parser.getTypeName(),
        data: parsedData
      });
    }
    
    // 5. 生成單一HTML
    const htmlContent = HTMLGenerator.generateMultiSheet(allSheets, fileName);
    Logger.log('HTML 生成完成');
    
    // 6. 儲存到Drive
    const cleanFileName = fileName.replace(/\.[^/.]+$/, '');
    const fileUrl = DriveManager.saveHTML(htmlContent, cleanFileName + '.html');
    Logger.log('檔案已儲存: ' + fileUrl);
    
    // 7. 計算總筆數
    let totalRecords = 0;
    for (let i = 0; i < allSheets.length; i++) {
      totalRecords += allSheets[i].data.data.length;
    }
    
    // 8. 寫入結果到Sheet
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    const result = {
      sheetName: '全部(' + allSheets.length + '個分頁)',
      type: 'MultiSheet',
      recordCount: totalRecords,
      url: fileUrl,
      status: 'success'
    };
    
    SheetsLogger.logResults(fileName, [result], duration);
    
    // 9. 清除暫存
    scriptProperties.deleteProperty('tempWorkbook');
    
    return { 
      success: true, 
      message: '成功匯入 ' + allSheets.length + ' 個分頁，共 ' + totalRecords + ' 筆資料', 
      results: [result]
    };
    
  } catch (error) {
    Logger.log('匯入錯誤: ' + error.toString());
    Logger.log('錯誤堆疊: ' + error.stack);
    return { 
      success: false, 
      message: '匯入失敗: ' + error.toString() 
    };
  }
}

/**
 * 取得分頁預覽（前N列）
 */
function getSheetPreview(worksheet, rows) {
  const data = worksheet.data;
  const preview = [];
  
  for (let i = 0; i < Math.min(rows, data.length); i++) {
    preview.push(data[i]);
  }
  
  return preview;
}


/**
 * 自動偵測標題列（找出第一個每行都有值的列）
 */
function detectHeaderRow(worksheet) {
  const data = worksheet.data;
  
  if (data.length === 0) {
    return 1; // 預設第一列
  }
  
  // 尋找第一個每行都有值的列
  for (let i = 0; i < Math.min(10, data.length); i++) {
    const row = data[i];
    let allCellsHaveValue = true;
    let nonEmptyCount = 0;
    
    // 檢查這一列的每個儲存格
    for (let j = 0; j < row.length; j++) {
      const cell = row[j];
      
      // 如果儲存格有值
      if (cell !== null && cell !== undefined && String(cell).trim() !== '') {
        nonEmptyCount++;
      } else {
        // 如果前面已經有內容，但這個是空的，代表不是標題列
        if (nonEmptyCount > 0) {
          allCellsHaveValue = false;
          break;
        }
      }
    }
    
    // 如果這一列至少有3個欄位有值，且每個有效欄位都有值，就是標題列
    if (allCellsHaveValue && nonEmptyCount >= 3) {
      return i + 1; // 轉換為 1-based 索引
    }
  }
  
  // 如果找不到，預設第一列
  return 1;
}

/**
 * 檢查指定列是否有合併儲存格
 */
function checkMergedCells(worksheet, headerRow) {
  const data = worksheet.data;
  if (headerRow < 1 || headerRow > data.length) {
    return false;
  }
  
  const row = data[headerRow - 1]; // 轉為0-based索引
  
  // 簡單檢查：如果連續多個空白儲存格，可能是合併的結果
  let emptyCount = 0;
  let hasContent = false;
  
  for (let i = 0; i < row.length; i++) {
    const cell = row[i];
    if (cell === null || cell === undefined || String(cell).trim() === '') {
      if (hasContent) {
        emptyCount++;
        if (emptyCount >= 2) {
          // 有內容後出現連續2個以上空白，可能是合併
          return true;
        }
      }
    } else {
      hasContent = true;
      emptyCount = 0;
    }
  }
  
  return false;
}

/**
 * 清除資料
 */
function handleClear() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }
    
    return { success: true, message: '資料已清除' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 解析Excel檔案 (使用Drive API轉換)
 */
function parseExcelFile(base64Data, fileName) {
  try {
    // 移除data URL前綴
    const base64 = base64Data.split(',')[1];
    const bytes = Utilities.base64Decode(base64);
    
    // 建立臨時檔案
    const blob = Utilities.newBlob(bytes, 
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
      fileName);
    
    // 上傳到Drive (臨時)
    const tempFile = DriveApp.createFile(blob);
    const fileId = tempFile.getId();
    
    // 轉換為Google Sheets格式
    const resource = {
      title: 'temp_' + fileName,
      mimeType: 'application/vnd.google-apps.spreadsheet'
    };
    
    const spreadsheet = Drive.Files.copy(resource, fileId);
    const ssId = spreadsheet.id;
    
    // 開啟轉換後的試算表
    const ss = SpreadsheetApp.openById(ssId);
    const sheets = ss.getSheets();
    
    // 讀取所有分頁資料
    const workbook = {
      SheetNames: [],
      Sheets: {}
    };
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      workbook.SheetNames.push(sheetName);
      
      // 讀取資料
      const data = sheet.getDataRange().getValues();
      workbook.Sheets[sheetName] = {
        data: data,
        name: sheetName
      };
    });
    
    // 刪除臨時檔案
    tempFile.setTrashed(true);
    DriveApp.getFileById(ssId).setTrashed(true);
    
    return workbook;
    
  } catch (error) {
    throw new Error('解析Excel檔案失敗: ' + error.toString());
  }
}
