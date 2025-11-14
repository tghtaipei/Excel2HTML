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
  HTML_FOLDER_NAME: '長照特約單位HTML下載',
  APP_VERSION: 'Drive TXT Export v1.0',
  
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

    // 儲存HTML至Google Drive文字檔
    const htmlFile = saveHtmlToDrive(integratedHTML);
    
    // 計算統計資訊
    const endTime = new Date();
    const processingTime = ((endTime - startTime) / 1000).toFixed(2);
    const totalInstitutions = workbookData.sheets.reduce(function(sum, sheet) { return sum + sheet.dataCount; }, 0);
    
    // 寫入Google Sheet
    writeToSheet(workbookData, htmlFile, processingTime, totalInstitutions);
    
    return {
      success: true,
      message: '匯入成功！',
      details: {
        分頁數: workbookData.sheets.length,
        總機構數: totalInstitutions,
        處理時間: processingTime + '秒',
        分頁列表: workbookData.sheets.map(function(s) { return s.name + ' (' + s.dataCount + '筆)'; }),
        TXT下載連結: htmlFile.url,
        程式版本: CONFIG.APP_VERSION
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
 * @param {Object} htmlFile - Google Drive上的HTML檔案資訊
 * @param {String} processingTime - 處理時間
 * @param {Number} totalInstitutions - 總機構數
 */
function writeToSheet(workbookData, htmlFile, processingTime, totalInstitutions) {
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
  sheet.setColumnWidth(7, 200); // 程式版本
  
  // 寫入第一列：轉換資訊
  const metaData = [
    [
      '處理時間(秒)',
      '總分頁數',
      '成功數',
      '失敗數',
      '總機構數',
      '處理日期',
      '程式版本'
    ],
    [
      processingTime,
      workbookData.totalSheets,
      workbookData.sheets.length,
      workbookData.totalSheets - workbookData.sheets.length,
      totalInstitutions,
      new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }),
      CONFIG.APP_VERSION
    ]
  ];

  sheet.getRange(1, 1, 2, 7).setValues(metaData);
  
  // 設定第一列格式
  sheet.getRange(1, 1, 1, 7)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange(2, 1, 1, 7)
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');
  
  // 寫入HTML檔案資訊與連結
  sheet.getRange(3, 1, 1, 5).setValues([[
    'HTML檔案名稱',
    '檔案大小 (KB)',
    '建立時間',
    'Drive 頁面連結',
    '直接下載連結'
  ]]);

  sheet.getRange(3, 1, 1, 5)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange(4, 1, 1, 5).setValues([[
    htmlFile.name,
    (htmlFile.size / 1024).toFixed(2),
    Utilities.formatDate(htmlFile.createdAt, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss'),
    '',
    ''
  ]]);

  sheet.getRange(4, 1, 1, 3)
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');

  sheet.getRange(4, 4)
    .setFormula('=HYPERLINK("' + htmlFile.url + '", "檢視 Drive 檔案")')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');

  sheet.getRange(4, 5)
    .setFormula('=HYPERLINK("' + htmlFile.downloadUrl + '", "直接下載 TXT")')
    .setBackground('#e8f0fe')
    .setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 250);
  sheet.setColumnWidth(5, 250);

  // 使用說明
  sheet.getRange(6, 1)
    .setValue('📝 使用說明：\n1. 點擊「檢視 Drive 檔案」可確認檔案資訊\n2. 點擊「直接下載 TXT」可獲得純檔案內容，便於嵌入網站\n3. 若需重新產出資料，請重新上傳 Excel 檔案並等待系統更新連結\n\n📂 檔案名稱：' + htmlFile.name + '\n🔗 檢視連結：' + htmlFile.url + '\n⬇️ 下載連結：' + htmlFile.downloadUrl + '\n🆕 程式版本：' + CONFIG.APP_VERSION)
    .setBackground('#fff3cd')
    .setWrap(true)
    .setVerticalAlignment('top');

  // 紀錄最新檔案ID於文件層屬性，供後續程式使用
  PropertiesService.getDocumentProperties().setProperty('LATEST_HTML_FILE_ID', htmlFile.id);

  // 凍結前四列
  sheet.setFrozenRows(4);

  Logger.log('資料已成功寫入Sheet並提供TXT下載連結');
}

/**
 * 將產出的HTML儲存為Google Drive文字檔
 * @param {string} htmlCode - 完整的HTML原始碼
 * @returns {Object} 包含檔案資訊的物件
 */
function saveHtmlToDrive(htmlCode) {
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss');
  const fileName = '長照特約單位資料_' + timestamp + '.txt';
  const blob = Utilities.newBlob(htmlCode, 'text/plain', fileName);
  const folder = ensureHtmlFolder();
  const file = folder.createFile(blob);
  file.setDescription('由長照特約單位匯入系統產生的HTML原始碼');
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    id: file.getId(),
    name: fileName,
    url: file.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId(),
    size: blob.getBytes().length,
    createdAt: new Date()
  };
}

function ensureHtmlFolder() {
  const props = PropertiesService.getDocumentProperties();
  let folderId = props.getProperty('HTML_DOWNLOAD_FOLDER_ID');
  let folder = null;

  if (folderId) {
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (error) {
      Logger.log('找不到既有資料夾，將重新建立: ' + error.toString());
      folder = null;
    }
  }

  if (!folder) {
    const folders = DriveApp.getFoldersByName(CONFIG.HTML_FOLDER_NAME);
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.HTML_FOLDER_NAME);
    props.setProperty('HTML_DOWNLOAD_FOLDER_ID', folder.getId());
  }

  return folder;
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
  const docProps = PropertiesService.getDocumentProperties();
  const fileId = docProps.getProperty('LATEST_HTML_FILE_ID');

  if (!fileId) {
    return null;
  }

  try {
    const file = DriveApp.getFileById(fileId);
    return file.getBlob().getDataAsString('utf-8');
  } catch (error) {
    Logger.log('讀取TXT檔案失敗: ' + error.toString());
    return null;
  }
}

/**
 * 提供前端檢視使用的系統資訊
 * @returns {Object} 包含程式版本與最新TXT檔案資訊
 */
function getAppMetadata() {
  const docProps = PropertiesService.getDocumentProperties();
  const fileId = docProps.getProperty('LATEST_HTML_FILE_ID');
  let latestFile = null;

  if (fileId) {
    try {
      const file = DriveApp.getFileById(fileId);
      latestFile = {
        name: file.getName(),
        url: file.getUrl(),
        downloadUrl: 'https://drive.google.com/uc?export=download&id=' + fileId,
        updatedAt: Utilities.formatDate(file.getLastUpdated(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss')
      };
    } catch (error) {
      Logger.log('讀取最新TXT檔案失敗: ' + error.toString());
    }
  }

  return {
    version: CONFIG.APP_VERSION,
    latestFile: latestFile
  };
}
