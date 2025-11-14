const SheetsLogger = {
  
  /**
   * 記錄結果到Sheet
   */
  logResults: function(fileName, results, duration) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // 確保有標題列
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['匯入時間', '檔案名稱', '分頁名稱', '類型', '筆數', 'URL', '狀態', '處理時間(秒)']);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#d9d9d9');
    }
    
    const timestamp = new Date();
    
    // 寫入每個分頁的結果
    results.forEach(result => {
      sheet.appendRow([
        timestamp,
        fileName,
        result.sheetName,
        result.type,
        result.recordCount,
        result.url,
        result.status,
        duration
      ]);
    });
    
    // 自動調整欄寬
    sheet.autoResizeColumns(1, 8);
  }
};
