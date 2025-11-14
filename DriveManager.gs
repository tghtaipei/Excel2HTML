const DriveManager = {
  
  /**
   * 儲存HTML到Drive
   */
  saveHTML: function(htmlContent, fileName) {
    try {
      // 取得或建立目標資料夾
      const folder = this.getOrCreateFolder('匯入資料');
      
      // 建立HTML檔案
      const blob = Utilities.newBlob(htmlContent, 'text/html', fileName);
      const file = folder.createFile(blob);
      
      // 設定共享權限 (任何人可檢視)
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      return file.getUrl();
      
    } catch (error) {
      throw new Error('儲存檔案失敗: ' + error.toString());
    }
  },
  
  /**
   * 取得或建立資料夾
   */
  getOrCreateFolder: function(folderName) {
    const folders = DriveApp.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      return folders.next();
    } else {
      return DriveApp.createFolder(folderName);
    }
  }
};
