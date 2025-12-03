const HTMLGenerator = {
  
  /**
   * 生成多分頁HTML（完整獨立HTML檔案）
   */
  generateMultiSheet: function(allSheets, fileName) {
    let html = '<!DOCTYPE html>\n';
    html += '<html lang="zh-TW">\n';
    html += '<head>\n';
    html += '  <meta charset="UTF-8">\n';
    html += '  <meta name="viewport" content="width=device-width, initial-scale=1.0">\n';
    html += '  <title>' + this.escapeHtml(fileName) + '</title>\n';
    html += '  <style>\n';
    html += this.getStyles();
    html += '  </style>\n';
    html += '</head>\n';
    html += '<body>\n';
    html += '  <div class="container">\n';
    
    // 分頁切換按鈕
    html += '    <div class="tab-buttons">\n';
    for (let i = 0; i < allSheets.length; i++) {
      const activeClass = i === 0 ? ' active' : '';
      html += '      <button class="tab-btn' + activeClass + '" onclick="switchTab(' + i + ')">' + this.escapeHtml(allSheets[i].name) + '</button>\n';
    }
    html += '    </div>\n';
    html += '    \n';
    
    // 各分頁內容
    for (let i = 0; i < allSheets.length; i++) {
      const sheet = allSheets[i];
      const displayStyle = i === 0 ? 'block' : 'none';
      
      html += '    <div class="tab-content" id="tab' + i + '" data-title="' + this.escapeHtml(sheet.name) + '" style="display:' + displayStyle + '">\n';
      html += this.generateSheetContent(sheet.data, i);
      html += '    </div>\n';
    }
    
    html += '  </div>\n';
    html += '  \n';
    html += '  <script>\n';
    html += this.getScripts(allSheets);
    html += '  </script>\n';
    html += '</body>\n';
    html += '</html>';
    
    return html;
  },
  
  /**
   * 生成單一分頁內容
   */
  generateSheetContent: function(parsedData, tabIndex) {
    const title = parsedData.title;
    const headers = parsedData.headers;
    const data = parsedData.data;
    
    // 檢測是否有行政區欄位
    const districtInfo = this.detectDistricts(data);
    
    // 檢測勾選欄位
    const checkboxColumns = this.detectCheckboxColumns(headers, data);
    
    let html = '';
    
    // 搜尋框
    html += '      <div class="search-box">\n';
    html += '        <input type="text" id="searchInput' + tabIndex + '" placeholder="輸入關鍵字搜尋..." class="search-input">\n';
    html += '        <button onclick="search(' + tabIndex + ')" class="btn btn-primary">搜尋</button>\n';
    html += '        <button onclick="clearSearch(' + tabIndex + ')" class="btn btn-secondary">清除</button>\n';
    html += '        <button onclick="exportPDF(' + tabIndex + ')" class="btn btn-export">📥 匯出PDF</button>\n';
    html += '        <span id="searchStats' + tabIndex + '" class="search-stats"></span>\n';
    html += '      </div>\n';
    html += '      \n';
    
    // 篩選區域
    if (districtInfo.hasDistrict || checkboxColumns.length > 0) {
      html += '      <div class="filters">\n';
      
      // 行政區篩選
      if (districtInfo.hasDistrict) {
        html += '        <div class="filter-section">\n';
        html += '          <div class="filter-title">行政區：</div>\n';
        html += '          <div class="filter-options">\n';
        for (let i = 0; i < districtInfo.districts.length; i++) {
          const district = districtInfo.districts[i];
          html += '            <label class="filter-label"><input type="checkbox" class="district-filter" data-tab="' + tabIndex + '" value="' + this.escapeHtml(district) + '" onchange="applyFilters(' + tabIndex + ')">' + this.escapeHtml(district) + '</label>\n';
        }
        // 如果有非台北市資料，增加"非臺北市"選項
        if (districtInfo.hasNonTaipei) {
          html += '            <label class="filter-label"><input type="checkbox" class="district-filter" data-tab="' + tabIndex + '" value="非臺北市" onchange="applyFilters(' + tabIndex + ')"><strong style="color: #d32f2f;">非臺北市</strong></label>\n';
        }
        html += '          </div>\n';
        html += '        </div>\n';
      }
      
      // 其他過濾條件
      if (checkboxColumns.length > 0) {
        html += '        <div class="filter-section">\n';
        html += '          <div class="filter-title">其他過濾條件：</div>\n';
        html += '          <div class="filter-options">\n';
        for (let i = 0; i < checkboxColumns.length; i++) {
          const col = checkboxColumns[i];
          html += '            <label class="filter-label"><input type="checkbox" class="column-filter" data-tab="' + tabIndex + '" data-column="' + col.index + '" onchange="applyFilters(' + tabIndex + ')"><strong>' + this.escapeHtml(col.name) + '</strong></label>\n';
        }
        html += '          </div>\n';
        html += '        </div>\n';
      }
      
      html += '      </div>\n';
      html += '      \n';
    }
    
    // 資料表格
    html += '      <table class="data-table" id="dataTable' + tabIndex + '">\n';
    html += '        <thead>\n';
    html += '          <tr>\n';
    for (let i = 0; i < headers.length; i++) {
      html += '            <th>' + this.escapeHtml(headers[i]) + '</th>\n';
    }
    html += '          </tr>\n';
    html += '        </thead>\n';
    html += '        <tbody>\n';
    
    // 只輸出有內容的資料列
    for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
      const row = data[rowIndex];
      
      // 檢查是否為空白列
      let isEmpty = true;
      for (let cellIndex = 0; cellIndex < row.length; cellIndex++) {
        const cellValue = row[cellIndex];
        if (cellValue !== null && cellValue !== undefined && String(cellValue).trim() !== '') {
          isEmpty = false;
          break;
        }
      }
      
      if (!isEmpty) {
        html += '          <tr data-row="' + rowIndex + '"';
        
        // 添加行政區屬性
        if (districtInfo.hasDistrict) {
          const rowDistrict = this.findDistrictInRow(row);
          if (rowDistrict) {
            html += ' data-district="' + this.escapeHtml(rowDistrict) + '"';
          }
        }
        
        // 添加勾選欄位屬性
        for (let i = 0; i < checkboxColumns.length; i++) {
          const col = checkboxColumns[i];
          const cellValue = String(row[col.index]).trim().toLowerCase();
          const isChecked = this.isCheckboxChecked(cellValue);
          html += ' data-checkbox-' + col.index + '="' + (isChecked ? 'checked' : 'unchecked') + '"';
        }
        
        html += '>\n';
        
        for (let cellIndex = 0; cellIndex < row.length; cellIndex++) {
          html += '            <td>' + this.escapeHtml(String(row[cellIndex])) + '</td>\n';
        }
        html += '          </tr>\n';
      }
    }
    
    html += '        </tbody>\n';
    html += '      </table>\n';
    html += '      \n';
    html += '      <div class="footer">\n';
    html += '        總計 <span id="count' + tabIndex + '">' + data.length + '</span> 筆資料\n';
    html += '      </div>\n';
    
    return html;
  },
  
  /**
   * 偵測行政區
   */
  detectDistricts: function(data) {
    const districts = ['中正', '中山', '萬華', '信義', '大安', '文山', '內湖', '南港', '北投', '士林', '大同', '松山'];
    const nonTaipeiKeywords = ['新北市', '基隆市', '桃園市', '花蓮縣'];
    const foundDistricts = [];
    let hasNonTaipei = false;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      for (let j = 0; j < row.length; j++) {
        const cellValue = String(row[j]);
        
        // 檢測台北市行政區
        for (let k = 0; k < districts.length; k++) {
          const district = districts[k];
          if (cellValue.indexOf(district) !== -1 && foundDistricts.indexOf(district) === -1) {
            foundDistricts.push(district);
          }
        }
        
        // 檢測非台北市關鍵字
        if (!hasNonTaipei) {
          for (let k = 0; k < nonTaipeiKeywords.length; k++) {
            const keyword = nonTaipeiKeywords[k];
            const keywordIndex = cellValue.indexOf(keyword);
            if (keywordIndex !== -1) {
              // 檢查關鍵字後5個字內是否符合條件
              const afterKeyword = cellValue.substring(keywordIndex + keyword.length, keywordIndex + keyword.length + 5);
              
              // 花蓮縣：需要有"市"或"鄉"
              if (keyword === '花蓮縣') {
                if (afterKeyword.indexOf('市') !== -1 || afterKeyword.indexOf('鄉') !== -1) {
                  hasNonTaipei = true;
                  break;
                }
              } 
              // 其他（新北市、基隆市、桃園市）：需要有"區"
              else {
                if (afterKeyword.indexOf('區') !== -1) {
                  hasNonTaipei = true;
                  break;
                }
              }
            }
          }
        }
      }
    }
    
    return {
      hasDistrict: foundDistricts.length > 0,
      districts: foundDistricts,
      hasNonTaipei: hasNonTaipei && foundDistricts.length > 0  // 只有在有台北市行政區時才顯示
    };
  },
  
  /**
   * 在資料列中尋找行政區
   */
  findDistrictInRow: function(row) {
    const districts = ['中正', '中山', '萬華', '信義', '大安', '文山', '內湖', '南港', '北投', '士林', '大同', '松山'];
    const nonTaipeiKeywords = ['新北市', '基隆市', '桃園市', '花蓮縣'];
    
    let foundDistrict = null;
    let hasNonTaipei = false;
    
    for (let j = 0; j < row.length; j++) {
      const cellValue = String(row[j]);
      
      // 檢查台北市行政區
      if (!foundDistrict) {
        for (let k = 0; k < districts.length; k++) {
          const district = districts[k];
          if (cellValue.indexOf(district) !== -1) {
            foundDistrict = district;
            break;
          }
        }
      }
      
      // 檢查非台北市關鍵字
      if (!hasNonTaipei) {
        for (let k = 0; k < nonTaipeiKeywords.length; k++) {
          const keyword = nonTaipeiKeywords[k];
          const keywordIndex = cellValue.indexOf(keyword);
          if (keywordIndex !== -1) {
            // 檢查關鍵字後5個字內是否符合條件
            const afterKeyword = cellValue.substring(keywordIndex + keyword.length, keywordIndex + keyword.length + 5);
            
            // 花蓮縣：需要有"市"或"鄉"
            if (keyword === '花蓮縣') {
              if (afterKeyword.indexOf('市') !== -1 || afterKeyword.indexOf('鄉') !== -1) {
                hasNonTaipei = true;
                break;
              }
            } 
            // 其他（新北市、基隆市、桃園市）：需要有"區"
            else {
              if (afterKeyword.indexOf('區') !== -1) {
                hasNonTaipei = true;
                break;
              }
            }
          }
        }
      }
    }
    
    // 如果有非台北市關鍵字，返回"非臺北市"
    if (hasNonTaipei) {
      return '非臺北市';
    }
    
    return foundDistrict;
  },
  
  /**
   * 偵測勾選欄位
   */
  detectCheckboxColumns: function(headers, data) {
    const checkboxColumns = [];
    
    for (let colIndex = 0; colIndex < headers.length; colIndex++) {
      let checkedCount = 0;
      let uncheckedCount = 0;
      let sampleSize = Math.min(data.length, 50);
      
      for (let rowIndex = 0; rowIndex < sampleSize; rowIndex++) {
        if (data[rowIndex] && data[rowIndex][colIndex] !== undefined) {
          const cellValue = String(data[rowIndex][colIndex]).trim().toLowerCase();
          if (this.isCheckboxChecked(cellValue)) {
            checkedCount++;
          } else if (this.isCheckboxUnchecked(cellValue)) {
            uncheckedCount++;
          }
        }
      }
      
      // 如果超過30%的資料是勾選相關，認定為勾選欄位
      if (checkedCount + uncheckedCount > sampleSize * 0.3) {
        checkboxColumns.push({
          index: colIndex,
          name: headers[colIndex]
        });
      }
    }
    
    return checkboxColumns;
  },
  
  /**
   * 判斷是否為已勾選
   */
  isCheckboxChecked: function(value) {
    const checkedValues = ['v', 'V', '✓', '✔', '√', 'true', 'yes', '是', 'o', 'O', '●', '⊙'];
    return checkedValues.indexOf(value) !== -1 || value === 'v';
  },
  
  /**
   * 判斷是否為未勾選
   */
  isCheckboxUnchecked: function(value) {
    return value === '' || value === 'x' || value === 'X' || value === 'false' || value === 'no' || value === '否';
  },
  
  /**
 * CSS樣式
 */
getStyles: function() {
  let css = '';
  
  // 基礎設定
  css += '* { box-sizing: border-box; margin: 0; padding: 0; }\n';
  css += 'body { font-family: Arial, "Microsoft JhengHei", sans-serif; background: linear-gradient(135deg, #e0f7fa 0%, #b3e5fc 100%); padding: 20px; min-height: 100vh; font-size: 18px; line-height: 1.2; }\n';
  css += '.container { max-width: 1400px; margin: 0 auto; background: white; padding: 20px; border-radius: 12px; box-shadow: 0 10px 40px rgba(0,0,0,0.2); }\n';
  
  // 分頁按鈕
  css += '.tab-buttons { display: flex; gap: 8px; margin-bottom: 15px; border-bottom: 3px solid #b3e5fc; padding-bottom: 10px; flex-wrap: wrap; }\n';
  css += '.tab-btn { padding: 12px 24px; background: linear-gradient(135deg, #4fc3f7 0%, #0288d1 100%); color: white; border: none; border-radius: 8px; cursor: pointer; font-size: 18px; font-weight: 600; transition: all 0.3s; box-shadow: 0 2px 8px rgba(3, 169, 244, 0.3); line-height: 1.2; }\n';
  css += '.tab-btn:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(3, 169, 244, 0.4); }\n';
  css += '.tab-btn.active { background: linear-gradient(135deg, #0288d1 0%, #01579b 100%); box-shadow: 0 4px 12px rgba(2, 136, 209, 0.5); }\n';
  css += '.tab-content { display: none; }\n';
  
  // 搜尋框
  css += '.search-box { margin-bottom: 15px; padding: 18px; background: linear-gradient(135deg, #e1f5fe 0%, #b3e5fc 100%); border-radius: 10px; display: flex; gap: 12px; align-items: center; box-shadow: 0 2px 8px rgba(3, 169, 244, 0.2); flex-wrap: wrap; }\n';
  css += '.search-input { flex: 1; min-width: 250px; padding: 12px 18px; border: 2px solid #4fc3f7; border-radius: 8px; font-size: 18px; background: white; transition: all 0.3s; line-height: 1.2; }\n';
  css += '.search-input:focus { outline: none; border-color: #0288d1; box-shadow: 0 0 0 3px rgba(2, 136, 209, 0.1); }\n';
  css += '.btn { padding: 12px 24px; border: none; border-radius: 8px; cursor: pointer; font-weight: 600; font-size: 18px; transition: all 0.3s; white-space: nowrap; line-height: 1.2; }\n';
  css += '.btn-primary { background: linear-gradient(135deg, #0288d1 0%, #01579b 100%); color: white; }\n';
  css += '.btn-primary:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(2, 136, 209, 0.4); }\n';
  css += '.btn-secondary { background: linear-gradient(135deg, #00acc1 0%, #00838f 100%); color: white; }\n';
  css += '.btn-secondary:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0, 172, 193, 0.4); }\n';
  css += '.btn-export { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); color: white; }\n';
  css += '.btn-export:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(245, 87, 108, 0.4); }\n';
  css += '.search-stats { color: #d32f2f; font-size: 18px; font-weight: 600; margin-left: auto; line-height: 1.2; }\n';
  
  // 篩選區域
  css += '.filters { margin-bottom: 15px; padding: 18px; background: linear-gradient(135deg, #b3e5fc 0%, #81d4fa 100%); border-radius: 10px; box-shadow: 0 2px 8px rgba(3, 169, 244, 0.2); }\n';
  css += '.filter-section { margin-bottom: 15px; }\n';
  css += '.filter-section:last-child { margin-bottom: 0; }\n';
  css += '.filter-title { color: #01579b; margin-bottom: 10px; font-size: 18px; font-weight: 700; line-height: 1.2; }\n';
  css += '.filter-options { display: flex; flex-wrap: wrap; gap: 18px; }\n';
  css += '.filter-label { cursor: pointer; color: #01579b; font-size: 18px; transition: all 0.2s; display: inline-flex; align-items: center; line-height: 1.2; }\n';
  css += '.filter-label:hover { color: #0288d1; }\n';
  css += '.filter-label input[type="checkbox"] { margin-right: 8px; width: 20px; height: 20px; cursor: pointer; }\n';
  
  // 資料表格
  css += '.data-table { width: 100%; border-collapse: collapse; margin-bottom: 15px; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }\n';
  css += '.data-table th { background: linear-gradient(135deg, #4fc3f7 0%, #0288d1 100%); color: white; padding: 16px 14px; text-align: left; font-weight: 600; font-size: 18px; position: sticky; top: 0; z-index: 10; line-height: 1.2; }\n';
  css += '.data-table td { padding: 14px; border-bottom: 1px solid #e1f5fe; color: #263238; font-size: 18px; line-height: 1.2; }\n';
  css += '.data-table tr:nth-child(even) { background: #f1f8fb; }\n';
  css += '.data-table tr:hover { background: #e1f5fe; transition: background 0.2s; }\n';
  css += '.data-table tr.hidden { display: none; }\n';
  
  // 高亮標記
  css += '.highlight { background: #fff59d; font-weight: bold; padding: 3px 6px; border-radius: 3px; }\n';
  css += '.current-highlight { background: #ffeb3b; }\n';
  
  // 頁尾
  css += '.footer { text-align: center; color: #455a64; font-size: 18px; padding: 18px; background: linear-gradient(135deg, #e1f5fe 0%, #b3e5fc 100%); border-radius: 8px; font-weight: 600; line-height: 1.2; }\n';
  css += '.footer span { color: #d32f2f; font-size: 20px; font-weight: 700; }\n';
  


  // 列印樣式
  css += '@media print {\n';
  css += '  body { background: white !important; }\n';
  css += '  .container { max-width: none !important; padding: 20px !important; box-shadow: none !important; }\n';
  css += '  \n';
  css += '  /* 隱藏UI元素 */\n';
  css += '  .header { display: none !important; }\n';
  css += '  .tab-buttons { display: none !important; }\n';
  css += '  .search-box { display: none !important; }\n';
  css += '  .filters { display: none !important; }\n';
  css += '  .footer { display: none !important; }\n';
  css += '  .btn { display: none !important; }\n';
  css += '  .stats { display: none !important; }\n';
  css += '  \n';
  css += '  /* 預設隱藏所有分頁 */\n';
  css += '  .tab-content { display: none !important; }\n';
  css += '  \n';
  css += '  /* 只顯示標記為 printing-active 的分頁 */\n';
  css += '  .tab-content.printing-active { display: block !important; }\n';
  css += '  \n';
  css += '  /* 在分頁內容前顯示標題 */\n';
  css += '  .tab-content.printing-active::before { \n';
  css += '    content: attr(data-title); \n';
  css += '    display: block; \n';
  css += '    font-size: 20px; \n';
  css += '    font-weight: bold; \n';
  css += '    text-align: center; \n';
  css += '    margin-bottom: 20px; \n';
  css += '    padding-bottom: 10px; \n';
  css += '    border-bottom: 2px solid #333; \n';
  css += '  }\n';
  css += '  \n';
  css += '  /* 隱藏被篩選的資料列 */\n';
  css += '  .data-table tr.hidden { display: none !important; }\n';
  css += '  \n';
  css += '  /* 表格樣式 */\n';
  css += '  .data-table { \n';
  css += '    width: 100% !important; \n';
  css += '    border-collapse: collapse !important; \n';
  css += '    font-size: 10px !important; \n';
  css += '    page-break-inside: auto;\n';
  css += '  }\n';
  css += '  \n';
  css += '  .data-table thead { display: table-header-group; }\n';
  css += '  .data-table tr { page-break-inside: avoid; page-break-after: auto; }\n';
  css += '  \n';
  css += '  .data-table th { \n';
  css += '    background: #d0d0d0 !important; \n';
  css += '    color: #000 !important; \n';
  css += '    border: 1px solid #333 !important; \n';
  css += '    padding: 8px !important;\n';
  css += '    font-weight: bold !important;\n';
  css += '    text-align: center !important;\n';
  css += '    -webkit-print-color-adjust: exact !important;\n';
  css += '    print-color-adjust: exact !important;\n';
  css += '  }\n';
  css += '  \n';
  css += '  .data-table td { \n';
  css += '    border: 1px solid #666 !important; \n';
  css += '    padding: 6px !important;\n';
  css += '    color: #000 !important;\n';
  css += '  }\n';
  css += '  \n';
  css += '  @page { \n';
  css += '    size: A4 landscape; \n';
  css += '    margin: 1.5cm;\n';
  css += '  }\n';
  css += '}\n';
  css += '\n';
  css += '}\n';
  css += '\n';
  
  // 響應式設計
  css += '@media screen and (max-width: 768px) {\n';
  css += '  body { font-size: 16px; padding: 10px; }\n';
  css += '  .container { padding: 15px; }\n';
  css += '  .tab-btn { padding: 10px 18px; font-size: 16px; }\n';
  css += '  .search-input { font-size: 16px; padding: 10px 15px; }\n';
  css += '  .btn { padding: 10px 18px; font-size: 16px; }\n';
  css += '  .search-stats { font-size: 16px; }\n';
  css += '  .filter-title { font-size: 16px; }\n';
  css += '  .filter-label { font-size: 16px; }\n';
  css += '  .filter-label input[type="checkbox"] { width: 18px; height: 18px; }\n';
  css += '  .data-table th { font-size: 16px; padding: 12px 10px; }\n';
  css += '  .data-table td { font-size: 16px; padding: 10px; }\n';
  css += '  .footer { font-size: 16px; }\n';
  css += '  .footer span { font-size: 18px; }\n';
  css += '}\n';
  
  return css;
},
  
  /**
   * JavaScript功能
   */
  getScripts: function(allSheets) {
    let js = '';
    js += 'var currentTab = 0;\n';
    js += '\n';
    js += 'function switchTab(tabIndex) {\n';
    js += '  var contents = document.querySelectorAll(".tab-content");\n';
    js += '  for (var i = 0; i < contents.length; i++) {\n';
    js += '    contents[i].style.display = "none";\n';
    js += '  }\n';
    js += '  var buttons = document.querySelectorAll(".tab-btn");\n';
    js += '  for (var i = 0; i < buttons.length; i++) {\n';
    js += '    buttons[i].classList.remove("active");\n';
    js += '  }\n';
    js += '  document.getElementById("tab" + tabIndex).style.display = "block";\n';
    js += '  buttons[tabIndex].classList.add("active");\n';
    js += '  currentTab = tabIndex;\n';
    js += '}\n';
    js += '\n';
    js += 'function search(tabIndex) {\n';
    js += '  var keyword = document.getElementById("searchInput" + tabIndex).value.trim();\n';
    js += '  if (!keyword) return;\n';
    js += '  var table = document.getElementById("dataTable" + tabIndex);\n';
    js += '  var rows = table.querySelectorAll("tbody tr");\n';
    js += '  var matchCount = 0;\n';
    js += '  for (var i = 0; i < rows.length; i++) {\n';
    js += '    var row = rows[i];\n';
    js += '    var cells = row.querySelectorAll("td");\n';
    js += '    var hasMatch = false;\n';
    js += '    for (var j = 0; j < cells.length; j++) {\n';
    js += '      var cell = cells[j];\n';
    js += '      var text = cell.textContent;\n';
    js += '      if (text.indexOf(keyword) !== -1) {\n';
    js += '        hasMatch = true;\n';
    js += '        matchCount++;\n';
    js += '        var regex = new RegExp("(" + escapeRegex(keyword) + ")", "gi");\n';
    js += '        cell.innerHTML = text.replace(regex, "<span class=\\"highlight\\">$1</span>");\n';
    js += '      }\n';
    js += '    }\n';
    js += '    if (hasMatch) { row.classList.remove("hidden"); } else { row.classList.add("hidden"); }\n';
    js += '  }\n';
    js += '  document.getElementById("searchStats" + tabIndex).textContent = "找到 " + matchCount + " 個符合項目";\n';
    js += '  updateCount(tabIndex);\n';
    js += '}\n';
    js += '\n';
    js += 'function clearSearch(tabIndex) {\n';
    js += '  document.getElementById("searchInput" + tabIndex).value = "";\n';
    js += '  document.getElementById("searchStats" + tabIndex).textContent = "";\n';
    js += '  var table = document.getElementById("dataTable" + tabIndex);\n';
    js += '  var rows = table.querySelectorAll("tbody tr");\n';
    js += '  for (var i = 0; i < rows.length; i++) {\n';
    js += '    var cells = rows[i].querySelectorAll("td");\n';
    js += '    for (var j = 0; j < cells.length; j++) {\n';
    js += '      cells[j].innerHTML = cells[j].textContent;\n';
    js += '    }\n';
    js += '  }\n';
    js += '  applyFilters(tabIndex);\n';
    js += '}\n';
    js += '\n';
    js += 'function applyFilters(tabIndex) {\n';
    js += '  var table = document.getElementById("dataTable" + tabIndex);\n';
    js += '  var rows = table.querySelectorAll("tbody tr");\n';
    js += '  var districtFilters = [];\n';
    js += '  var districtCheckboxes = document.querySelectorAll(".district-filter[data-tab=\\"" + tabIndex + "\\"]:checked");\n';
    js += '  for (var i = 0; i < districtCheckboxes.length; i++) { districtFilters.push(districtCheckboxes[i].value); }\n';
    js += '  var columnFilters = [];\n';
    js += '  var columnCheckboxes = document.querySelectorAll(".column-filter[data-tab=\\"" + tabIndex + "\\"]:checked");\n';
    js += '  for (var i = 0; i < columnCheckboxes.length; i++) { columnFilters.push(columnCheckboxes[i].getAttribute("data-column")); }\n';
    js += '  for (var i = 0; i < rows.length; i++) {\n';
    js += '    var row = rows[i];\n';
    js += '    var show = true;\n';
    js += '    if (districtFilters.length > 0) {\n';
    js += '      var rowDistrict = row.getAttribute("data-district");\n';
    js += '      if (districtFilters.indexOf(rowDistrict) === -1) { show = false; }\n';
    js += '    }\n';
    js += '    for (var j = 0; j < columnFilters.length; j++) {\n';
    js += '      var col = columnFilters[j];\n';
    js += '      var rowValue = row.getAttribute("data-checkbox-" + col);\n';
    js += '      if (rowValue !== "checked") { show = false; break; }\n';
    js += '    }\n';
    js += '    if (show) { row.classList.remove("hidden"); } else { row.classList.add("hidden"); }\n';
    js += '  }\n';
    js += '  updateCount(tabIndex);\n';
    js += '}\n';
    js += '\n';
    js += 'function updateCount(tabIndex) {\n';
    js += '  var table = document.getElementById("dataTable" + tabIndex);\n';
    js += '  var visibleRows = table.querySelectorAll("tbody tr:not(.hidden)");\n';
    js += '  document.getElementById("count" + tabIndex).textContent = visibleRows.length;\n';
    js += '}\n';
    js += '\n';
    js += 'function escapeRegex(str) {\n';
    js += '  return str.replace(/[.*+?^${}()|[\\]\\\\]/g, "\\\\$&");\n';
    js += '}\n';
    js += '\n';
    js += 'function exportPDF(tabIndex) {\n';
    js += '  // 取得分頁\n';
    js += '  var tabContent = document.getElementById("tab" + tabIndex);\n';
    js += '  if (!tabContent) {\n';
    js += '    alert("找不到分頁");\n';
    js += '    return;\n';
    js += '  }\n';
    js += '  \n';
    js += '  // 檢查是否有可見資料\n';
    js += '  var table = document.getElementById("dataTable" + tabIndex);\n';
    js += '  if (!table) {\n';
    js += '    alert("找不到表格");\n';
    js += '    return;\n';
    js += '  }\n';
    js += '  \n';
    js += '  var visibleRows = table.querySelectorAll("tbody tr:not(.hidden)");\n';
    js += '  if (visibleRows.length === 0) {\n';
    js += '    alert("沒有可匯出的資料（請檢查篩選條件）");\n';
    js += '    return;\n';
    js += '  }\n';
    js += '  \n';
    js += '  // 取得分頁名稱\n';
    js += '  var tabBtn = document.querySelectorAll(".tab-btn")[tabIndex];\n';
    js += '  var sheetName = tabBtn ? tabBtn.textContent.trim() : "表格";\n';
    js += '  \n';
    js += '  // 儲存原始狀態\n';
    js += '  var allContents = document.querySelectorAll(".tab-content");\n';
    js += '  var originalDisplays = [];\n';
    js += '  var originalClasses = [];\n';
    js += '  \n';
    js += '  for (var i = 0; i < allContents.length; i++) {\n';
    js += '    originalDisplays[i] = allContents[i].style.display;\n';
    js += '    originalClasses[i] = allContents[i].className;\n';
    js += '  }\n';
    js += '  \n';
    js += '  // 設定要列印的分頁\n';
    js += '  for (var i = 0; i < allContents.length; i++) {\n';
    js += '    if (i === tabIndex) {\n';
    js += '      allContents[i].style.display = "block";\n';
    js += '      allContents[i].className = allContents[i].className + " printing-active";\n';
    js += '    } else {\n';
    js += '      allContents[i].style.display = "none";\n';
    js += '    }\n';
    js += '  }\n';
    js += '  \n';
    js += '  // 設定文件標題\n';
    js += '  var originalTitle = document.title;\n';
    js += '  document.title = sheetName;\n';
    js += '  \n';
    js += '  // 標記正在列印\n';
    js += '  document.body.className = (document.body.className || "") + " is-printing";\n';
    js += '  \n';
    js += '  // 延遲執行列印，確保樣式已套用\n';
    js += '  setTimeout(function() {\n';
    js += '    window.print();\n';
    js += '    \n';
    js += '    // 恢復原始狀態\n';
    js += '    setTimeout(function() {\n';
    js += '      document.title = originalTitle;\n';
    js += '      document.body.className = document.body.className.replace(" is-printing", "");\n';
    js += '      \n';
    js += '      for (var i = 0; i < allContents.length; i++) {\n';
    js += '        allContents[i].style.display = originalDisplays[i];\n';
    js += '        allContents[i].className = originalClasses[i];\n';
    js += '      }\n';
    js += '    }, 500);\n';
    js += '  }, 300);\n';
    js += '}\n';
    js += '\n';
    js += 'document.addEventListener("DOMContentLoaded", function() {\n';
    js += '  switchTab(0);\n';
    js += '});\n';
    return js;
  },
  
  /**
   * HTML跳脫
   */
  escapeHtml: function(text) {
    const map = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'
    };
    return String(text).replace(/[&<>"']/g, function(m) { 
      return map[m]; 
    });
  }
};
