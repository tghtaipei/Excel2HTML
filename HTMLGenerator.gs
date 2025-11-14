const HTMLGenerator = {
  
  /**
   * 生成多分頁HTML（適合插入Froala Editor）
   */
  generateMultiSheet: function(allSheets, fileName) {
    let html = '';
    
    // 內嵌樣式
    html += '<style>\n';
    html += this.getStyles();
    html += '</style>\n';
    html += '\n';
    
    // 容器開始
    html += '<div class="excel-html-container">\n';
    
    // 分頁切換按鈕
    html += '  <div class="tab-buttons">\n';
    for (let i = 0; i < allSheets.length; i++) {
      const activeClass = i === 0 ? ' active' : '';
      html += '    <button class="tab-btn' + activeClass + '" onclick="switchExcelTab(' + i + ')">' + this.escapeHtml(allSheets[i].name) + '</button>\n';
    }
    html += '  </div>\n';
    html += '  \n';
    
    // 各分頁內容
    for (let i = 0; i < allSheets.length; i++) {
      const sheet = allSheets[i];
      const displayStyle = i === 0 ? 'block' : 'none';
      
      html += '  <div class="excel-tab-content" id="exceltab' + i + '" style="display:' + displayStyle + '">\n';
      html += this.generateSheetContent(sheet.data, i);
      html += '  </div>\n';
    }
    
    html += '</div>\n';
    html += '\n';
    
    // JavaScript
    html += '<script>\n';
    html += this.getScripts(allSheets);
    html += '</script>\n';
    
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
    html += '    <div class="excel-search-box">\n';
    html += '      <input type="text" id="excelSearchInput' + tabIndex + '" placeholder="輸入關鍵字搜尋..." class="excel-search-input">\n';
    html += '      <button onclick="excelSearch(' + tabIndex + ')" class="excel-btn excel-btn-primary">搜尋</button>\n';
    html += '      <button onclick="excelClearSearch(' + tabIndex + ')" class="excel-btn excel-btn-secondary">清除</button>\n';
    html += '      <span id="excelSearchStats' + tabIndex + '" class="excel-search-stats"></span>\n';
    html += '    </div>\n';
    html += '    \n';
    
    // 篩選區域
    if (districtInfo.hasDistrict || checkboxColumns.length > 0) {
      html += '    <div class="excel-filters">\n';
      
      // 行政區篩選
      if (districtInfo.hasDistrict) {
        html += '      <div class="excel-filter-section">\n';
        html += '        <div class="excel-filter-title">行政區：</div>\n';
        html += '        <div class="excel-filter-options">\n';
        for (let i = 0; i < districtInfo.districts.length; i++) {
          const district = districtInfo.districts[i];
          html += '          <label class="excel-filter-label"><input type="checkbox" class="excel-district-filter" data-tab="' + tabIndex + '" value="' + this.escapeHtml(district) + '" onchange="excelApplyFilters(' + tabIndex + ')">' + this.escapeHtml(district) + '</label>\n';
        }
        html += '        </div>\n';
        html += '      </div>\n';
      }
      
      // 其他過濾條件
      if (checkboxColumns.length > 0) {
        html += '      <div class="excel-filter-section">\n';
        html += '        <div class="excel-filter-title">其他過濾條件：</div>\n';
        html += '        <div class="excel-filter-options">\n';
        for (let i = 0; i < checkboxColumns.length; i++) {
          const col = checkboxColumns[i];
          html += '          <label class="excel-filter-label"><input type="checkbox" class="excel-column-filter" data-tab="' + tabIndex + '" data-column="' + col.index + '" onchange="excelApplyFilters(' + tabIndex + ')"><strong>' + this.escapeHtml(col.name) + '</strong></label>\n';
        }
        html += '        </div>\n';
        html += '      </div>\n';
      }
      
      html += '    </div>\n';
      html += '    \n';
    }
    
    // 資料表格
    html += '    <table class="excel-data-table" id="excelDataTable' + tabIndex + '">\n';
    html += '      <thead>\n';
    html += '        <tr>\n';
    for (let i = 0; i < headers.length; i++) {
      html += '          <th>' + this.escapeHtml(headers[i]) + '</th>\n';
    }
    html += '        </tr>\n';
    html += '      </thead>\n';
    html += '      <tbody>\n';
    
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
        html += '        <tr data-row="' + rowIndex + '"';
        
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
          html += '          <td>' + this.escapeHtml(String(row[cellIndex])) + '</td>\n';
        }
        html += '        </tr>\n';
      }
    }
    
    html += '      </tbody>\n';
    html += '    </table>\n';
    html += '    \n';
    html += '    <div class="excel-footer">\n';
    html += '      總計 <span id="excelCount' + tabIndex + '">' + data.length + '</span> 筆資料\n';
    html += '    </div>\n';
    
    return html;
  },
  
  /**
   * 偵測行政區
   */
  detectDistricts: function(data) {
    const districts = ['中正', '中山', '萬華', '信義', '大安', '文山', '內湖', '南港', '北投', '士林', '大同'];
    const foundDistricts = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      for (let j = 0; j < row.length; j++) {
        const cellValue = String(row[j]);
        for (let k = 0; k < districts.length; k++) {
          const district = districts[k];
          if (cellValue.indexOf(district) !== -1 && foundDistricts.indexOf(district) === -1) {
            foundDistricts.push(district);
          }
        }
      }
    }
    
    return {
      hasDistrict: foundDistricts.length > 0,
      districts: foundDistricts
    };
  },
  
  /**
   * 在資料列中尋找行政區
   */
  findDistrictInRow: function(row) {
    const districts = ['中正', '中山', '萬華', '信義', '大安', '文山', '內湖', '南港', '北投', '士林', '大同'];
    
    for (let j = 0; j < row.length; j++) {
      const cellValue = String(row[j]);
      for (let k = 0; k < districts.length; k++) {
        const district = districts[k];
        if (cellValue.indexOf(district) !== -1) {
          return district;
        }
      }
    }
    return null;
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
   * CSS樣式（所有class加上excel-前綴避免衝突）
   */
  getStyles: function() {
    let css = '';
    css += '.excel-html-container { font-family: Arial, "Microsoft JhengHei", sans-serif; max-width: 100%; margin: 20px 0; }\n';
    
    css += '.tab-buttons { display: flex; gap: 8px; margin-bottom: 15px; border-bottom: 3px solid #b3e5fc; padding-bottom: 10px; flex-wrap: wrap; }\n';
    css += '.tab-btn { padding: 10px 20px; background: linear-gradient(135deg, #4fc3f7 0%, #0288d1 100%); color: white; border: none; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 600; transition: all 0.3s; box-shadow: 0 2px 8px rgba(3, 169, 244, 0.3); }\n';
    css += '.tab-btn:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(3, 169, 244, 0.4); }\n';
    css += '.tab-btn.active { background: linear-gradient(135deg, #0288d1 0%, #01579b 100%); box-shadow: 0 4px 12px rgba(2, 136, 209, 0.5); }\n';
    css += '.excel-tab-content { display: none; }\n';
    
    css += '.excel-search-box { margin-bottom: 15px; padding: 15px; background: linear-gradient(135deg, #e1f5fe 0%, #b3e5fc 100%); border-radius: 10px; display: flex; gap: 10px; align-items: center; box-shadow: 0 2px 8px rgba(3, 169, 244, 0.2); }\n';
    css += '.excel-search-input { flex: 1; padding: 10px 15px; border: 2px solid #4fc3f7; border-radius: 8px; font-size: 14px; background: white; transition: all 0.3s; }\n';
    css += '.excel-search-input:focus { outline: none; border-color: #0288d1; box-shadow: 0 0 0 3px rgba(2, 136, 209, 0.1); }\n';
    css += '.excel-btn { padding: 10px 20px; border: none; border-radius: 8px; cursor: pointer; font-weight: 600; font-size: 14px; transition: all 0.3s; }\n';
    css += '.excel-btn-primary { background: linear-gradient(135deg, #0288d1 0%, #01579b 100%); color: white; }\n';
    css += '.excel-btn-primary:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(2, 136, 209, 0.4); }\n';
    css += '.excel-btn-secondary { background: linear-gradient(135deg, #00acc1 0%, #00838f 100%); color: white; }\n';
    css += '.excel-btn-secondary:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0, 172, 193, 0.4); }\n';
    css += '.excel-search-stats { color: #d32f2f; font-size: 14px; font-weight: 600; }\n';
    
    css += '.excel-filters { margin-bottom: 15px; padding: 15px; background: linear-gradient(135deg, #b3e5fc 0%, #81d4fa 100%); border-radius: 10px; box-shadow: 0 2px 8px rgba(3, 169, 244, 0.2); }\n';
    css += '.excel-filter-section { margin-bottom: 12px; }\n';
    css += '.excel-filter-section:last-child { margin-bottom: 0; }\n';
    css += '.excel-filter-title { color: #01579b; margin-bottom: 8px; font-size: 15px; font-weight: 700; }\n';
    css += '.excel-filter-options { display: flex; flex-wrap: wrap; gap: 15px; }\n';
    css += '.excel-filter-label { cursor: pointer; color: #01579b; font-size: 14px; transition: all 0.2s; display: inline-flex; align-items: center; }\n';
    css += '.excel-filter-label:hover { color: #0288d1; }\n';
    css += '.excel-filter-label input[type="checkbox"] { margin-right: 6px; width: 16px; height: 16px; cursor: pointer; }\n';
    
    css += '.excel-data-table { width: 100%; border-collapse: collapse; margin-bottom: 15px; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }\n';
    css += '.excel-data-table th { background: linear-gradient(135deg, #4fc3f7 0%, #0288d1 100%); color: white; padding: 14px 12px; text-align: left; font-weight: 600; font-size: 14px; }\n';
    css += '.excel-data-table td { padding: 12px; border-bottom: 1px solid #e1f5fe; color: #263238; font-size: 13px; }\n';
    css += '.excel-data-table tr:nth-child(even) { background: #f1f8fb; }\n';
    css += '.excel-data-table tr:hover { background: #e1f5fe; transition: background 0.2s; }\n';
    css += '.excel-data-table tr.hidden { display: none; }\n';
    
    css += '.excel-highlight { background: #fff59d; font-weight: bold; padding: 2px 4px; border-radius: 3px; }\n';
    css += '.excel-current-highlight { background: #ffeb3b; }\n';
    
    css += '.excel-footer { text-align: center; color: #455a64; font-size: 14px; padding: 15px; background: linear-gradient(135deg, #e1f5fe 0%, #b3e5fc 100%); border-radius: 8px; font-weight: 600; }\n';
    css += '.excel-footer span { color: #d32f2f; font-size: 16px; }\n';
    
    return css;
  },
  
  /**
   * JavaScript功能（所有函數加上excel前綴避免衝突）
   */
  getScripts: function(allSheets) {
    let js = '';
    js += 'var excelCurrentTab = 0;\n';
    js += '\n';
    js += 'function switchExcelTab(tabIndex) {\n';
    js += '  var contents = document.querySelectorAll(".excel-tab-content");\n';
    js += '  for (var i = 0; i < contents.length; i++) {\n';
    js += '    contents[i].style.display = "none";\n';
    js += '  }\n';
    js += '  var buttons = document.querySelectorAll(".tab-btn");\n';
    js += '  for (var i = 0; i < buttons.length; i++) {\n';
    js += '    buttons[i].classList.remove("active");\n';
    js += '  }\n';
    js += '  document.getElementById("exceltab" + tabIndex).style.display = "block";\n';
    js += '  buttons[tabIndex].classList.add("active");\n';
    js += '  excelCurrentTab = tabIndex;\n';
    js += '}\n';
    js += '\n';
    js += 'function excelSearch(tabIndex) {\n';
    js += '  var keyword = document.getElementById("excelSearchInput" + tabIndex).value.trim();\n';
    js += '  if (!keyword) return;\n';
    js += '  var table = document.getElementById("excelDataTable" + tabIndex);\n';
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
    js += '        var regex = new RegExp("(" + excelEscapeRegex(keyword) + ")", "gi");\n';
    js += '        cell.innerHTML = text.replace(regex, "<span class=\\"excel-highlight\\">$1</span>");\n';
    js += '      }\n';
    js += '    }\n';
    js += '    if (hasMatch) { row.classList.remove("hidden"); } else { row.classList.add("hidden"); }\n';
    js += '  }\n';
    js += '  document.getElementById("excelSearchStats" + tabIndex).textContent = "找到 " + matchCount + " 個符合項目";\n';
    js += '  excelUpdateCount(tabIndex);\n';
    js += '}\n';
    js += '\n';
    js += 'function excelClearSearch(tabIndex) {\n';
    js += '  document.getElementById("excelSearchInput" + tabIndex).value = "";\n';
    js += '  document.getElementById("excelSearchStats" + tabIndex).textContent = "";\n';
    js += '  var table = document.getElementById("excelDataTable" + tabIndex);\n';
    js += '  var rows = table.querySelectorAll("tbody tr");\n';
    js += '  for (var i = 0; i < rows.length; i++) {\n';
    js += '    var cells = rows[i].querySelectorAll("td");\n';
    js += '    for (var j = 0; j < cells.length; j++) {\n';
    js += '      cells[j].innerHTML = cells[j].textContent;\n';
    js += '    }\n';
    js += '  }\n';
    js += '  excelApplyFilters(tabIndex);\n';
    js += '}\n';
    js += '\n';
    js += 'function excelApplyFilters(tabIndex) {\n';
    js += '  var table = document.getElementById("excelDataTable" + tabIndex);\n';
    js += '  var rows = table.querySelectorAll("tbody tr");\n';
    js += '  var districtFilters = [];\n';
    js += '  var districtCheckboxes = document.querySelectorAll(".excel-district-filter[data-tab=\\"" + tabIndex + "\\"]:checked");\n';
    js += '  for (var i = 0; i < districtCheckboxes.length; i++) { districtFilters.push(districtCheckboxes[i].value); }\n';
    js += '  var columnFilters = [];\n';
    js += '  var columnCheckboxes = document.querySelectorAll(".excel-column-filter[data-tab=\\"" + tabIndex + "\\"]:checked");\n';
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
    js += '  excelUpdateCount(tabIndex);\n';
    js += '}\n';
    js += '\n';
    js += 'function excelUpdateCount(tabIndex) {\n';
    js += '  var table = document.getElementById("excelDataTable" + tabIndex);\n';
    js += '  var visibleRows = table.querySelectorAll("tbody tr:not(.hidden)");\n';
    js += '  document.getElementById("excelCount" + tabIndex).textContent = visibleRows.length;\n';
    js += '}\n';
    js += '\n';
    js += 'function excelEscapeRegex(str) {\n';
    js += '  return str.replace(/[.*+?^${}()|[\\]\\\\]/g, "\\\\$&");\n';
    js += '}\n';
    js += '\n';
    js += 'if (typeof excelInitialized === "undefined") {\n';
    js += '  excelInitialized = true;\n';
    js += '  if (document.readyState === "loading") {\n';
    js += '    document.addEventListener("DOMContentLoaded", function() { switchExcelTab(0); });\n';
    js += '  } else {\n';
    js += '    switchExcelTab(0);\n';
    js += '  }\n';
    js += '}\n';
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
