/**
 * HTML生成器 - HTMLGenerator.gs
 * 
 * 功能：生成整合的HTML頁面，包含分頁切換、搜尋、篩選功能
 */

/**
 * 生成整合HTML
 * @param {Array} sheets - 所有分頁資料
 * @returns {String} 完整的HTML原始碼
 */
function generateIntegratedHTML(sheets) {
  // 提取所有行政區和特約碼別
  const allDistricts = extractAllDistricts(sheets);
  const allContractCodes = extractAllContractCodes(sheets);
  
  // 生成HTML
  let html = generateHTMLHeader();
  html += generateHTMLStyles();
  html += generateHTMLBody(sheets, allDistricts, allContractCodes);
  html += generateHTMLScripts(sheets);
  html += '</html>';
  
  // 壓縮HTML（移除不必要的空白和換行）
  html = compressHTML(html);
  
  Logger.log('HTML大小: ' + html.length + ' 字元');
  
  return html;
}

/**
 * 壓縮HTML（移除多餘空白，保留功能）
 */
function compressHTML(html) {
  // 移除HTML註解（但保留條件註解）
  html = html.replace(/<!--(?!\[if)[\s\S]*?-->/g, '');
  
  // 移除多餘的空白行（連續的空行合併為一行）
  html = html.replace(/\n\s*\n\s*\n/g, '\n\n');
  
  // 不進行激進的壓縮，保持基本結構
  // 只移除行首和行尾的空白
  const lines = html.split('\n');
  const compressed = lines.map(function(line) {
    return line.trim();
  }).filter(function(line) {
    return line.length > 0; // 移除完全空白的行
  }).join('\n');
  
  Logger.log('壓縮前: ' + html.length + ' 字元, 壓縮後: ' + compressed.length + ' 字元');
  Logger.log('壓縮率: ' + ((1 - compressed.length / html.length) * 100).toFixed(1) + '%');
  
  return compressed;
}

/**
 * 提取所有行政區
 */
function extractAllDistricts(sheets) {
  const districts = new Set();
  
  sheets.forEach(function(sheet) {
    sheet.data.forEach(function(row) {
      const districtText = row.服務區別 || '';
      const districtArray = districtText.split(/[、,，\n]/);
      
      districtArray.forEach(function(district) {
        const cleaned = district.trim();
        if (cleaned && cleaned !== '全區') {
          districts.add(cleaned);
        }
      });
    });
  });
  
  const result = Array.from(districts).sort();
  return result;
}

/**
 * 生成HTML Header
 */
function generateHTMLHeader() {
  return `<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>臺北市政府衛生局長照2.0特約服務單位一覽表</title>
`;
}

/**
 * 生成HTML樣式
 */
function generateHTMLStyles() {
  return `  <style>
    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
    }
    
    body {
      font-family: 'Microsoft JhengHei', 'PingFang TC', sans-serif;
      background-color: #f5f7fa;
      color: #333;
      line-height: 1.6;
    }
    
    .container {
      max-width: 1400px;
      margin: 0 auto;
      padding: 20px;
    }
    
    /* 標題區 */
    .header {
      background: linear-gradient(135deg, #1a73e8 0%, #4285f4 100%);
      color: white;
      padding: 30px;
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(26, 115, 232, 0.3);
      margin-bottom: 20px;
    }
    
    .header h1 {
      font-size: 28px;
      margin-bottom: 10px;
      font-weight: 600;
    }
    
    .header p {
      font-size: 14px;
      opacity: 0.9;
    }
    
    /* 控制面板 */
    .control-panel {
      background: white;
      padding: 20px;
      border-radius: 8px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
      margin-bottom: 20px;
    }
    
    .search-bar {
      display: flex;
      gap: 10px;
      margin-bottom: 15px;
      flex-wrap: wrap;
    }
    
    .search-input {
      flex: 1;
      min-width: 200px;
      padding: 10px 15px;
      border: 2px solid #e0e0e0;
      border-radius: 4px;
      font-size: 14px;
      transition: border-color 0.3s;
    }
    
    .search-input:focus {
      outline: none;
      border-color: #1a73e8;
    }
    
    .btn {
      padding: 10px 20px;
      border: none;
      border-radius: 4px;
      font-size: 14px;
      cursor: pointer;
      transition: all 0.3s;
      font-weight: 500;
    }
    
    .btn-primary {
      background: #1a73e8;
      color: white;
    }
    
    .btn-primary:hover {
      background: #1557b0;
      box-shadow: 0 2px 4px rgba(26, 115, 232, 0.4);
    }
    
    .btn-secondary {
      background: #f1f3f4;
      color: #5f6368;
    }
    
    .btn-secondary:hover {
      background: #e8eaed;
    }
    
    /* 篩選器 */
    .filters {
      display: flex;
      gap: 15px;
      flex-wrap: wrap;
      align-items: center;
    }
    
    .filter-group {
      display: flex;
      align-items: center;
      gap: 8px;
    }
    
    .filter-group label {
      font-size: 14px;
      color: #5f6368;
      font-weight: 500;
    }
    
    .filter-select {
      padding: 8px 12px;
      border: 1px solid #dadce0;
      border-radius: 4px;
      font-size: 14px;
      background: white;
      cursor: pointer;
    }
    
    /* 特約碼別篩選 */
    .contract-filters {
      display: flex;
      gap: 15px;
      flex-wrap: wrap;
      margin-top: 10px;
      padding-top: 10px;
      border-top: 1px solid #e0e0e0;
    }
    
    .contract-filter-item {
      display: flex;
      align-items: center;
      gap: 5px;
    }
    
    .contract-filter-item input[type="checkbox"] {
      width: 16px;
      height: 16px;
      cursor: pointer;
    }
    
    .contract-filter-item label {
      font-size: 13px;
      cursor: pointer;
      color: #5f6368;
    }
    
    /* 統計資訊 */
    .stats {
      display: flex;
      gap: 10px;
      margin-top: 10px;
      font-size: 13px;
      color: #5f6368;
    }
    
    .stat-item {
      padding: 5px 10px;
      background: #e8f0fe;
      border-radius: 4px;
    }
    
    /* 分頁標籤 */
    .tabs {
      display: flex;
      gap: 5px;
      margin-bottom: 20px;
      overflow-x: auto;
      background: white;
      padding: 10px;
      border-radius: 8px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    .tab {
      padding: 12px 24px;
      background: #f1f3f4;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 500;
      color: #5f6368;
      transition: all 0.3s;
      white-space: nowrap;
    }
    
    .tab:hover {
      background: #e8eaed;
    }
    
    .tab.active {
      background: #1a73e8;
      color: white;
      box-shadow: 0 2px 4px rgba(26, 115, 232, 0.3);
    }
    
    /* 表格容器 */
    .table-container {
      background: white;
      border-radius: 8px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
      overflow: hidden;
    }
    
    .sheet-content {
      display: none;
    }
    
    .sheet-content.active {
      display: block;
    }
    
    /* 表格標題 */
    .table-title {
      background: #f8f9fa;
      padding: 15px 20px;
      border-bottom: 2px solid #1a73e8;
    }
    
    .table-title h2 {
      font-size: 18px;
      color: #1a73e8;
      margin-bottom: 5px;
    }
    
    .table-title p {
      font-size: 13px;
      color: #5f6368;
    }
    
    /* 表格 */
    .data-table {
      width: 100%;
      border-collapse: collapse;
    }
    
    .data-table thead {
      background: #1a73e8;
      color: white;
      position: sticky;
      top: 0;
      z-index: 10;
    }
    
    .data-table th {
      padding: 12px 8px;
      text-align: left;
      font-size: 13px;
      font-weight: 600;
      border-right: 1px solid rgba(255,255,255,0.2);
    }
    
    .data-table th:last-child {
      border-right: none;
    }
    
    .data-table tbody tr {
      border-bottom: 1px solid #e0e0e0;
      transition: background-color 0.2s;
    }
    
    .data-table tbody tr:hover {
      background-color: #f8f9fa;
    }
    
    .data-table tbody tr.hidden {
      display: none;
    }
    
    .data-table td {
      padding: 10px 8px;
      font-size: 13px;
      border-right: 1px solid #f0f0f0;
    }
    
    .data-table td:last-child {
      border-right: none;
    }
    
    /* 特約碼別欄位 */
    .contract-code {
      text-align: center;
      font-size: 16px;
      color: #34a853;
    }
    
    /* 搜尋高亮 */
    .highlight {
      background-color: #fff59d;
      padding: 2px 4px;
      border-radius: 2px;
      font-weight: 600;
    }
    
    /* 無資料提示 */
    .no-data {
      padding: 40px;
      text-align: center;
      color: #5f6368;
      font-size: 14px;
    }
    
    /* 載入動畫 */
    .loading {
      display: none;
      text-align: center;
      padding: 20px;
      color: #5f6368;
    }
    
    /* 響應式設計 */
    @media (max-width: 768px) {
      .container {
        padding: 10px;
      }
      
      .header h1 {
        font-size: 22px;
      }
      
      .search-bar {
        flex-direction: column;
      }
      
      .search-input {
        width: 100%;
      }
      
      .filters {
        flex-direction: column;
        align-items: stretch;
      }
      
      .filter-group {
        flex-direction: column;
        align-items: stretch;
      }
      
      .tabs {
        flex-wrap: wrap;
      }
      
      .data-table {
        font-size: 12px;
      }
      
      .data-table th,
      .data-table td {
        padding: 8px 4px;
      }
    }
  </style>
`;
}

/**
 * 生成HTML Body
 */
function generateHTMLBody(sheets, allDistricts, allContractCodes) {
  let html = `</head>
<body>
  <div class="container">
    <!-- 標題區 -->
    <div class="header">
      <h1>臺北市政府衛生局長照2.0特約服務單位一覽表</h1>
      <p>資料更新日期：${sheets[0].updateDate || ''} | 共 ${sheets.length} 個服務類別</p>
    </div>
    
    <!-- 控制面板 -->
    <div class="control-panel">
      <!-- 搜尋列 -->
      <div class="search-bar">
        <input type="text" id="searchInput" class="search-input" placeholder="搜尋機構名稱或地址..." />
        <button class="btn btn-primary" onclick="performSearch()">🔍 搜尋</button>
        <button class="btn btn-secondary" onclick="clearSearch()">✕ 清除</button>
      </div>
      
      <!-- 篩選器 -->
      <div class="filters">
        <div class="filter-group">
          <label for="districtFilter">行政區：</label>
          <select id="districtFilter" class="filter-select" onchange="applyFilters()">
            <option value="">全部區域</option>
`;
  
  // 加入行政區選項
  allDistricts.forEach(function(district) {
    html += `            <option value="${district}">${district}</option>\n`;
  });
  
  html += `          </select>
        </div>
      </div>
      
      <!-- 特約碼別篩選 -->
      <div class="contract-filters" id="contractFilters">
        <label style="font-weight: 600; color: #5f6368;">特約碼別：</label>
      </div>
      
      <!-- 統計資訊 -->
      <div class="stats" id="stats">
        <span class="stat-item">總機構數：<strong id="totalCount">0</strong></span>
        <span class="stat-item">顯示機構數：<strong id="displayCount">0</strong></span>
      </div>
    </div>
    
    <!-- 分頁標籤 -->
    <div class="tabs" id="tabs">
`;
  
  // 生成分頁標籤
  sheets.forEach(function(sheet, index) {
    const tabName = getSimpleTabName(sheet.sheetName);
    const activeClass = index === 0 ? ' active' : '';
    html += `      <button class="tab${activeClass}" onclick="switchTab(${index})">${tabName}</button>\n`;
  });
  
  html += `    </div>
    
    <!-- 表格容器 -->
    <div class="table-container">
`;
  
  // 生成各分頁內容
  sheets.forEach(function(sheet, index) {
    html += generateSheetContent(sheet, index);
  });
  
  html += `    </div>
  </div>
  
`;
  
  return html;
}

/**
 * 生成分頁內容
 */
function generateSheetContent(sheet, index) {
  const activeClass = index === 0 ? ' active' : '';
  
  let html = `      <div class="sheet-content${activeClass}" id="sheet${index}">
        <div class="table-title">
          <h2>${sheet.title}</h2>
          <p>資料筆數：${sheet.dataCount} 筆</p>
        </div>
        <div style="overflow-x: auto;">
          <table class="data-table" id="table${index}">
            <thead>
              <tr>
                <th style="min-width: 50px;">序號</th>
                <th style="min-width: 200px;">機構名稱</th>
                <th style="min-width: 120px;">服務區別</th>
                <th style="min-width: 80px;">郵遞區號</th>
                <th style="min-width: 250px;">機構地址</th>
                <th style="min-width: 120px;">聯絡電話</th>
                <th style="min-width: 80px;">聯絡窗口</th>
`;
  
  // 特約碼別欄位標題
  sheet.contractCodes.forEach(function(code) {
    html += `                <th style="min-width: 60px; text-align: center;">${code.code}</th>\n`;
  });
  
  html += `              </tr>
            </thead>
            <tbody>
`;
  
  // 資料列
  sheet.data.forEach(function(row, rowIndex) {
    html += `              <tr data-row="${rowIndex}">\n`;
    html += `                <td>${escapeHtml(row.序號)}</td>\n`;
    html += `                <td>${escapeHtml(row.機構名稱)}</td>\n`;
    html += `                <td>${escapeHtml(row.服務區別)}</td>\n`;
    html += `                <td>${escapeHtml(row.郵遞區號)}</td>\n`;
    html += `                <td>${escapeHtml(row.機構地址)}</td>\n`;
    html += `                <td>${escapeHtml(row.聯絡電話)}</td>\n`;
    html += `                <td>${escapeHtml(row.聯絡窗口)}</td>\n`;
    
    // 特約碼別
    sheet.contractCodes.forEach(function(code) {
      const hasContract = row.特約碼別[code.code];
      const checkMark = hasContract ? '✓' : '';
      html += `                <td class="contract-code">${checkMark}</td>\n`;
    });
    
    html += `              </tr>\n`;
  });
  
  html += `            </tbody>
          </table>
        </div>
      </div>
`;
  
  return html;
}

/**
 * 生成JavaScript腳本
 */
function generateHTMLScripts(sheets) {
  // 將sheets資料轉換為JSON字串
  const sheetsJSON = JSON.stringify(sheets.map(function(sheet) {
    return {
      name: sheet.sheetName,
      title: sheet.title,
      dataCount: sheet.dataCount,
      contractCodes: sheet.contractCodes,
      data: sheet.data
    };
  }));
  
  // 使用Base64編碼避免複製時的引號轉義問題
  const sheetsDataBase64 = Utilities.base64Encode(sheetsJSON, Utilities.Charset.UTF_8);
  
  let html = `  <script>
    // 資料 (Base64編碼，避免複製時引號問題)
    const sheetsDataBase64 = '${sheetsDataBase64}';
    
    // 解碼資料
    function base64Decode(str) {
      try {
        return decodeURIComponent(atob(str).split('').map(function(c) {
          return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
        }).join(''));
      } catch(e) {
        console.error('Base64解碼失敗:', e);
        return null;
      }
    }
    
    const sheetsData = JSON.parse(base64Decode(sheetsDataBase64));
    let currentTab = 0;
    
    // 初始化 - 使用多重保險機制
    // 方案1: DOMContentLoaded (DOM解析完成時觸發,最早)
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', function() {
        console.log('DOMContentLoaded 觸發');
        initializeApp();
      });
    } else {
      // DOM已經載入完成
      console.log('DOM已就緒,直接初始化');
      initializeApp();
    }
    
    // 方案2: window.onload (所有資源載入完成,較晚)
    window.onload = function() {
      console.log('window.onload 觸發');
      // 如果之前沒初始化成功,再試一次
      setTimeout(initializeApp, 100);
    };
    
    // 統一的初始化函數
    function initializeApp() {
      console.log('開始初始化應用...');
      
      try {
        initContractFilters();
        console.log('✓ 特約碼別篩選器初始化完成');
      } catch(e) {
        console.error('❌ 初始化特約碼別篩選失敗:', e);
        console.error('錯誤堆疊:', e.stack);
      }
      
      try {
        updateStats();
        console.log('✓ 統計資訊更新完成');
      } catch(e) {
        console.error('❌ 更新統計失敗:', e);
      }
    }
    
    // 初始化特約碼別篩選器
    function initContractFilters() {
      console.log('正在查找 contractFilters 容器...');
      
      const container = document.getElementById('contractFilters');
      
      // 安全檢查:如果找不到容器,記錄詳細資訊並返回
      if (!container) {
        console.error('❌ 找不到 contractFilters 容器元素');
        console.log('document.readyState:', document.readyState);
        console.log('所有ID元素:', Array.from(document.querySelectorAll('[id]')).map(function(el) { return el.id; }));
        
        // 嘗試查找所有可能的容器
        var allDivs = document.querySelectorAll('div');
        console.log('頁面共有', allDivs.length, '個div元素');
        
        var contractRelated = document.querySelectorAll('[class*="contract"]');
        console.log('包含contract的元素:', contractRelated.length);
        
        return;
      }
      
      console.log('✓ 找到 contractFilters 容器');
      
      const codes = new Set();
      
      // 檢查 sheetsData 是否存在
      if (!sheetsData || sheetsData.length === 0) {
        console.warn('sheetsData 為空或不存在');
        return;
      }
      
      console.log('sheetsData 包含', sheetsData.length, '個分頁');
      
      sheetsData.forEach(function(sheet, idx) {
        console.log('處理分頁', idx + 1, ':', sheet.name);
        if (sheet.contractCodes && sheet.contractCodes.length > 0) {
          console.log('  - 找到', sheet.contractCodes.length, '個特約碼別');
          sheet.contractCodes.forEach(function(code) {
            codes.add(code.code);
          });
        } else {
          console.log('  - 此分頁沒有特約碼別');
        }
      });
      
      const sortedCodes = Array.from(codes).sort();
      console.log('總共', sortedCodes.length, '個唯一特約碼別:', sortedCodes);
      
      if (sortedCodes.length === 0) {
        console.warn('沒有找到任何特約碼別');
        return;
      }
      
      // 清空容器(保留label)
      const existingLabel = container.querySelector('label');
      container.innerHTML = '';
      if (existingLabel) {
        container.appendChild(existingLabel);
        console.log('✓ 保留了原有的label');
      }
      
      // 創建checkbox
      sortedCodes.forEach(function(code) {
        const div = document.createElement('div');
        div.className = 'contract-filter-item';
        div.innerHTML = \`
          <input type="checkbox" id="contract_\${code}" value="\${code}" onchange="applyFilters()">
          <label for="contract_\${code}">\${code}</label>
        \`;
        container.appendChild(div);
      });
      
      console.log('✓ 成功創建', sortedCodes.length, '個特約碼別篩選器');
    }
    
    // 切換分頁
    function switchTab(index) {
      currentTab = index;
      
      // 更新分頁標籤
      const tabs = document.querySelectorAll('.tab');
      tabs.forEach(function(tab, i) {
        tab.classList.toggle('active', i === index);
      });
      
      // 更新內容
      const contents = document.querySelectorAll('.sheet-content');
      contents.forEach(function(content, i) {
        content.classList.toggle('active', i === index);
      });
      
      // 重新應用篩選
      applyFilters();
    }
    
    // 執行搜尋
    function performSearch() {
      applyFilters();
    }
    
    // 清除搜尋
    function clearSearch() {
      const searchInput = document.getElementById('searchInput');
      const districtFilter = document.getElementById('districtFilter');
      
      if (searchInput) searchInput.value = '';
      if (districtFilter) districtFilter.value = '';
      
      // 清除所有特約碼別勾選
      document.querySelectorAll('.contract-filter-item input').forEach(function(cb) {
        cb.checked = false;
      });
      
      applyFilters();
    }
    
    // 應用篩選
    function applyFilters() {
      const searchInput = document.getElementById('searchInput');
      const districtFilter = document.getElementById('districtFilter');
      
      const searchText = searchInput ? searchInput.value.toLowerCase().trim() : '';
      const districtValue = districtFilter ? districtFilter.value : '';
      
      // 取得勾選的特約碼別
      const selectedCodes = [];
      document.querySelectorAll('.contract-filter-item input:checked').forEach(function(cb) {
        selectedCodes.push(cb.value);
      });
      
      const table = document.getElementById('table' + currentTab);
      if (!table) return;
      
      const rows = table.querySelectorAll('tbody tr');
      let displayCount = 0;
      
      rows.forEach(function(row, index) {
        const rowData = sheetsData[currentTab].data[index];
        let show = true;
        
        // 關鍵字搜尋
        if (searchText) {
          const name = (rowData['機構名稱'] || '').toLowerCase();
          const address = (rowData['機構地址'] || '').toLowerCase();
          if (!name.includes(searchText) && !address.includes(searchText)) {
            show = false;
          }
        }
        
        // 行政區篩選
        if (districtValue && show) {
          const districts = (rowData['服務區別'] || '').split(/[、,，\\n]/);
          const hasDistrict = districts.some(function(d) { return d.trim() === districtValue; });
          if (!hasDistrict) {
            show = false;
          }
        }
        
        // 特約碼別篩選
        if (selectedCodes.length > 0 && show) {
          const hasAnyCode = selectedCodes.some(function(code) { return rowData['特約碼別'][code]; });
          if (!hasAnyCode) {
            show = false;
          }
        }
        
        // 顯示/隱藏列
        row.classList.toggle('hidden', !show);
        
        if (show) {
          displayCount++;
          highlightText(row, searchText);
        } else {
          removeHighlight(row);
        }
      });
      
      updateStats(displayCount);
    }
    
    // 高亮文字
    function highlightText(row, searchText) {
      if (!searchText) {
        removeHighlight(row);
        return;
      }
      
      const cells = row.querySelectorAll('td');
      cells.forEach(function(cell, index) {
        if (index < 2 || index === 4) { // 機構名稱或地址
          const originalText = sheetsData[currentTab].data[parseInt(row.dataset.row)];
          let text = '';
          
          if (index === 1) text = originalText['機構名稱'] || '';
          else if (index === 4) text = originalText['機構地址'] || '';
          
          if (text) {
            const regex = new RegExp('(' + escapeRegex(searchText) + ')', 'gi');
            const highlightedText = text.replace(regex, '<span class="highlight">$1</span>');
            cell.innerHTML = highlightedText;
          }
        }
      });
    }
    
    // 移除高亮
    function removeHighlight(row) {
      const cells = row.querySelectorAll('td');
      cells.forEach(function(cell, index) {
        if (index < 2 || index === 4) {
          const originalText = sheetsData[currentTab].data[parseInt(row.dataset.row)];
          let text = '';
          
          if (index === 1) text = originalText['機構名稱'] || '';
          else if (index === 4) text = originalText['機構地址'] || '';
          
          if (text) {
            cell.textContent = text;
          }
        }
      });
    }
    
    // 更新統計
    function updateStats(displayCount) {
      const totalCount = sheetsData[currentTab].dataCount;
      const totalElement = document.getElementById('totalCount');
      const displayElement = document.getElementById('displayCount');
      
      if (totalElement) {
        totalElement.textContent = totalCount;
      }
      if (displayElement) {
        displayElement.textContent = displayCount !== undefined ? displayCount : totalCount;
      }
    }
    
    // 轉義正則表達式特殊字符
    function escapeRegex(str) {
      var specials = ['.', '*', '+', '?', '^', '$', '{', '}', '(', ')', '|', '[', ']', '\\\\'];
      for (var i = 0; i < specials.length; i++) {
        str = str.split(specials[i]).join('\\\\' + specials[i]);
      }
      return str;
    }
  </script>
`;
  
  return html;
}

/**
 * 輔助函數：轉義HTML特殊字符
 */
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
    .replace(/\n/g, '<br>');
}

/**
 * 輔助函數：簡化分頁名稱
 */
function getSimpleTabName(sheetName) {
  if (sheetName.includes('專業')) return '專業服務';
  if (sheetName.includes('住宿')) return '住宿式';
  if (sheetName.includes('社區')) return '社區式';
  if (sheetName.includes('居家')) return '居家式';
  if (sheetName.includes('巷弄')) return '巷弄長照站';
  return sheetName;
}