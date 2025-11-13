/**
 * 資料解析器 - Parser.gs
 * 
 * 功能：解析不同格式的Excel分頁資料
 */

/**
 * 清理字符串,移除可能導致JSON錯誤的特殊字符
 * @param {String} str - 要清理的字符串
 * @returns {String} 清理後的字符串
 */
function cleanString(str) {
  if (!str && str !== 0) return '';
  
  // 轉換為字符串
  str = String(str);
  
  // 移除控制字符
  str = str.replace(/[\x00-\x1F\x7F-\x9F]/g, '');
  
  // 移除零寬空格等不可見字符
  str = str.replace(/[\u200B-\u200D\uFEFF]/g, '');
  
  // 移除所有雙引號(避免JSON錯誤)
  str = str.replace(/"/g, '');
  
  // 規範化空白字符
  str = str.replace(/\s+/g, ' ').trim();
  
  return str;
}

// ==================== 模式A解析器 (簡單格式) ====================

/**
 * 解析模式A格式（8欄，單層標題）
 * 適用於：住宿式、居家式、巷弄長照站
 * 
 * @param {Array} values - 分頁資料陣列
 * @param {Object} formatType - 格式類型資訊
 * @returns {Object} 解析後的資料
 */
function parseModeA(values, formatType) {
  const result = {
    title: '',
    updateDate: '',
    totalCount: 0,
    headers: [],
    contractCodes: [],
    data: [],
    dataCount: 0
  };
  
  // 解析標題列（第1列）
  const titleText = values[formatType.titleRow].join('');
  result.title = titleText.trim();
  
  // 提取更新日期
  const dateMatch = titleText.match(/(\d{3})年(\d{1,2})月(\d{1,2})日/);
  if (dateMatch) {
    result.updateDate = `${dateMatch[1]}/${dateMatch[2]}/${dateMatch[3]}`;
  }
  
  // 提取總數
  const countMatch = titleText.match(/特約共(\d+)家/);
  if (countMatch) {
    result.totalCount = parseInt(countMatch[1]);
  }
  
  // 解析欄位標題（第2列）
  const headerRow = values[formatType.headerRow];
  result.headers = headerRow.filter(function(cell) { return cell !== ''; }).map(function(cell) { return String(cell).trim(); });
  
  // 識別特約碼別欄位（通常是最後1-2欄）
  const contractCodeStartIdx = 7; // 第8欄開始
  for (let i = contractCodeStartIdx; i < result.headers.length; i++) {
    const header = result.headers[i];
    // 提取特約碼別代碼（如SC05, GA09等）
    const codeMatch = header.match(/\(([A-Z]{2}\d{2})\)/);
    if (codeMatch) {
      result.contractCodes.push({
        index: i,
        code: codeMatch[1],
        name: header
      });
    }
  }
  
  // 解析資料列
  for (let i = formatType.dataStartRow; i < values.length; i++) {
    const row = values[i];
    
    // 跳過空列
    if (row.every(function(cell) { return !cell; })) {
      continue;
    }
    
    const rowData = {
      序號: cleanString(row[0]),
      機構名稱: cleanString(row[1]),
      服務區別: cleanString(row[2]),
      郵遞區號: cleanString(row[3]),
      機構地址: cleanString(row[4]),
      聯絡電話: cleanString(row[5]),
      聯絡窗口: cleanString(row[6]),
      特約碼別: {}
    };
    
    // 處理特約碼別
    result.contractCodes.forEach(function(code) {
      const value = row[code.index];
      rowData.特約碼別[code.code] = (value === 'V' || value === 'v' || value === '✓');
    });
    
    result.data.push(rowData);
  }
  
  result.dataCount = result.data.length;
  
  Logger.log('模式A解析完成: ' + result.dataCount + ' 筆資料');
  
  return result;
}

// ==================== 模式B解析器 (複雜格式) ====================

/**
 * 解析模式B格式（13-15欄，雙層標題）
 * 適用於：專業服務、社區式
 * 
 * @param {Array} values - 分頁資料陣列
 * @param {Object} formatType - 格式類型資訊
 * @returns {Object} 解析後的資料
 */
function parseModeB(values, formatType) {
  const result = {
    title: '',
    updateDate: '',
    totalCount: 0,
    headers: [],
    contractCodes: [],
    contractGroups: [],
    data: [],
    dataCount: 0
  };
  
  // 解析標題列（第1列）
  const titleText = values[formatType.titleRow].join('');
  result.title = titleText.trim();
  
  // 提取更新日期
  const dateMatch = titleText.match(/(\d{3})年(\d{1,2})月(\d{1,2})日/);
  if (dateMatch) {
    result.updateDate = `${dateMatch[1]}/${dateMatch[2]}/${dateMatch[3]}`;
  }
  
  // 提取總數
  const countMatch = titleText.match(/特約共(\d+)家/);
  if (countMatch) {
    result.totalCount = parseInt(countMatch[1]);
  }
  
  // 解析雙層標題
  const headerRow1 = values[formatType.headerRow2]; // 第2列
  const headerRow2 = values[formatType.headerRow];  // 第3列
  
  // 基本欄位（前7欄）
  result.headers = headerRow2.slice(0, 7).filter(function(cell) { return cell !== ''; }).map(function(cell) { return String(cell).trim(); });
  
  // 解析特約碼別分組
  let currentGroup = null;
  const contractCodeStartIdx = 7;
  
  for (let i = contractCodeStartIdx; i < headerRow2.length; i++) {
    const cell1 = String(headerRow1[i] || '').trim();
    const cell2 = String(headerRow2[i] || '').trim();
    
    // 檢查是否是新的分組
    if (cell1 && cell1 !== '') {
      currentGroup = {
        name: cell1,
        codes: []
      };
      result.contractGroups.push(currentGroup);
    }
    
    // 提取特約碼別代碼
    if (cell2) {
      const code = cell2;
      const codeInfo = {
        index: i,
        code: code,
        group: currentGroup ? currentGroup.name : '未分組'
      };
      
      result.contractCodes.push(codeInfo);
      if (currentGroup) {
        currentGroup.codes.push(code);
      }
    }
  }
  
  // 解析資料列
  for (let i = formatType.dataStartRow; i < values.length; i++) {
    const row = values[i];
    
    // 跳過空列
    if (row.every(function(cell) { return !cell; })) {
      continue;
    }
    
    const rowData = {
      序號: cleanString(row[0]),
      機構名稱: cleanString(row[1]),
      服務區別: cleanString(row[2]),
      郵遞區號: cleanString(row[3]),
      機構地址: cleanString(row[4]),
      聯絡電話: cleanString(row[5]),
      聯絡窗口: cleanString(row[6]),
      特約碼別: {}
    };
    
    // 處理特約碼別
    result.contractCodes.forEach(function(code) {
      const value = row[code.index];
      rowData.特約碼別[code.code] = (value === 'V' || value === 'v' || value === '✓');
    });
    
    result.data.push(rowData);
  }
  
  result.dataCount = result.data.length;
  
  Logger.log('模式B解析完成: ' + result.dataCount + ' 筆資料');
  
  return result;
}

// ==================== 輔助函數 ====================

/**
 * 提取行政區列表
 * @param {Array} data - 資料陣列
 * @returns {Array} 唯一的行政區列表
 */
function extractDistricts(data) {
  const districts = new Set();
  
  data.forEach(function(row) {
    const districtText = row.服務區別 || '';
    const districtArray = districtText.split(/[、,，\n]/);
    
    districtArray.forEach(function(district) {
      const cleaned = district.trim();
      if (cleaned && cleaned !== '全區') {
        districts.add(cleaned);
      }
    });
  });
  
  const result = Array.from(districts).sort();
  result.unshift('全區'); // 在最前面加入「全區」選項
  
  return result;
}

/**
 * 提取所有特約碼別
 * @param {Array} sheets - 所有分頁資料
 * @returns {Object} 特約碼別對應表
 */
function extractAllContractCodes(sheets) {
  const codeMap = {};
  
  sheets.forEach(function(sheet) {
    sheet.contractCodes.forEach(function(code) {
      if (!codeMap[code.code]) {
        codeMap[code.code] = {
          code: code.code,
          name: code.name || code.code,
          group: code.group || '未分組'
        };
      }
    });
  });
  
  return codeMap;
}