/**
 * 解析器工廠
 */
const ParserFactory = {
  
  /**
   * 取得適當的解析器
   */
  getParser: function(sheetName, worksheet) {
    // 嘗試懲戒解析器
    const disciplineParser = new DisciplineParser();
    if (disciplineParser.canHandle(sheetName, worksheet)) {
      return disciplineParser;
    }
    
    // 嘗試特約單位解析器
    const contractParser = new ContractUnitParser();
    if (contractParser.canHandle(sheetName, worksheet)) {
      return contractParser;
    }
    
    // 使用預設解析器
    return new DefaultParser();
  }
};

/**
 * 基礎解析器類別
 */
function BaseParser() {
  this.typeName = 'Unknown';
}

BaseParser.prototype.getTypeName = function() {
  return this.typeName;
};

BaseParser.prototype.canHandle = function(sheetName, worksheet) {
  return false;
};

BaseParser.prototype.parse = function(worksheet, sheetName) {
  throw new Error('parse() must be implemented');
};

BaseParser.prototype.worksheetToArray = function(worksheet) {
  return worksheet.data;
};

/**
 * 預設解析器
 */
function DefaultParser() {
  BaseParser.call(this);
  this.typeName = 'Default';
}

DefaultParser.prototype = Object.create(BaseParser.prototype);
DefaultParser.prototype.constructor = DefaultParser;

DefaultParser.prototype.canHandle = function(sheetName, worksheet) {
  return true;
};

DefaultParser.prototype.parse = function(worksheet, sheetName) {
  const data = this.worksheetToArray(worksheet);
  
  if (data.length === 0) {
    return { title: sheetName, headers: [], data: [] };
  }
  
  // 假設第一列是標題
  const headers = data[0];
  const rows = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // 過濾完全空白的列
    let hasContent = false;
    for (let j = 0; j < row.length; j++) {
      if (row[j] !== null && row[j] !== undefined && row[j] !== '') {
        hasContent = true;
        break;
      }
    }
    if (hasContent) {
      rows.push(row);
    }
  }
  
  return {
    title: sheetName,
    headers: headers,
    data: rows
  };
};

/**
 * 懲戒格式解析器
 */
function DisciplineParser() {
  BaseParser.call(this);
  this.typeName = 'Discipline';
}

DisciplineParser.prototype = Object.create(BaseParser.prototype);
DisciplineParser.prototype.constructor = DisciplineParser;

DisciplineParser.prototype.canHandle = function(sheetName, worksheet) {
  const keywords = ['懲戒', 'discipline', '處分'];
  const nameLower = sheetName.toLowerCase();
  
  for (let i = 0; i < keywords.length; i++) {
    if (nameLower.indexOf(keywords[i]) !== -1) {
      return true;
    }
  }
  return false;
};

DisciplineParser.prototype.parse = function(worksheet, sheetName) {
  const data = this.worksheetToArray(worksheet);
  
  if (data.length < 2) {
    return { title: sheetName, headers: [], data: [] };
  }
  
  // 尋找標題列
  let headerRowIndex = 0;
  
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const row = data[i];
    for (let j = 0; j < row.length; j++) {
      const cell = String(row[j]);
      if (cell.indexOf('姓名') !== -1 || cell.indexOf('案由') !== -1) {
        headerRowIndex = i;
        break;
      }
    }
    if (headerRowIndex > 0) break;
  }
  
  const headers = data[headerRowIndex];
  const rows = [];
  
  for (let i = headerRowIndex + 1; i < data.length; i++) {
    const row = data[i];
    let hasContent = false;
    for (let j = 0; j < row.length; j++) {
      if (row[j] !== null && row[j] !== undefined && row[j] !== '') {
        hasContent = true;
        break;
      }
    }
    if (hasContent) {
      rows.push(row);
    }
  }
  
  return {
    title: sheetName,
    headers: headers,
    data: rows
  };
};

/**
 * 特約單位格式解析器
 */
function ContractUnitParser() {
  BaseParser.call(this);
  this.typeName = 'ContractUnit';
}

ContractUnitParser.prototype = Object.create(BaseParser.prototype);
ContractUnitParser.prototype.constructor = ContractUnitParser;

ContractUnitParser.prototype.canHandle = function(sheetName, worksheet) {
  const keywords = ['特約', '服務單位', '醫療院所', '診所', '醫院'];
  const nameLower = sheetName.toLowerCase();
  
  for (let i = 0; i < keywords.length; i++) {
    if (nameLower.indexOf(keywords[i]) !== -1) {
      return true;
    }
  }
  
  // 檢查第一列是否包含特約單位相關欄位
  const data = worksheet.data;
  if (data.length > 0) {
    const firstRow = data[0];
    for (let j = 0; j < firstRow.length; j++) {
      const cell = String(firstRow[j]);
      if (cell.indexOf('特約類別') !== -1 || cell.indexOf('機構名稱') !== -1) {
        return true;
      }
    }
  }
  
  return false;
};

ContractUnitParser.prototype.parse = function(worksheet, sheetName) {
  const data = this.worksheetToArray(worksheet);
  
  if (data.length < 2) {
    return { title: sheetName, headers: [], data: [] };
  }
  
  // 尋找標題列（可能在第2或第3列）
  let headerRowIndex = 0;
  
  for (let i = 0; i < Math.min(3, data.length); i++) {
    const row = data[i];
    let headerCount = 0;
    
    for (let j = 0; j < row.length; j++) {
      const cell = String(row[j]).trim();
      if (cell.length > 0 && cell.length < 20) {
        headerCount++;
      }
    }
    
    // 如果這一列有多個短文字，很可能是標題列
    if (headerCount >= 3) {
      headerRowIndex = i;
      break;
    }
  }
  
  const headers = data[headerRowIndex];
  const rows = [];
  
  for (let i = headerRowIndex + 1; i < data.length; i++) {
    const row = data[i];
    let hasContent = false;
    
    for (let j = 0; j < row.length; j++) {
      if (row[j] !== null && row[j] !== undefined && row[j] !== '') {
        hasContent = true;
        break;
      }
    }
    
    if (hasContent) {
      rows.push(row);
    }
  }
  
  return {
    title: sheetName,
    headers: headers,
    data: rows
  };
};

/**
 * 基礎解析器類別
 */
function BaseParser() {
  this.typeName = 'Unknown';
}

BaseParser.prototype.getTypeName = function() {
  return this.typeName;
};

BaseParser.prototype.canHandle = function(sheetName, worksheet) {
  return false;
};

BaseParser.prototype.parse = function(worksheet, sheetName) {
  throw new Error('parse() must be implemented');
};

BaseParser.prototype.parseWithHeaderRow = function(worksheet, sheetName, headerRow) {
  const data = this.worksheetToArray(worksheet);
  
  if (data.length === 0 || headerRow < 1 || headerRow > data.length) {
    return { title: sheetName, headers: [], data: [] };
  }
  
  // 使用指定的標題列（轉為0-based索引）
  const headers = data[headerRow - 1];
  const rows = [];
  
  // 從標題列的下一列開始讀取資料
  for (let i = headerRow; i < data.length; i++) {
    const row = data[i];
    
    // 過濾完全空白的列
    let hasContent = false;
    for (let j = 0; j < row.length; j++) {
      if (row[j] !== null && row[j] !== undefined && String(row[j]).trim() !== '') {
        hasContent = true;
        break;
      }
    }
    
    if (hasContent) {
      rows.push(row);
    }
  }
  
  return {
    title: sheetName,
    headers: headers,
    data: rows
  };
};

BaseParser.prototype.worksheetToArray = function(worksheet) {
  return worksheet.data;
};

/**
 * 預設解析器
 */
function DefaultParser() {
  BaseParser.call(this);
  this.typeName = 'Default';
}

DefaultParser.prototype = Object.create(BaseParser.prototype);
DefaultParser.prototype.constructor = DefaultParser;

DefaultParser.prototype.canHandle = function(sheetName, worksheet) {
  return true;
};

DefaultParser.prototype.parse = function(worksheet, sheetName) {
  return this.parseWithHeaderRow(worksheet, sheetName, 1);
};

/**
 * 懲戒格式解析器
 */
function DisciplineParser() {
  BaseParser.call(this);
  this.typeName = 'Discipline';
}

DisciplineParser.prototype = Object.create(BaseParser.prototype);
DisciplineParser.prototype.constructor = DisciplineParser;

DisciplineParser.prototype.canHandle = function(sheetName, worksheet) {
  const keywords = ['懲戒', 'discipline', '處分'];
  const nameLower = sheetName.toLowerCase();
  
  for (let i = 0; i < keywords.length; i++) {
    if (nameLower.indexOf(keywords[i]) !== -1) {
      return true;
    }
  }
  return false;
};

DisciplineParser.prototype.parse = function(worksheet, sheetName) {
  return this.parseWithHeaderRow(worksheet, sheetName, 1);
};

/**
 * 特約單位格式解析器
 */
function ContractUnitParser() {
  BaseParser.call(this);
  this.typeName = 'ContractUnit';
}

ContractUnitParser.prototype = Object.create(BaseParser.prototype);
ContractUnitParser.prototype.constructor = ContractUnitParser;

ContractUnitParser.prototype.canHandle = function(sheetName, worksheet) {
  const keywords = ['特約', '服務單位', '醫療院所', '診所', '醫院'];
  const nameLower = sheetName.toLowerCase();
  
  for (let i = 0; i < keywords.length; i++) {
    if (nameLower.indexOf(keywords[i]) !== -1) {
      return true;
    }
  }
  
  const data = worksheet.data;
  if (data.length > 0) {
    const firstRow = data[0];
    for (let j = 0; j < firstRow.length; j++) {
      const cell = String(firstRow[j]);
      if (cell.indexOf('特約類別') !== -1 || cell.indexOf('機構名稱') !== -1) {
        return true;
      }
    }
  }
  
  return false;
};

ContractUnitParser.prototype.parse = function(worksheet, sheetName) {
  return this.parseWithHeaderRow(worksheet, sheetName, 1);
};
