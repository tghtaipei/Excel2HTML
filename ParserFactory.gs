/**
 * 解析器工廠
 */
const ParserFactory = {
  
  /**
   * 取得適當的解析器
   */
  getParser: function(sheetName, worksheet) {
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
