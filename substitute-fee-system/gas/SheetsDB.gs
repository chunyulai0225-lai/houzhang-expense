/**
 * SheetsDB.gs — 通用 Google Sheets CRUD helper。
 *
 * 每個 domain service（Classification.gs／FeeCalculation.gs…）都透過這幾個
 * 函式存取資料，不會各自重複寫一次「把整張表讀出來、找表頭欄位」這種邏輯。
 *
 * 所有資料列都當作「物件的陣列」處理：readRows() 回傳 [{col1:val1,...}, ...]，
 * appendRow()/updateRow() 收物件、依表頭順序寫回對應欄位。
 */

function getSheet(name) {
  var sheet = getSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error("找不到分頁「" + name + "」，請先執行 setupSheets()");
  return sheet;
}

function getHeaders(name) {
  var sheet = getSheet(name);
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

// 讀出整張表（不含表頭），每一列轉成 {headerName: value} 物件。
// 值一律以字串型式回傳（getSheet 已把欄位設成純文字格式），由呼叫端自行轉型別，
// 避免 Sheets 自動型別推斷造成「0 被讀成空字串」這類跟金錢有關的錯誤。
function readRows(name) {
  var sheet = getSheet(name);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol === 0) return [];
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var rows = [];
  for (var r = 0; r < values.length; r++) {
    var obj = { __row: r + 2 }; // 記住實際列號，updateRow/deleteRow 用得到
    var allBlank = true;
    for (var c = 0; c < headers.length; c++) {
      var v = values[r][c];
      if (v !== "" && v !== null && v !== undefined) allBlank = false;
      obj[headers[c]] = v === "" ? null : v;
    }
    if (!allBlank) rows.push(obj);
  }
  return rows;
}

function findRows(name, predicate) {
  return readRows(name).filter(predicate);
}

function findOne(name, predicate) {
  var rows = findRows(name, predicate);
  return rows.length > 0 ? rows[0] : null;
}

function findById(name, id) {
  return findOne(name, function (r) { return r.id === id; });
}

// 新增一列。obj 沒填的欄位一律留空字串，欄位順序完全依照 SHEET_SCHEMAS 定義。
function appendRow(name, obj) {
  var sheet = getSheet(name);
  var headers = getHeaders(name);
  var row = headers.map(function (h) {
    var v = obj[h];
    return v === undefined || v === null ? "" : v;
  });
  sheet.appendRow(row);
  return obj;
}

// 依 id 局部更新（只覆蓋 patch 裡出現的欄位，其餘欄位維持原值）。
function updateRow(name, id, patch) {
  var sheet = getSheet(name);
  var headers = getHeaders(name);
  var existing = findById(name, id);
  if (!existing) throw new Error("在「" + name + "」找不到 id=" + id);

  var merged = {};
  headers.forEach(function (h) {
    merged[h] = Object.prototype.hasOwnProperty.call(patch, h) ? patch[h] : existing[h];
  });
  var rowValues = headers.map(function (h) {
    var v = merged[h];
    return v === undefined || v === null ? "" : v;
  });
  sheet.getRange(existing.__row, 1, 1, headers.length).setValues([rowValues]);
  return merged;
}

function deleteRow(name, id) {
  var sheet = getSheet(name);
  var existing = findById(name, id);
  if (!existing) return false;
  sheet.deleteRow(existing.__row);
  return true;
}

// 批次新增（匯入用，逐列 appendRow 在 200~300 列時仍偏慢，改用一次性 setValues）。
function appendRows(name, objs) {
  if (objs.length === 0) return;
  var sheet = getSheet(name);
  var headers = getHeaders(name);
  var values = objs.map(function (obj) {
    return headers.map(function (h) {
      var v = obj[h];
      return v === undefined || v === null ? "" : v;
    });
  });
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, values.length, headers.length).setValues(values);
}
