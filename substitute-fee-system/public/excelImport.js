// excelImport.js — 瀏覽器端 Excel 解析（用 SheetJS）。
//
// GAS 沒有 exceljs，沒辦法在後端解析 Excel 二進位檔案，所以把「讀檔＋抓表頭欄位」
// 整段搬到瀏覽器執行，解析結果的形狀跟 Node 版本 excelImportService.ts 的
// ParsedExcelRow 完全一樣，這樣才能直接把 rows 陣列 POST 給 GAS 的
// importSubstituteRows action，後續解析（日期/節次標準化、版本化）完全不用改寫。
//
// 欄位比對規則、表頭列偵測、資料列判定「是否為空白列」的邏輯，都是逐行
// 對照 excelImportService.ts 的 parseWorkbook()／findHeaderRowNumber() 抄過來的，
// 不是重新設計。

const EXACT_COLUMN_ALIASES = {
  rowNumber: ["序", "序號"],
  originalTeacherCode: ["原教師代碼", "原教師代号"],
  originalTeacherName: ["原教師"],
  dateText: ["日期"],
  leaveType: ["假別"],
  hoursOrDaysText: ["時數天數", "時數/天數", "時數／天數"],
  periodText: ["節次"],
  className: ["班級"],
  subject: ["科目"],
  substituteTeacherCode: ["代課教師代碼", "代課教師代号"],
  substituteTeacherName: ["代課教師"],
  teacherCert: ["教師證"],
  payGrade: ["薪等"],
  homeroomFeeText: ["代導師"],
  periodCountText: ["節數"],
};

// 局部比對（表頭常包含會逐月變動的文字，例如金額、天數），只在完全比對找不到時才嘗試。
const FUZZY_COLUMN_MARKERS = {
  homeroomFeeText: "代導師費",
  dailyOrHalfDayWageText: "日薪",
  substitutePeriodFeeText: "代課鐘點",
};

const REQUIRED_FIELDS = ["dateText", "originalTeacherName", "substituteTeacherName", "periodText"];

function normalizeHeader(text) {
  return String(text).replace(/\s+/g, "");
}

function cellToString(ws, addr) {
  const cell = ws[addr];
  if (!cell) return null;
  const v = cell.v;
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v.toISOString();
  const s = String(v).trim();
  return s === "" ? null : s;
}

function isRowEffectivelyEmpty(row) {
  return !(row.dateText || row.originalTeacherName || row.substituteTeacherName || row.periodText || row.className || row.subject);
}

// 讀取檔案並回傳 SheetJS workbook 物件，供 listWorkbookSheetsFromFile()／
// extractSheetRows() 共用（同一個 workbook 只解析一次二進位）。
function readWorkbookFromFile(file) {
  return file.arrayBuffer().then((buf) => XLSX.read(buf, { type: "array", cellDates: true }));
}

// 對應 excelImportService.ts 的 listWorkbookSheets()：列出所有工作表名稱與列數，
// 供上傳前先讓管理者確認要匯入哪一個 Sheet。
function listWorkbookSheetsFromWorkbook(workbook) {
  return workbook.SheetNames.map((name) => {
    const ws = workbook.Sheets[name];
    const ref = ws["!ref"];
    let rowCount = 0;
    if (ref) {
      const range = XLSX.utils.decode_range(ref);
      rowCount = range.e.r - range.s.r + 1;
    }
    return { name, rowCount };
  });
}

// 跟 server.ts /api/monthly-imports/inspect 完全一樣的猜測規則，只是搬到前端。
function suggestStaffTypeFromSheetName(name) {
  if (name.includes("非BD")) return "NON_BD";
  if (name.includes("BD")) return "BD";
  return "UNKNOWN";
}

// 在前 10 列中找出「看起來最像表頭」的那一列：比對得到的欄位數最多者勝出。
function findHeaderRowNumber(ws, range) {
  const maxScanRow = Math.min(range.s.r + 10, range.e.r + 1);
  let bestRow = range.s.r;
  let bestScore = -1;
  for (let r = range.s.r; r < maxScanRow; r++) {
    let score = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const text = cellToString(ws, XLSX.utils.encode_cell({ r, c }));
      if (!text) continue;
      const normalized = normalizeHeader(text);
      const isExact = Object.values(EXACT_COLUMN_ALIASES).some((aliases) => aliases.some((alias) => normalizeHeader(alias) === normalized));
      if (isExact) score += 1;
    }
    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }
  return bestRow;
}

// 對應 excelImportService.ts 的 parseWorkbook()：回傳 {headers, rows}，
// rows 的形狀跟 GAS 端 api_importSubstituteRows() 預期的 ParsedExcelRow 一致。
function extractSheetRows(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) throw new Error(`找不到名為「${sheetName}」的工作表`);
  const ref = ws["!ref"];
  if (!ref) throw new Error("這個工作表是空的");
  const range = XLSX.utils.decode_range(ref);

  const headerRowIdx = findHeaderRowNumber(ws, range);
  const headers = [];
  const columnMap = {};
  const claimedFields = new Set();

  for (let c = range.s.c; c <= range.e.c; c++) {
    const text = cellToString(ws, XLSX.utils.encode_cell({ r: headerRowIdx, c }));
    if (!text) continue;
    headers.push(text);
    const normalized = normalizeHeader(text);
    for (const [field, aliases] of Object.entries(EXACT_COLUMN_ALIASES)) {
      if (claimedFields.has(field)) continue;
      if (aliases.some((alias) => normalizeHeader(alias) === normalized)) {
        columnMap[field] = c;
        claimedFields.add(field);
        break;
      }
    }
  }
  // 完全比對結束後，剩下的欄位再嘗試局部比對（表頭包含變動文字的情況）
  for (let c = range.s.c; c <= range.e.c; c++) {
    const text = cellToString(ws, XLSX.utils.encode_cell({ r: headerRowIdx, c }));
    if (!text) continue;
    const normalized = normalizeHeader(text);
    for (const [field, marker] of Object.entries(FUZZY_COLUMN_MARKERS)) {
      if (claimedFields.has(field)) continue;
      if (normalized.includes(marker)) {
        columnMap[field] = c;
        claimedFields.add(field);
      }
    }
  }

  const missingRequired = REQUIRED_FIELDS.filter((f) => !(f in columnMap));
  if (missingRequired.length > 0) {
    const labels = missingRequired.map((f) => EXACT_COLUMN_ALIASES[f][0]).join("、");
    throw new Error(`Excel 缺少必要欄位：${labels}（找到的欄位：${headers.join("、") || "無"}）`);
  }

  const headerTextByCol = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const text = cellToString(ws, XLSX.utils.encode_cell({ r: headerRowIdx, c }));
    if (text) headerTextByCol[c] = text;
  }

  function get(r, field) {
    const c = columnMap[field];
    if (c === undefined) return undefined;
    return cellToString(ws, XLSX.utils.encode_cell({ r, c })) || undefined;
  }

  const rows = [];
  let dataRowSeq = 0;
  for (let r = headerRowIdx + 1; r <= range.e.r; r++) {
    const raw = {};
    Object.keys(headerTextByCol).forEach((cKey) => {
      const c = Number(cKey);
      raw[headerTextByCol[c]] = cellToString(ws, XLSX.utils.encode_cell({ r, c }));
    });

    const parsed = {
      rowNumber: 0,
      raw,
      originalTeacherCode: get(r, "originalTeacherCode"),
      originalTeacherName: get(r, "originalTeacherName"),
      dateText: get(r, "dateText"),
      leaveType: get(r, "leaveType"),
      hoursOrDaysText: get(r, "hoursOrDaysText"),
      periodText: get(r, "periodText"),
      className: get(r, "className"),
      subject: get(r, "subject"),
      substituteTeacherCode: get(r, "substituteTeacherCode"),
      substituteTeacherName: get(r, "substituteTeacherName"),
      teacherCert: get(r, "teacherCert"),
      payGrade: get(r, "payGrade"),
      homeroomFeeText: get(r, "homeroomFeeText"),
      dailyOrHalfDayWageText: get(r, "dailyOrHalfDayWageText"),
      substitutePeriodFeeText: get(r, "substitutePeriodFeeText"),
      periodCountText: get(r, "periodCountText"),
    };

    if (isRowEffectivelyEmpty(parsed)) continue;
    dataRowSeq += 1;
    parsed.rowNumber = dataRowSeq;
    rows.push(parsed);
  }

  return { headers, rows };
}

// ---------- 對帳用：解析使用者上傳的原始「給出納」Excel ----------
// 對應 reconciliationService.ts 的 parseUploadedChunaWorkbook()：掃描所有工作表，
// 找出「代課教師」或「教師」這種姓名欄，搭配「節數」「金額／總額」欄，自動找出
// 對應的欄位位置，不假設固定在第幾欄。

function parseUploadedChunaWorkbook(workbook) {
  const rows = [];
  workbook.SheetNames.forEach((sheetName) => {
    const ws = workbook.Sheets[sheetName];
    const ref = ws["!ref"];
    if (!ref) return;
    const range = XLSX.utils.decode_range(ref);

    let nameCol = -1, periodCol = -1, amountCol = -1, headerRow = -1;
    const maxScanRow = Math.min(range.s.r + 15, range.e.r + 1);
    for (let r = range.s.r; r < maxScanRow; r++) {
      let foundName = -1, foundPeriod = -1, foundAmount = -1;
      const maxScanCol = Math.min(range.s.c + 30, range.e.c + 1);
      for (let c = range.s.c; c < maxScanCol; c++) {
        const text = (cellToString(ws, XLSX.utils.encode_cell({ r, c })) || "").replace(/\s|\n/g, "");
        if (!text) continue;
        if (text.includes("教師") && !text.includes("代碼") && foundName === -1) foundName = c;
        if (text.includes("節數") && foundPeriod === -1) foundPeriod = c;
        if ((text.includes("總額") || text.includes("金額")) && foundAmount === -1) foundAmount = c;
      }
      if (foundName !== -1 && foundPeriod !== -1 && foundAmount !== -1) {
        nameCol = foundName;
        periodCol = foundPeriod;
        amountCol = foundAmount;
        headerRow = r;
        break;
      }
    }

    if (headerRow === -1) return; // 這個工作表看起來不是給出納格式，略過

    for (let r = headerRow + 1; r <= range.e.r; r++) {
      const name = (cellToString(ws, XLSX.utils.encode_cell({ r, c: nameCol })) || "").trim();
      const periodText = (cellToString(ws, XLSX.utils.encode_cell({ r, c: periodCol })) || "").trim();
      const amountText = (cellToString(ws, XLSX.utils.encode_cell({ r, c: amountCol })) || "").trim();
      if (!name || name.includes("總計") || name.includes("小計")) continue;
      const periodCount = Number(periodText);
      const amount = Number(amountText);
      if (Number.isNaN(periodCount) || Number.isNaN(amount)) continue;
      rows.push({ name, periodCount, amount });
    }
  });
  return rows;
}
