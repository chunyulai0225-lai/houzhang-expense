/**
 * Chuna.gs — 給出納彙總／匯出、對帳。逐行對照 chunaService.ts / reconciliationService.ts。
 *
 * 完全不重新計算金額，直接沿用 FeeCalculation.gs 已驗證的 summarizeTeacherMonthlyFees。
 * 自費代課獨立存在 SelfFunded 分頁（見 MonthlyClose.gs 開頭說明），天生不會出現在
 * monthlyImportIds 對應的 SubstituteRecords 裡，不需要像 Node 版本那樣額外用
 * fileName 排除虛擬容器批次。
 *
 * Excel 匯出限制（GAS 沒有 exceljs）：改用「建立暫存 Google Sheet → 寫入資料 →
 * 用 DriveApp 匯出成 .xlsx 二進位 → base64 編碼回傳給前端 → 前端解碼觸發下載 →
 * 刪除暫存檔」。排版（分組標題／表頭／逐列／總計列）跟 Node 版本的
 * generateChunaExcelBuffer 一致，只是產生方式不同，見 api_generateChunaExcel()。
 */

var CHUNA_SECTION_HEADER = ["代課教師代碼", "代課教師", "節數", "金額", "備註"];

function enrichWithPayrollCode(rows) {
  return rows
    .map(function (r) {
      var person = r.substituteTeacherId ? findById("Persons", r.substituteTeacherId) : null;
      return Object.assign({}, r, { payrollCode: person ? (person.payrollCode || null) : null });
    })
    // 目前真實給出納檔案看不出明確排序規則，這裡採用「節數由多到少」當預設，
    // 貼近真實檔案大部分列的呈現方式。
    .sort(function (a, b) { return b.totalCount - a.totalCount; });
}

// 給出納是公費核銷用的名冊，自費代課是學校/單位自行吸收的支出，不應該混在同一張表，
// 所以這裡只接受公費 Excel 匯入批次的 monthlyImportId（自費代課本來就不在 SubstituteRecords）。
function getChunaSummary(monthlyImportIds) {
  var idSet = {};
  monthlyImportIds.forEach(function (id) { idSet[id] = true; });
  var imports = readRows("MonthlyImports").filter(function (i) { return idSet[i.id]; });
  var bdIds = imports.filter(function (i) { return i.sourceStaffType === "BD"; }).map(function (i) { return i.id; });
  var nonBdIds = imports.filter(function (i) { return i.sourceStaffType === "NON_BD"; }).map(function (i) { return i.id; });
  var unknownIds = imports.filter(function (i) { return i.sourceStaffType === "UNKNOWN"; }).map(function (i) { return i.id; });

  return {
    bd: enrichWithPayrollCode(bdIds.length > 0 ? summarizeTeacherMonthlyFees(bdIds) : []),
    nonBd: enrichWithPayrollCode(nonBdIds.length > 0 ? summarizeTeacherMonthlyFees(nonBdIds) : []),
    unknown: enrichWithPayrollCode(unknownIds.length > 0 ? summarizeTeacherMonthlyFees(unknownIds) : []),
  };
}

function api_getChunaSummary(payload) {
  requireField(payload, "monthlyImportIds", "monthlyImportIds");
  return getChunaSummary(payload.monthlyImportIds);
}

function writeChunaSection(sheet, startRow, title, rows) {
  var r = startRow;
  sheet.getRange(r, 1).setValue(title);
  r += 1;
  sheet.getRange(r, 1, 1, CHUNA_SECTION_HEADER.length).setValues([CHUNA_SECTION_HEADER]);
  r += 1;
  var totalCount = 0, totalAmount = 0;
  rows.forEach(function (row) {
    sheet.getRange(r, 1, 1, 5).setValues([[row.payrollCode || "", row.substituteTeacherName, row.totalCount, Number(row.totalAmount), ""]]);
    totalCount += row.totalCount;
    totalAmount += Number(row.totalAmount);
    r += 1;
  });
  sheet.getRange(r, 1, 1, 5).setValues([["", "總計", totalCount, totalAmount, ""]]);
  r += 2;
  return r;
}

// 產生給出納 .xlsx。GAS 沒有 exceljs，改用「暫存 Google Sheet → DriveApp 匯出 xlsx →
// base64」，用完即刪暫存檔，不會留在使用者的雲端硬碟裡。
function api_generateChunaExcel(payload) {
  requireField(payload, "monthlyImportIds", "monthlyImportIds");
  requireField(payload, "year", "年");
  requireField(payload, "month", "月");
  var year = Number(payload.year), month = Number(payload.month);
  var summary = getChunaSummary(payload.monthlyImportIds);

  var tempName = "_temp_chuna_" + year + "_" + month + "_" + newId();
  var tempSs = SpreadsheetApp.create(tempName);
  var sheet = tempSs.getSheets()[0];
  sheet.setName((year + "年" + month + "月給出納").slice(0, 100));

  var r = 1;
  if (summary.nonBd.length > 0) r = writeChunaSection(sheet, r, year + "年" + month + "月公費代課（編制外）", summary.nonBd);
  if (summary.bd.length > 0) r = writeChunaSection(sheet, r, year + "年" + month + "月公費代課（編制內）", summary.bd);
  if (summary.unknown.length > 0) r = writeChunaSection(sheet, r, year + "年" + month + "月公費代課（未標示來源）", summary.unknown);
  for (var c = 1; c <= 5; c++) sheet.setColumnWidth(c, 140);

  SpreadsheetApp.flush();
  var fileId = tempSs.getId();
  var blob = DriveApp.getFileById(fileId).getAs(MimeType.MICROSOFT_EXCEL);
  var base64 = Utilities.base64Encode(blob.getBytes());
  DriveApp.getFileById(fileId).setTrashed(true);

  return { fileName: year + "年" + month + "月給出納.xlsx", mimeType: MimeType.MICROSOFT_EXCEL, base64: base64 };
}

// ---------- 對帳 ----------
// 讀取使用者「上傳的原始給出納」逐教師跟系統算出來的結果比對。不會只比總金額；
// 每個差異都嘗試從資料本身找「可能原因」，但不宣稱「已確認原因」。
//
// GAS 限制：解析上傳的 .xlsx 二進位（原本用 exceljs）無法在 GAS 端執行，改成跟
// Excel 匯入一樣的做法——由前端用 SheetJS 在瀏覽器解析成 [{name, periodCount, amount}]
// 陣列後，直接把解析結果 POST 過來（payload.uploadedRows），這裡只做比對本身。

function guessPossibleReason(name, monthlyImportIds, diffDirection) {
  if (diffDirection === "less" || diffDirection === "onlyOriginal") {
    var idSet = {};
    monthlyImportIds.forEach(function (id) { idSet[id] = true; });
    var matchingRaw = readRows("RawRecords").filter(function (raw) { return idSet[raw.monthlyImportId] && raw.substituteTeacherText === name; });
    if (matchingRaw.length === 0) {
      return "尚無法從資料自動判斷原因，請人工確認（例如：原始給出納可能包含固定課表／整月代理等系統目前資料來源沒有涵蓋的項目）";
    }
    var processedRawIds = {};
    readRows("SubstituteRecords").forEach(function (r) { if (r.rawRecordId) processedRawIds[r.rawRecordId] = true; });
    var unprocessed = matchingRaw.filter(function (raw) { return !processedRawIds[raw.id]; });
    if (unprocessed.length > 0) {
      return "可能原因：有 " + unprocessed.length + " 筆該教師的原始資料因為日期／節次格式問題未能成功匯入（可在匯入錯誤列表查證）";
    }
    return "尚無法從資料自動判斷原因，請人工確認（例如：原始給出納可能包含固定課表／整月代理等系統目前資料來源沒有涵蓋的項目）";
  }
  if (diffDirection === "more") {
    return "可能原因：系統包含原始給出納未列出的紀錄，或原始給出納對這筆金額另有特殊處理（例如超鐘點另計）——請人工確認";
  }
  return null;
}

function reconcile(monthlyImportIds, uploadedRows) {
  var summary = getChunaSummary(monthlyImportIds);
  var allSystemRows = [].concat(summary.bd, summary.nonBd, summary.unknown);

  // 同一個姓名可能同時出現在 BD／非BD／未標示來源這幾個分組（例如同一位教師這個月
  // 同時有編制內跟編制外的代課紀錄）——這裡要把節數／金額加總起來，絕對不能讓後面
  // 的分組直接蓋掉前面的，否則系統總額會悄悄少算，對帳就失去意義了。
  var systemByName = {};
  var systemNameOrder = [];
  allSystemRows.forEach(function (r) {
    var key = String(r.substituteTeacherName).replace("（未配對）", "").trim();
    var existing = systemByName[key];
    if (existing) {
      existing.generalCount += r.generalCount;
      existing.generalAmount = decToStr(Number(existing.generalAmount) + Number(r.generalAmount));
      existing.overtimeCount += r.overtimeCount;
      existing.overtimeAmount = decToStr(Number(existing.overtimeAmount) + Number(r.overtimeAmount));
      existing.projectCount += r.projectCount;
      existing.projectAmount = decToStr(Number(existing.projectAmount) + Number(r.projectAmount));
      existing.totalCount += r.totalCount;
      existing.totalAmount = decToStr(Number(existing.totalAmount) + Number(r.totalAmount));
    } else {
      systemByName[key] = Object.assign({}, r);
      systemNameOrder.push(key);
    }
  });

  var uploadedByName = {};
  (uploadedRows || []).forEach(function (r) { uploadedByName[String(r.name).trim()] = r; });

  var allNamesSet = {};
  systemNameOrder.forEach(function (n) { allNamesSet[n] = true; });
  Object.keys(uploadedByName).forEach(function (n) { allNamesSet[n] = true; });
  var allNames = Object.keys(allNamesSet);

  var rows = [];
  var totalSystem = 0, totalOriginal = 0;

  allNames.forEach(function (name) {
    var sys = systemByName[name];
    var orig = uploadedByName[name];
    var systemAmount = sys ? Number(sys.totalAmount) : null;
    var originalAmount = orig ? Number(orig.amount) : null;
    totalSystem += systemAmount || 0;
    totalOriginal += originalAmount || 0;

    var status, possibleReason = null;
    var amountDiff = systemAmount !== null && originalAmount !== null ? systemAmount - originalAmount : null;

    if (sys && !orig) {
      status = "ONLY_SYSTEM";
    } else if (!sys && orig) {
      status = "ONLY_ORIGINAL";
      possibleReason = guessPossibleReason(name, monthlyImportIds, "onlyOriginal");
    } else if (amountDiff === 0) {
      status = "MATCH";
    } else if (amountDiff !== null && amountDiff < 0) {
      status = "SYSTEM_LESS";
      possibleReason = guessPossibleReason(name, monthlyImportIds, "less");
    } else if (amountDiff !== null && amountDiff > 0) {
      status = "SYSTEM_MORE";
      possibleReason = guessPossibleReason(name, monthlyImportIds, "more");
    } else {
      status = "UNCERTAIN";
    }

    rows.push({
      name: name, systemPeriodCount: sys ? sys.totalCount : null, originalPeriodCount: orig ? orig.periodCount : null,
      systemAmount: systemAmount, originalAmount: originalAmount, amountDiff: amountDiff, status: status, possibleReason: possibleReason,
    });
  });

  rows.sort(function (a, b) { return Math.abs(b.amountDiff || 0) - Math.abs(a.amountDiff || 0); });

  return { rows: rows, totals: { systemAmount: totalSystem, originalAmount: totalOriginal, diff: totalSystem - totalOriginal } };
}

function api_reconcile(payload) {
  requireField(payload, "monthlyImportIds", "monthlyImportIds");
  requireField(payload, "uploadedRows", "uploadedRows");
  var result = reconcile(payload.monthlyImportIds, payload.uploadedRows);

  // 附加稽核紀錄（見 Setup.gs 開頭說明）：每次執行對帳多寫一筆歷史快照，
  // 不是比對邏輯的資料來源，即時比對永遠重新算。
  var runId = newId();
  result.rows.forEach(function (row) {
    appendRow("Reconciliation", {
      id: newId(), runId: runId, monthlyImportIds: JSON.stringify(payload.monthlyImportIds), name: row.name,
      systemPeriodCount: row.systemPeriodCount === null ? "" : row.systemPeriodCount,
      originalPeriodCount: row.originalPeriodCount === null ? "" : row.originalPeriodCount,
      systemAmount: row.systemAmount === null ? "" : row.systemAmount,
      originalAmount: row.originalAmount === null ? "" : row.originalAmount,
      amountDiff: row.amountDiff === null ? "" : row.amountDiff,
      status: row.status, possibleReason: row.possibleReason || "", createdAt: nowIso(), createdBy: payload.changedBy || "",
    });
  });

  return result;
}
