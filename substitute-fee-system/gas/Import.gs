/**
 * Import.gs — 公費代課 Excel 匯入（GAS 端）。
 *
 * Excel 檔案的「讀取二進位＋抓表頭欄位」搬到瀏覽器用 SheetJS 做（見 app.js
 * parseExcelFileToRows()），這裡收到的 payload.rows 已經是跟 Node 版本
 * ParsedExcelRow 完全一樣形狀的物件陣列。從這裡開始（日期/節次解析、
 * 版本化、寫 RawRecords/SubstituteRecords）逐行對照 excelImportService.ts
 * 的 importSubstituteExcel()，邏輯完全一致，只是換成操作 Sheets。
 */

var WEEKDAY_CHAR_MAP = { "一": "MON", "二": "TUE", "三": "WED", "四": "THU", "五": "FRI", "六": "SAT", "日": "SUN", "天": "SUN" };
var CHINESE_DIGIT_MAP = { "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7 };
var DATE_PATTERN = /^(\d{1,2})[-\/](\d{1,2})(?:\((一|二|三|四|五|六|日)\))?(?:\s*\d{1,2}:\d{2}\s*~\s*\d{1,2}:\d{2})?$/;
// 真實 2026/09 資料出現的日期「區間」格式，例如「09-11(五) 12:00 ~ 09-14(一) 16:00」——
// 前後各是一個完整的「MM-DD(星期) HH:mm」，中間用 ~ 連接。這跟 DATE_PATTERN 本來就支援的
// 「單日 + 同一天的起訖時間」（例如「6/18(四) 13:50 ~ 15:50」）是兩種不同的格式，
// 不能混在同一個 pattern 裡。這種區間實際代表的是「原教師請假/出差的期間」，不是
// 「要展開成好幾筆代課紀錄」——真正的代課節次與計費數量另外由「時數天數」欄位決定
// （見下面 parseHoursOrDaysToPeriodCount()），這裡只需要把區間的起始日當成這筆
// SubstituteRecord 的代課日期（單純作為日期欄位的定位點，不是用來推算實際代課日），
// 不猜區間裡哪幾天才是真正代課日、也不會展開成多筆紀錄。
var DATE_RANGE_PATTERN =
  /^(\d{1,2})[-\/](\d{1,2})(?:\((一|二|三|四|五|六|日)\))?\s*\d{1,2}:\d{2}\s*~\s*(\d{1,2})[-\/](\d{1,2})(?:\((一|二|三|四|五|六|日)\))?\s*\d{1,2}:\d{2}$/;

function buildParsedDateFromParts(year, month, day, textWeekdayChar, expectedMonth) {
  if (month < 1 || month > 12) return null;
  var date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  var computedWeekday = WEEKDAY_BY_JS_INDEX[date.getUTCDay()];
  var weekdayMismatch = Boolean(textWeekdayChar) && WEEKDAY_CHAR_MAP[textWeekdayChar] !== computedWeekday;
  var monthMismatch = month !== expectedMonth;
  return { date: date.toISOString().slice(0, 10), weekday: computedWeekday, weekdayMismatch: weekdayMismatch, monthMismatch: monthMismatch };
}

// 回傳三種結果之一：
//   {date, weekday, weekdayMismatch, monthMismatch, isDateRange?} —— 可以正常解析成一天
//     （isDateRange 為 true 時，date/weekday 是「區間起始日」，只作為定位點使用）
//   {error}                                                        —— 真的無法辨識
function parseDateText(text, year, expectedMonth) {
  var trimmed = String(text).trim();
  var match = trimmed.match(DATE_PATTERN);
  if (!match) {
    var rangeMatch = trimmed.match(DATE_RANGE_PATTERN);
    if (rangeMatch) {
      var startParsed = buildParsedDateFromParts(year, Number(rangeMatch[1]), Number(rangeMatch[2]), rangeMatch[3], expectedMonth);
      if (!startParsed) return { error: '日期區間的起始日不合理："' + text + '"' };
      startParsed.isDateRange = true;
      return startParsed;
    }
    return { error: '無法解析日期格式："' + text + '"，預期格式如 "06-18(四)" 或 "6/18(四) 13:50 ~ 15:50"' };
  }
  var month = Number(match[1]);
  var day = Number(match[2]);
  if (month < 1 || month > 12) return { error: '日期月份不合理："' + text + '"' };
  var parsed = buildParsedDateFromParts(year, month, day, match[3], expectedMonth);
  if (!parsed) return { error: '日期不存在："' + text + '"' };
  return parsed;
}

// 只認得「純日數」這種能安全、明確換算成一般代課節次計費數量的格式（例如「5日」
// 「8日」「12日」）：時數天數是 N 天、搭配單一節次欄位，代表這個代課教師在請假／
// 出差期間內每天都代同一節，共 N 次一般代課節次。混合單位（例如「3日4時」）或純
// 時數（例如「6時」）目前沒有既有、明確又正確的換算規則，寧可回傳無法解析、
// 讓呼叫端標記「待確認」，也不要自行推導、用錯誤的假設產生假的計費數字。
function parseHoursOrDaysToPeriodCount(text) {
  var trimmed = String(text || "").trim();
  var match = trimmed.match(/^(\d+)日$/);
  if (match) return { periodCount: Number(match[1]) };
  return { error: true };
}

// 回傳三種結果之一：
//   {periodCode}                                      —— 可以對應到系統既有的節次代碼
//   {isSpecialPeriod: true, specialPeriodText: "..."} —— 辨識出是特殊節次（例如「導師時間」），
//                                                        但不是既有的 P1~P7／早自修／午休
//   {error}                                            —— 真的無法辨識
function parsePeriodText(text, validCodes) {
  var trimmed = String(text).trim();
  function check(code) {
    return validCodes[code] ? { periodCode: code } : { error: "節次代碼「" + code + "」尚未設定於系統（原始文字：\"" + text + "\"）" };
  }
  if (trimmed === "早自修") return check("EARLY_STUDY");
  if (trimmed === "午休") return check("LUNCH");
  // 「導師時間」是真實存在的節次文字，但不是既有 P1~P7 的其中一節，不能硬轉成某一節
  // （代導師費規則尚未確認，見 FeeCalculation.gs 開頭說明——這裡故意不猜）。
  if (trimmed === "導師時間") return { isSpecialPeriod: true, specialPeriodText: trimmed };
  var match = trimmed.match(/^第([0-9一二三四五六七]+)節$/);
  if (match) {
    var raw = match[1];
    var num = /^[0-9]+$/.test(raw) ? Number(raw) : CHINESE_DIGIT_MAP[raw];
    if (num) return check("P" + num);
  }
  return { error: '無法解析節次："' + text + '"' };
}

function computeNextVersionNo(semesterId, year, month, sourceStaffType) {
  var rows = readRows("MonthlyImports").filter(function (r) {
    return r.semesterId === semesterId && Number(r.year) === year && Number(r.month) === month && r.sourceStaffType === sourceStaffType;
  });
  var maxV = 0;
  rows.forEach(function (r) { if (Number(r.versionNo) > maxV) maxV = Number(r.versionNo); });
  return maxV + 1;
}

// payload: { semesterId, year, month, fileName, sheetName, sourceStaffType, importedBy, rows: [ParsedExcelRow], detectedHeaders: [string] }
function api_importSubstituteRows(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "year", "年");
  requireField(payload, "month", "月");
  var semester = findById("Semesters", payload.semesterId);
  if (!semester) throw new Error("找不到學期");
  assertMonthNotLocked(Number(payload.year), Number(payload.month));

  var periodSlots = readRows("PeriodSlots");
  var validPeriodCodes = {};
  periodSlots.forEach(function (p) { validPeriodCodes[p.code] = true; });
  var sourceStaffType = payload.sourceStaffType || "UNKNOWN";
  var rows = payload.rows || [];

  var previousActive = readRows("MonthlyImports").filter(function (r) {
    return r.semesterId === payload.semesterId && Number(r.year) === Number(payload.year) &&
      Number(r.month) === Number(payload.month) && r.sourceStaffType === sourceStaffType && r.status === "ACTIVE";
  });
  var nextVersionNo = computeNextVersionNo(payload.semesterId, Number(payload.year), Number(payload.month), sourceStaffType);

  var monthlyImport = {
    id: newId(), semesterId: payload.semesterId, year: Number(payload.year), month: Number(payload.month),
    importedAt: nowIso(), fileName: payload.fileName || "", importedBy: payload.importedBy || "",
    totalCount: rows.length, successCount: 0, errorCount: 0, versionNo: nextVersionNo, status: "ACTIVE",
    sourceStaffType: sourceStaffType, sourceSheetName: payload.sheetName || "", note: "", createdAt: nowIso(),
  };
  appendRow("MonthlyImports", monthlyImport);

  previousActive.forEach(function (b) {
    updateRow("MonthlyImports", b.id, { status: "SUPERSEDED" });
    writeChangeLog("monthly_imports", b.id, "status", "ACTIVE", "SUPERSEDED", payload.importedBy, "重新匯入，被新批次取代（" + monthlyImport.id + "）");
  });

  var issues = [];
  var successCount = 0;
  var rawRowsToInsert = [];
  var recordRowsToInsert = [];

  rows.forEach(function (row) {
    var rawId = newId();
    rawRowsToInsert.push({
      id: rawId, monthlyImportId: monthlyImport.id, rowNumber: row.rowNumber,
      originalTeacherCodeText: row.originalTeacherCode || "", originalTeacherText: row.originalTeacherName || "",
      dateText: row.dateText || "", leaveTypeText: row.leaveType || "", hoursOrDaysText: row.hoursOrDaysText || "",
      periodText: row.periodText || "", classText: row.className || "", subjectText: row.subject || "",
      substituteTeacherCodeText: row.substituteTeacherCode || "", substituteTeacherText: row.substituteTeacherName || "",
      teacherCertText: row.teacherCert || "", payGradeText: row.payGrade || "", homeroomFeeText: row.homeroomFeeText || "",
      dailyOrHalfDayWageText: row.dailyOrHalfDayWageText || "", substitutePeriodFeeText: row.substitutePeriodFeeText,
      periodCountText: row.periodCountText || "", sheetName: payload.sheetName || "",
      rawJson: JSON.stringify(row.raw || {}), processedRecordId: "", createdAt: nowIso(),
    });

    var rowIssues = [];
    if (!row.dateText) rowIssues.push({ rowNumber: row.rowNumber, fieldName: "日期", message: "缺少日期" });
    if (!row.originalTeacherName) rowIssues.push({ rowNumber: row.rowNumber, fieldName: "原教師", message: "缺少原教師" });
    if (!row.substituteTeacherName) rowIssues.push({ rowNumber: row.rowNumber, fieldName: "代課教師", message: "缺少代課教師" });
    if (!row.periodText) rowIssues.push({ rowNumber: row.rowNumber, fieldName: "節次", message: "缺少節次" });

    var parsedDate = null;
    if (row.dateText) {
      var dr = parseDateText(row.dateText, Number(payload.year), Number(payload.month));
      if (dr.error) {
        rowIssues.push({ rowNumber: row.rowNumber, fieldName: "日期", message: dr.error });
      } else {
        // 日期可能是單一天，也可能是一個區間（例如「09-11(五) 12:00 ~ 09-14(一) 16:00」）——
        // 兩者都視為可以正常解析，不會因為是區間就判定成錯誤或待確認；區間本身不會展開
        // 成多筆紀錄，parsedDate.date 只是區間的起始日，作為這筆 SubstituteRecord 的
        // 日期定位點使用（見 parseDateText 開頭說明）。
        parsedDate = dr;
      }
    }
    var periodCode = null;
    if (row.periodText) {
      var pr = parsePeriodText(row.periodText, validPeriodCodes);
      if (pr.isSpecialPeriod) {
        // 「導師時間」這類特殊節次：保留原始文字，不硬轉成 P1~P7 的某一節，也不猜代課費
        // （代導師費／日薪／半日薪的計算規則尚未確認）。原始資料照樣完整保留，只是
        // 不建立 SubstituteRecord，記錄成「待確認」，等規則確認後再另外處理。
        rowIssues.push({
          rowNumber: row.rowNumber, fieldName: "節次",
          message: '特殊節次／待確認："' + pr.specialPeriodText + '"：尚未確認對應的代課費計算規則，不會自動歸類到既有節次，暫不建立代課紀錄',
        });
      } else if (pr.error) {
        rowIssues.push({ rowNumber: row.rowNumber, fieldName: "節次", message: pr.error });
      } else {
        periodCode = pr.periodCode;
      }
    }

    // 日期是「區間」、且節次是第一節～第七節時，這筆紀錄的計費數量（periodCount）
    // 一律依「時數天數」欄位判斷（例如「5日」→5 個一般代課節次），不會、也不能從
    // 日期區間本身推算出天數。目前只認得純日數格式（見 parseHoursOrDaysToPeriodCount
    // 開頭說明），無法安全解析時（缺漏、或像「3日4時」這種混合單位）不猜測，保留
    // RawRecord、標記「時數天數／待確認」，暫不建立 SubstituteRecord——跟日期區間
    // 本身無關，不可以誤植成日期相關的錯誤訊息。單一日期的既有資料完全不受影響
    // （periodCount 維持原本的空白，2026/06 迴歸不受任何影響）。
    var periodCount = null;
    if (parsedDate && parsedDate.isDateRange && periodCode && /^P[1-7]$/.test(periodCode)) {
      var hd = parseHoursOrDaysToPeriodCount(row.hoursOrDaysText);
      if (hd.error) {
        rowIssues.push({
          rowNumber: row.rowNumber, fieldName: "時數天數",
          message: '時數天數／待確認：日期為區間（"' + row.dateText + '"），節次「' + periodCode + '」需要明確的「時數天數」才能確認計費數量，' +
            (row.hoursOrDaysText ? '目前的文字："' + row.hoursOrDaysText + '"' : "目前欄位是空的") +
            "不是系統目前能安全辨識的格式（僅支援如「5日」這種純日數），暫不建立代課紀錄，請人工確認後改用可辨識的格式重新登錄",
        });
      } else {
        periodCount = hd.periodCount;
      }
    }

    if (rowIssues.length > 0 || !parsedDate || !periodCode) {
      issues = issues.concat(rowIssues);
      return;
    }

    var noteParts = [];
    if (parsedDate.weekdayMismatch) noteParts.push('日期文字標示的星期與實際計算不符（原始："' + row.dateText + '"）');
    if (parsedDate.monthMismatch) noteParts.push("此列日期月份與所選匯入月份（" + payload.month + "月）不同，請確認");
    if (parsedDate.isDateRange) {
      noteParts.push('日期原文為區間："' + row.dateText + '"，系統以區間起始日作為代課日期、不展開成多筆紀錄' +
        (periodCount !== null ? "；計費數量依時數天數設為 " + periodCount : ""));
    }

    recordRowsToInsert.push({
      id: newId(), rawRecordId: rawId, entryType: "EXCEL_IMPORT", monthlyImportId: monthlyImport.id,
      originalTeacherId: "", substituteTeacherId: "", date: parsedDate.date, weekday: parsedDate.weekday,
      periodCode: periodCode, className: row.className || "", subject: row.subject || "", leaveType: row.leaveType || "",
      rawHoursOrDays: row.hoursOrDaysText || "", periodCount: periodCount === null ? "" : periodCount, staffType: sourceStaffType,
      fundingSource: "UNDETERMINED", projectId: "", unitPrice: "", amount: "",
      classificationMethod: "GENERAL_DEFAULT", classificationRuleId: "", classifiedAt: "",
      autoFundingSource: "", autoClassificationMethod: "", autoClassificationRuleId: "", autoProjectId: "",
      conflictCandidatesJson: "", isManuallyModified: false, manualOverrideReason: "", manualOverrideAt: "",
      manualOverrideBy: "", note: noteParts.join("；"), createdBy: "", updatedBy: "",
      createdAt: nowIso(), updatedAt: nowIso(),
    });
    successCount += 1;
  });

  appendRows("RawRecords", rawRowsToInsert);
  appendRows("SubstituteRecords", recordRowsToInsert);
  if (issues.length > 0) {
    appendRows("ImportErrors", issues.map(function (i) {
      return { id: newId(), monthlyImportId: monthlyImport.id, rowNumber: i.rowNumber, fieldName: i.fieldName || "", message: i.message, createdAt: nowIso() };
    }));
  }

  var updatedImport = updateRow("MonthlyImports", monthlyImport.id, { successCount: successCount, errorCount: issues.length });

  return {
    monthlyImport: updatedImport, detectedHeaders: payload.detectedHeaders || [], totalCount: rows.length,
    successCount: successCount, errorCount: issues.length, errors: issues,
    supersededImportIds: previousActive.map(function (b) { return b.id; }),
  };
}

function api_listMonthlyImports(payload) {
  requireField(payload, "semesterId", "學期");
  var rows = readRows("MonthlyImports").filter(function (r) {
    if (r.semesterId !== payload.semesterId) return false;
    if (payload.year && Number(r.year) !== Number(payload.year)) return false;
    if (payload.month && Number(r.month) !== Number(payload.month)) return false;
    return true;
  }).map(stripRow);
  rows.sort(function (a, b) {
    if (a.year !== b.year) return b.year - a.year;
    if (a.month !== b.month) return b.month - a.month;
    return b.versionNo - a.versionNo;
  });
  return rows.map(hydrateMonthlyImport);
}

function hydrateMonthlyImport(r) {
  return {
    id: r.id, semesterId: r.semesterId, year: Number(r.year), month: Number(r.month), importedAt: r.importedAt,
    fileName: r.fileName, importedBy: r.importedBy, totalCount: Number(r.totalCount), successCount: Number(r.successCount),
    errorCount: Number(r.errorCount), versionNo: Number(r.versionNo), status: r.status, sourceStaffType: r.sourceStaffType,
    sourceSheetName: r.sourceSheetName, note: r.note, createdAt: r.createdAt,
  };
}

function api_getMonthlyImportDetail(payload) {
  requireField(payload, "id", "id");
  var detail = findById("MonthlyImports", payload.id);
  if (!detail) throw new Error("找不到匯入批次");
  var errors = readRows("ImportErrors").filter(function (e) { return e.monthlyImportId === payload.id; });
  errors.sort(function (a, b) { return Number(a.rowNumber || 0) - Number(b.rowNumber || 0); });
  var hydrated = hydrateMonthlyImport(detail);
  hydrated.errors = errors.map(stripRow);
  return hydrated;
}

function api_listSubstituteRecords(payload) {
  requireField(payload, "id", "id");
  var rows = readRows("SubstituteRecords").filter(function (r) { return r.monthlyImportId === payload.id; });
  rows.sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
  return rows.map(hydrateSubstituteRecord);
}

function getRawRecordRef(rawRecordId) {
  if (!rawRecordId) return null;
  var raw = findById("RawRecords", rawRecordId);
  return raw ? stripRow(raw) : null;
}

function hydrateSubstituteRecord(r) {
  var o = stripRow(r);
  o.periodCount = o.periodCount === "" || o.periodCount === null || o.periodCount === undefined ? null : Number(o.periodCount);
  o.unitPrice = o.unitPrice === "" ? null : decToStr(o.unitPrice);
  o.amount = o.amount === "" ? null : decToStr(o.amount);
  o.isManuallyModified = toBool(o.isManuallyModified);
  o.originalTeacherId = o.originalTeacherId || null;
  o.substituteTeacherId = o.substituteTeacherId || null;
  o.rawRecordId = o.rawRecordId || null;
  o.originalTeacher = getPersonRef(o.originalTeacherId);
  o.substituteTeacher = getPersonRef(o.substituteTeacherId);
  o.project = getProjectRef(o.projectId);
  o.rawRecord = getRawRecordRef(o.rawRecordId);
  return o;
}

function api_listUnmatchedTeacherReferences(payload) {
  requireField(payload, "id", "id");
  var records = readRows("SubstituteRecords").filter(function (r) {
    return r.monthlyImportId === payload.id && r.entryType === "EXCEL_IMPORT" && (!r.originalTeacherId || !r.substituteTeacherId);
  });
  var results = [];
  records.forEach(function (r) {
    var raw = getRawRecordRef(r.rawRecordId);
    if (!r.originalTeacherId && raw && raw.originalTeacherText) {
      results.push({ recordId: r.id, field: "original", rawName: raw.originalTeacherText, candidates: findPersonsByName(raw.originalTeacherText) });
    }
    if (!r.substituteTeacherId && raw && raw.substituteTeacherText) {
      results.push({ recordId: r.id, field: "substitute", rawName: raw.substituteTeacherText, candidates: findPersonsByName(raw.substituteTeacherText) });
    }
  });
  return results;
}

function findPersonsByName(name) {
  return readRows("Persons").filter(function (p) { return p.name === name; }).map(stripRow);
}

function api_resolveTeacherReference(payload) {
  requireField(payload, "id", "id");
  requireField(payload, "field", "field");
  requireField(payload, "personId", "personId");
  var record = findById("SubstituteRecords", payload.id);
  if (!record) throw new Error("找不到代課紀錄");
  assertImportMonthNotLocked(record.monthlyImportId);
  var fieldName = payload.field === "original" ? "originalTeacherId" : "substituteTeacherId";
  var oldValue = record[fieldName];
  var patch = {}; patch[fieldName] = payload.personId; patch.updatedAt = nowIso();
  var updated = updateRow("SubstituteRecords", payload.id, patch);
  writeChangeLog("substitute_records", payload.id, fieldName, oldValue, payload.personId, payload.changedBy, payload.reason || "人工配對教師");
  return hydrateSubstituteRecord(updated);
}

function api_autoApplyUnambiguousTeacherMatches(payload) {
  requireField(payload, "id", "id");
  var unmatched = api_listUnmatchedTeacherReferences(payload);
  var applied = 0;
  unmatched.forEach(function (u) {
    if (u.candidates.length === 1) {
      api_resolveTeacherReference({ id: u.recordId, field: u.field, personId: u.candidates[0].id, changedBy: payload.changedBy, reason: "系統依姓名找到唯一相符人員，經管理者一鍵套用" });
      applied += 1;
    }
  });
  var remaining = api_listUnmatchedTeacherReferences(payload);
  return { appliedCount: applied, remainingCount: remaining.length };
}
