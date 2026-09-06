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

// 回傳 {date, weekday, weekdayMismatch, monthMismatch} 或 {error}
function parseDateText(text, year, expectedMonth) {
  var trimmed = String(text).trim();
  var match = trimmed.match(DATE_PATTERN);
  if (!match) return { error: '無法解析日期格式："' + text + '"，預期格式如 "06-18(四)" 或 "6/18(四) 13:50 ~ 15:50"' };
  var month = Number(match[1]);
  var day = Number(match[2]);
  if (month < 1 || month > 12) return { error: '日期月份不合理："' + text + '"' };
  var date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return { error: '日期不存在："' + text + '"' };
  var computedWeekday = WEEKDAY_BY_JS_INDEX[date.getUTCDay()];
  var textWeekdayChar = match[3];
  var weekdayMismatch = Boolean(textWeekdayChar) && WEEKDAY_CHAR_MAP[textWeekdayChar] !== computedWeekday;
  var monthMismatch = month !== expectedMonth;
  return { date: date.toISOString().slice(0, 10), weekday: computedWeekday, weekdayMismatch: weekdayMismatch, monthMismatch: monthMismatch };
}

// 回傳 {periodCode} 或 {error}
function parsePeriodText(text, validCodes) {
  var trimmed = String(text).trim();
  function check(code) {
    return validCodes[code] ? { periodCode: code } : { error: "節次代碼「" + code + "」尚未設定於系統（原始文字：\"" + text + "\"）" };
  }
  if (trimmed === "早自修") return check("EARLY_STUDY");
  if (trimmed === "午休") return check("LUNCH");
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
      if (dr.error) rowIssues.push({ rowNumber: row.rowNumber, fieldName: "日期", message: dr.error });
      else parsedDate = dr;
    }
    var periodCode = null;
    if (row.periodText) {
      var pr = parsePeriodText(row.periodText, validPeriodCodes);
      if (pr.error) rowIssues.push({ rowNumber: row.rowNumber, fieldName: "節次", message: pr.error });
      else periodCode = pr.periodCode;
    }

    if (rowIssues.length > 0 || !parsedDate || !periodCode) {
      issues = issues.concat(rowIssues);
      return;
    }

    var noteParts = [];
    if (parsedDate.weekdayMismatch) noteParts.push('日期文字標示的星期與實際計算不符（原始："' + row.dateText + '"）');
    if (parsedDate.monthMismatch) noteParts.push("此列日期月份與所選匯入月份（" + payload.month + "月）不同，請確認");

    recordRowsToInsert.push({
      id: newId(), rawRecordId: rawId, entryType: "EXCEL_IMPORT", monthlyImportId: monthlyImport.id,
      originalTeacherId: "", substituteTeacherId: "", date: parsedDate.date, weekday: parsedDate.weekday,
      periodCode: periodCode, className: row.className || "", subject: row.subject || "", leaveType: row.leaveType || "",
      rawHoursOrDays: row.hoursOrDaysText || "", periodCount: "", staffType: sourceStaffType,
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
