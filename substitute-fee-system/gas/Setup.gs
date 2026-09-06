/**
 * Setup.gs — Google Sheets 資料表結構定義與一次性初始化。
 *
 * 這裡把 Node/Prisma 版本的資料模型「原封不動」對應成 Sheets 分頁＋欄位，
 * 沒有重新設計欄位或拿掉任何欄位。使用者原本已經建立的 14 個分頁
 * （Persons/Semesters/SubstituteRecords/RawRecords/MonthlyImports/WeeklyRules/
 * DateRules/Projects/FeeRules/SelfFunded/FeeCalculations/Reconciliation/
 * MonthlyLocks/ChangeLog）全部保留、欄位補齊；另外補上幾個 Node 版本本來就有、
 * 但使用者當初沒有另外開分頁的輔助資料表（ImportErrors／IssueAcknowledgements／
 * PeriodSlots／BdClassificationRules），純粹是「資料表放在哪一個分頁」的實作細節，
 * 不影響任何欄位、任何分類/計費規則。
 *
 * 只有一個地方跟 Node 版本的實體關聯稍微不同，且是刻意的、合理的調整：
 * 自費代課在 Node 版本裡借用了 SubstituteRecord + 一個「虛擬 MonthlyImport 容器」
 * 來繞過「monthlyImportId 不可為 null」的關聯式資料庫限制；Sheets 沒有外鍵限制，
 * 不需要這個繞路，所以自費代課直接使用使用者原本就建立好的獨立 SelfFunded 分頁。
 * 欄位、業務邏輯（不進分類引擎、不算入公費總額、給出納/對帳都要排除）完全比照
 * Node 版本的自費代課規則，只是不再需要那個「虛擬容器」的實作技巧。
 *
 * FeeCalculations／Reconciliation 這兩個使用者事先建立的分頁，Node 版本裡並不存在
 * 對應資料表（費用直接存在 SubstituteRecords.amount／unitPrice 上、對帳是即時運算、
 * 不落地保存——這是 Phase 9 明確定的原則：「月結首頁統計必須即時運算，不能存重複快照」）。
 * 這裡沒有違反這個原則：SubstituteRecords.amount/unitPrice 仍然是唯一的金額來源，
 * 給出納/對帳/分類統計仍然每次即時從 SubstituteRecords 現算；FeeCalculations／
 * Reconciliation 兩個分頁只拿來做「附加的稽核紀錄」（每次執行計算/對帳時多寫一筆
 * 歷史記錄，方便事後回溯「這次是誰在什麼時候跑的」），不是計算的資料來源，
 * 拿掉這兩個分頁不會讓任何功能算錯，純粹只是少了歷史記錄可以查。
 */

var SHEET_SCHEMAS = {
  // ---- 使用者已建立的 14 個分頁 ----
  Semesters: [
    "id", "schoolYear", "term", "status", "isCurrent", "startDate", "endDate",
    "overtimeMatchMode", "note", "createdAt", "updatedAt",
  ],
  Persons: [
    "id", "name", "payrollCode", "enrollmentStatus", "enrollDate", "leaveDate",
    "note", "createdAt", "updatedAt",
  ],
  Projects: [
    "id", "semesterId", "name", "isActive", "note", "createdAt", "updatedAt",
  ],
  WeeklyRules: [
    "id", "semesterId", "personId", "ruleType", "projectId", "weekday", "periodCode",
    "subject", "weeklyPeriods", "effectiveDate", "endDate", "note", "createdAt", "updatedAt",
  ],
  DateRules: [
    "id", "semesterId", "date", "personId", "periodCode", "originalClassificationNote",
    "overrideClassification", "projectId", "note", "isCancelled", "cancelledAt",
    "cancelledBy", "createdAt", "updatedAt",
  ],
  FeeRules: [
    "id", "semesterId", "feeType", "amount", "effectiveDate", "endDate", "note",
    "createdAt", "updatedAt",
  ],
  MonthlyImports: [
    "id", "semesterId", "year", "month", "importedAt", "fileName", "importedBy",
    "totalCount", "successCount", "errorCount", "versionNo", "status",
    "sourceStaffType", "sourceSheetName", "note", "createdAt",
  ],
  RawRecords: [
    "id", "monthlyImportId", "rowNumber", "originalTeacherCodeText", "originalTeacherText",
    "dateText", "leaveTypeText", "hoursOrDaysText", "periodText", "classText", "subjectText",
    "substituteTeacherCodeText", "substituteTeacherText", "teacherCertText", "payGradeText",
    "homeroomFeeText", "dailyOrHalfDayWageText", "substitutePeriodFeeText", "periodCountText",
    "sheetName", "rawJson", "processedRecordId", "createdAt",
  ],
  SubstituteRecords: [
    "id", "rawRecordId", "entryType", "monthlyImportId", "originalTeacherId", "substituteTeacherId",
    "date", "weekday", "periodCode", "className", "subject", "leaveType", "rawHoursOrDays",
    "periodCount", "staffType", "fundingSource", "projectId", "unitPrice", "amount",
    "classificationMethod", "classificationRuleId", "classifiedAt",
    "autoFundingSource", "autoClassificationMethod", "autoClassificationRuleId", "autoProjectId",
    "conflictCandidatesJson", "isManuallyModified", "manualOverrideReason", "manualOverrideAt",
    "manualOverrideBy", "note", "createdBy", "updatedBy", "createdAt", "updatedAt",
  ],
  SelfFunded: [
    "id", "semesterId", "year", "month", "date", "weekday", "originalTeacherId",
    "substituteTeacherId", "periodCode", "className", "subject", "amount", "unitPrice",
    "note", "createdBy", "createdAt", "updatedBy", "updatedAt",
  ],
  FeeCalculations: [
    // 附加稽核紀錄用（見上方檔案說明），不是金額的來源。
    "id", "recordId", "recordSource", "unitPrice", "amount", "feeType", "feeRuleId",
    "reason", "calculatedAt", "changedBy",
  ],
  Reconciliation: [
    // 每次「對帳」執行後的逐教師結果快照，供事後查歷史用；即時比對邏輯仍每次現算。
    "id", "runId", "monthlyImportIds", "name", "systemPeriodCount", "originalPeriodCount",
    "systemAmount", "originalAmount", "amountDiff", "status", "possibleReason", "createdAt",
    "createdBy",
  ],
  MonthlyLocks: [
    "id", "semesterId", "year", "month", "status", "lockedAt", "lockedBy",
    "unlockedAt", "unlockedBy", "note", "createdAt", "updatedAt",
  ],
  ChangeLog: [
    "id", "tableName", "recordId", "fieldName", "oldValue", "newValue",
    "changedAt", "changedBy", "reason", "createdAt",
  ],

  // ---- Node 版本本來就有、使用者沒另外開分頁的輔助資料表（自動建立） ----
  ImportErrors: [
    "id", "monthlyImportId", "rowNumber", "fieldName", "message", "createdAt",
  ],
  IssueAcknowledgements: [
    "id", "semesterId", "year", "month", "targetTable", "targetId", "reason",
    "acknowledgedBy", "acknowledgedAt", "createdAt", "updatedAt",
  ],
  PeriodSlots: [
    "id", "code", "displayName", "sortOrder", "isTeachingPeriod", "createdAt", "updatedAt",
  ],
  BdClassificationRules: [
    "id", "codeValue", "isBd", "note", "createdAt", "updatedAt",
  ],
  SchoolCalendarDays: [
    "id", "semesterId", "date", "weekday", "isTeachingDay", "note", "createdAt", "updatedAt",
  ],
};

/**
 * 一次性初始化：到 Apps Script 編輯器選這個函式、按執行即可。
 * 可以重複執行——已經存在的分頁只會補齊表頭（不會清空既有資料），
 * PeriodSlots／BdClassificationRules 只在完全空白時才寫入預設值。
 */
function setupSheets() {
  var ss = getSpreadsheet();
  Object.keys(SHEET_SCHEMAS).forEach(function (name) {
    var headers = SHEET_SCHEMAS[name];
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    var existing = headerRange.getValues()[0];
    var needsHeader = headers.some(function (h, i) { return existing[i] !== h; });
    if (needsHeader) {
      headerRange.setValues([headers]);
    }
    // 全部欄位一律用純文字格式，避免 Sheets 自動把日期字串／長數字轉成序列值，
    // 這對「Raw 金額 0 不可視為缺漏」這種必須逐字保留原文的規則格外重要。
    sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2000), headers.length).setNumberFormat("@");
    if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);
  });

  seedPeriodSlotsIfEmpty();
  seedBdClassificationRulesIfEmpty();

  Logger.log("setupSheets 完成，共 " + Object.keys(SHEET_SCHEMAS).length + " 個分頁。");
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function seedPeriodSlotsIfEmpty() {
  var rows = readRows("PeriodSlots");
  if (rows.length > 0) return;
  var defs = [
    { code: "EARLY_STUDY", displayName: "早自修", sortOrder: 1, isTeachingPeriod: false },
    { code: "P1", displayName: "第1節", sortOrder: 2, isTeachingPeriod: true },
    { code: "P2", displayName: "第2節", sortOrder: 3, isTeachingPeriod: true },
    { code: "P3", displayName: "第3節", sortOrder: 4, isTeachingPeriod: true },
    { code: "P4", displayName: "第4節", sortOrder: 5, isTeachingPeriod: true },
    { code: "LUNCH", displayName: "午休", sortOrder: 6, isTeachingPeriod: false },
    { code: "P5", displayName: "第5節", sortOrder: 7, isTeachingPeriod: true },
    { code: "P6", displayName: "第6節", sortOrder: 8, isTeachingPeriod: true },
    { code: "P7", displayName: "第7節", sortOrder: 9, isTeachingPeriod: true },
  ];
  defs.forEach(function (d) {
    appendRow("PeriodSlots", {
      id: newId(), code: d.code, displayName: d.displayName, sortOrder: d.sortOrder,
      isTeachingPeriod: d.isTeachingPeriod, createdAt: nowIso(), updatedAt: nowIso(),
    });
  });
}

function seedBdClassificationRulesIfEmpty() {
  var rows = readRows("BdClassificationRules");
  if (rows.length > 0) return;
  ["B", "D"].forEach(function (code) {
    appendRow("BdClassificationRules", {
      id: newId(), codeValue: code, isBd: true, note: "已確認：代碼 B/D → 編制內",
      createdAt: nowIso(), updatedAt: nowIso(),
    });
  });
}
