/**
 * MonthlyClose.gs — 月結守門、問題確認、自費代課、待處理清單、月結首頁。
 * 逐行對照 monthlyLockService.ts / issueAcknowledgementService.ts / selfFundedService.ts /
 * pendingIssuesService.ts / dashboardService.ts。
 *
 * 唯一跟 Node 版本不同的實作細節（刻意、已在 Setup.gs 說明）：自費代課不再借用
 * 「虛擬 MonthlyImport 容器」，直接使用使用者原本建立的 SelfFunded 分頁。因此這裡
 * 的 getBlockingIssues／getMonthlyDashboard／getChunaSummary 都是「SubstituteRecords
 * 只會有 EXCEL_IMPORT 資料、自費代課永遠只在 SelfFunded 分頁」，不需要另外用
 * entryType 排除自費代課——這跟 Node 版本排除的效果完全一致，只是資料放的位置不同。
 */

// ---------- 月結鎖定 ----------

function getActiveMonthlyImportIds(semesterId, year, month) {
  return readRows("MonthlyImports")
    .filter(function (r) { return r.semesterId === semesterId && Number(r.year) === year && Number(r.month) === month && r.status === "ACTIVE"; })
    .map(function (r) { return r.id; });
}

function getAcknowledgedIdSet(targetTable, targetIds) {
  var set = {};
  if (targetIds.length === 0) return set;
  var idSet = {};
  targetIds.forEach(function (id) { idSet[id] = true; });
  readRows("IssueAcknowledgements").forEach(function (a) {
    if (a.targetTable === targetTable && idSet[a.targetId]) set[a.targetId] = true;
  });
  return set;
}

// 計算某學期某年月「尚未被確認接受」的阻擋性問題數量。
// 只看 entryType=EXCEL_IMPORT 的紀錄——自費代課存在獨立的 SelfFunded 分頁，
// 天生就不會出現在這裡，不需要額外過濾。
function getBlockingIssues(semesterId, year, month) {
  var importIds = getActiveMonthlyImportIds(semesterId, year, month);
  if (importIds.length === 0) {
    return { teacherUnmatched: 0, conflict: 0, amountMissing: 0, importErrors: 0, total: 0 };
  }
  var importIdSet = {};
  importIds.forEach(function (id) { importIdSet[id] = true; });

  var allRecords = readRows("SubstituteRecords").filter(function (r) { return importIdSet[r.monthlyImportId] && r.entryType === "EXCEL_IMPORT"; });
  var teacherUnmatchedRows = allRecords.filter(function (r) { return r.classificationMethod === "TEACHER_UNMATCHED"; });
  var conflictRows = allRecords.filter(function (r) { return r.classificationMethod === "CONFLICT"; });
  var amountMissingRows = allRecords.filter(function (r) {
    return ["GENERAL", "OVERTIME", "PROJECT"].indexOf(r.fundingSource) !== -1 && (r.amount === "" || r.amount === null || r.amount === undefined);
  });
  var importErrorRows = readRows("ImportErrors").filter(function (e) { return importIdSet[e.monthlyImportId]; });

  var ackTeacherUnmatched = getAcknowledgedIdSet("SubstituteRecord", teacherUnmatchedRows.map(function (r) { return r.id; }));
  var ackConflict = getAcknowledgedIdSet("SubstituteRecord", conflictRows.map(function (r) { return r.id; }));
  var ackAmountMissing = getAcknowledgedIdSet("SubstituteRecord", amountMissingRows.map(function (r) { return r.id; }));
  var ackImportErrors = getAcknowledgedIdSet("MonthlyImportError", importErrorRows.map(function (r) { return r.id; }));

  var teacherUnmatched = teacherUnmatchedRows.filter(function (r) { return !ackTeacherUnmatched[r.id]; }).length;
  var conflict = conflictRows.filter(function (r) { return !ackConflict[r.id]; }).length;
  var amountMissing = amountMissingRows.filter(function (r) { return !ackAmountMissing[r.id]; }).length;
  var importErrors = importErrorRows.filter(function (r) { return !ackImportErrors[r.id]; }).length;

  return { teacherUnmatched: teacherUnmatched, conflict: conflict, amountMissing: amountMissing, importErrors: importErrors, total: teacherUnmatched + conflict + amountMissing + importErrors };
}

function getOrCreateLockRow(semesterId, year, month) {
  var existing = findOne("MonthlyLocks", function (r) { return Number(r.year) === year && Number(r.month) === month; });
  if (existing) return existing;
  return appendRow("MonthlyLocks", {
    id: newId(), semesterId: semesterId, year: year, month: month, status: "NOT_IMPORTED",
    lockedAt: "", lockedBy: "", unlockedAt: "", unlockedBy: "", note: "", createdAt: nowIso(), updatedAt: nowIso(),
  });
}

function getMonthlyLockStatus(semesterId, year, month) {
  var lock = getOrCreateLockRow(semesterId, year, month);
  return {
    year: Number(lock.year), month: Number(lock.month), status: lock.status, isLocked: lock.status === "LOCKED",
    lockedAt: lock.lockedAt || null, lockedBy: lock.lockedBy || null, unlockedAt: lock.unlockedAt || null, unlockedBy: lock.unlockedBy || null,
  };
}

// 給其他函式用的守門檢查：只看 (year, month) 是否已鎖定。找不到鎖定資料列時視為
// 「尚未鎖定」（NOT_IMPORTED 是預設狀態，本來就可以修改）。
function assertMonthNotLocked(year, month) {
  var lock = findOne("MonthlyLocks", function (r) { return Number(r.year) === year && Number(r.month) === month; });
  if (lock && lock.status === "LOCKED") {
    throw new Error(year + "年" + month + "月已經鎖定，不能修改。如需修改請先由管理者解除鎖定並留下理由。");
  }
}

function assertImportMonthNotLocked(monthlyImportId) {
  var monthlyImport = findById("MonthlyImports", monthlyImportId);
  if (!monthlyImport) throw new Error("找不到匯入批次");
  assertMonthNotLocked(Number(monthlyImport.year), Number(monthlyImport.month));
}

function assertRecordMonthNotLocked(substituteRecordId) {
  var record = findById("SubstituteRecords", substituteRecordId);
  if (!record) throw new Error("找不到代課紀錄");
  var monthlyImport = findById("MonthlyImports", record.monthlyImportId);
  if (!monthlyImport) throw new Error("找不到匯入批次");
  assertMonthNotLocked(Number(monthlyImport.year), Number(monthlyImport.month));
}

function lockMonth(semesterId, year, month, lockedBy) {
  if (isBlank(lockedBy)) throw new Error("鎖定必須填寫操作人");
  var issues = getBlockingIssues(semesterId, year, month);
  if (issues.total > 0) {
    throw new Error(
      "尚有 " + issues.total + " 筆未確認的阻擋問題無法鎖定：原教師未配對 " + issues.teacherUnmatched + " 筆、規則衝突 " + issues.conflict +
      " 筆、金額無法計算 " + issues.amountMissing + " 筆、匯入錯誤 " + issues.importErrors + " 筆。請先處理或在「待處理」頁面確認接受。"
    );
  }
  var lock = getOrCreateLockRow(semesterId, year, month);
  var oldStatus = lock.status;
  updateRow("MonthlyLocks", lock.id, { status: "LOCKED", lockedAt: nowIso(), lockedBy: lockedBy, unlockedAt: "", unlockedBy: "", updatedAt: nowIso() });
  writeChangeLog("monthly_locks", lock.id, "status", oldStatus, "LOCKED", lockedBy, "鎖定 " + year + "年" + month + "月");
  return getMonthlyLockStatus(semesterId, year, month);
}

function unlockMonth(semesterId, year, month, unlockedBy, reason) {
  if (isBlank(unlockedBy)) throw new Error("解鎖必須填寫操作人");
  if (isBlank(reason)) throw new Error("解鎖必須填寫理由");
  var lock = findOne("MonthlyLocks", function (r) { return Number(r.year) === year && Number(r.month) === month; });
  if (!lock || lock.status !== "LOCKED") throw new Error(year + "年" + month + "月目前不是鎖定狀態，不需要解鎖");
  updateRow("MonthlyLocks", lock.id, { status: "NOT_IMPORTED", unlockedAt: nowIso(), unlockedBy: unlockedBy, updatedAt: nowIso() });
  writeChangeLog("monthly_locks", lock.id, "status", "LOCKED", "NOT_IMPORTED", unlockedBy, "解鎖 " + year + "年" + month + "月：" + reason);
  return getMonthlyLockStatus(semesterId, year, month);
}

function api_getMonthlyLockStatus(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "year", "年");
  requireField(payload, "month", "月");
  return getMonthlyLockStatus(payload.semesterId, Number(payload.year), Number(payload.month));
}

function api_lockMonth(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "year", "年");
  requireField(payload, "month", "月");
  return lockMonth(payload.semesterId, Number(payload.year), Number(payload.month), payload.lockedBy);
}

function api_unlockMonth(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "year", "年");
  requireField(payload, "month", "月");
  requireField(payload, "reason", "解鎖理由");
  return unlockMonth(payload.semesterId, Number(payload.year), Number(payload.month), payload.unlockedBy, payload.reason);
}

// ---------- 月結問題確認 ----------
// 「已確認／接受」不是消除錯誤：只記錄「管理者已經看過這筆問題、同意讓它不再阻擋
// 這個月的月結」這件事。原始問題（SubstituteRecord 或 ImportError）完全不會被修改。

function getAcknowledgement(targetTable, targetId) {
  return findOne("IssueAcknowledgements", function (a) { return a.targetTable === targetTable && a.targetId === targetId; });
}

function api_acknowledgeIssue(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "year", "年");
  requireField(payload, "month", "月");
  requireField(payload, "targetTable", "targetTable");
  requireField(payload, "targetId", "targetId");
  requireField(payload, "reason", "理由");
  requireField(payload, "acknowledgedBy", "確認人");

  if (payload.targetTable === "SubstituteRecord") {
    if (!findById("SubstituteRecords", payload.targetId)) throw new Error("找不到這筆代課紀錄");
  } else if (payload.targetTable === "MonthlyImportError") {
    if (!findById("ImportErrors", payload.targetId)) throw new Error("找不到這筆匯入錯誤");
  } else {
    throw new Error("targetTable 只能是 SubstituteRecord 或 MonthlyImportError");
  }

  var existing = getAcknowledgement(payload.targetTable, payload.targetId);
  var ack;
  if (existing) {
    ack = updateRow("IssueAcknowledgements", existing.id, { reason: payload.reason, acknowledgedBy: payload.acknowledgedBy, acknowledgedAt: nowIso(), updatedAt: nowIso() });
  } else {
    ack = appendRow("IssueAcknowledgements", {
      id: newId(), semesterId: payload.semesterId, year: Number(payload.year), month: Number(payload.month),
      targetTable: payload.targetTable, targetId: payload.targetId, reason: payload.reason,
      acknowledgedBy: payload.acknowledgedBy, acknowledgedAt: nowIso(), createdAt: nowIso(), updatedAt: nowIso(),
    });
  }

  writeChangeLog("monthly_issue_acknowledgements", ack.id, existing ? "reason" : null, existing ? existing.reason : null, payload.reason, payload.acknowledgedBy,
    "確認接受問題（" + payload.targetTable + " " + payload.targetId + "）：" + payload.reason);
  return ack;
}

function api_revokeAcknowledgement(payload) {
  requireField(payload, "targetTable", "targetTable");
  requireField(payload, "targetId", "targetId");
  var existing = getAcknowledgement(payload.targetTable, payload.targetId);
  if (!existing) return { ok: true };
  deleteRow("IssueAcknowledgements", existing.id);
  writeChangeLog("monthly_issue_acknowledgements", existing.id, null, existing.reason, null, payload.changedBy,
    "撤銷確認（" + payload.targetTable + " " + payload.targetId + "），問題重新列為待處理");
  return { ok: true };
}

function api_listAcknowledgements(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "year", "年");
  requireField(payload, "month", "月");
  var rows = readRows("IssueAcknowledgements").filter(function (a) {
    return a.semesterId === payload.semesterId && Number(a.year) === Number(payload.year) && Number(a.month) === Number(payload.month);
  });
  rows.sort(function (a, b) { return String(b.acknowledgedAt).localeCompare(String(a.acknowledgedAt)); });
  return rows;
}

// ---------- 自費代課 ----------
// 不是 Excel 匯入資料，不跑分類引擎、不跑 Phase9-5 費用計算——金額由管理者建立時
// 直接輸入，避免污染 GENERAL/OVERTIME/PROJECT/CONFLICT 這一整套公費分類邏輯，
// 也不會被「待處理」清單誤判成需要處理的公費異常。獨立存在 SelfFunded 分頁
// （不是虛擬 MonthlyImport 容器，見檔案開頭說明）。

function api_createSelfFunded(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "year", "年");
  requireField(payload, "month", "月");
  requireField(payload, "date", "日期");
  requireField(payload, "substituteTeacherId", "代課教師");
  requireField(payload, "amount", "金額");
  requireField(payload, "createdBy", "建立人");
  assertMonthNotLocked(Number(payload.year), Number(payload.month));

  var dateOnly = toDateOnly(payload.date);
  var row = appendRow("SelfFunded", {
    id: newId(), semesterId: payload.semesterId, year: Number(payload.year), month: Number(payload.month),
    date: dateOnly, weekday: weekdayOfDateOnly(dateOnly),
    originalTeacherId: payload.originalTeacherId || "", substituteTeacherId: payload.substituteTeacherId,
    periodCode: payload.periodCode || "", className: payload.className || "", subject: payload.subject || "",
    amount: decToStr(Number(payload.amount)), unitPrice: decToStr(Number(payload.unitPrice !== undefined && payload.unitPrice !== "" ? payload.unitPrice : payload.amount)),
    note: payload.note || "", createdBy: payload.createdBy, createdAt: nowIso(), updatedBy: "", updatedAt: nowIso(),
  });

  writeChangeLog("self_funded", row.id, null, null, "建立自費代課：" + dateOnly + " 金額 " + row.amount, payload.createdBy, payload.note || "手動建立自費代課");
  return hydrateSelfFunded(row);
}

function api_updateSelfFunded(payload) {
  requireField(payload, "id", "id");
  requireField(payload, "updatedBy", "操作人");
  var existing = findById("SelfFunded", payload.id);
  if (!existing) throw new Error("找不到自費代課紀錄");
  assertMonthNotLocked(Number(existing.year), Number(existing.month));

  var patch = { updatedBy: payload.updatedBy, updatedAt: nowIso() };
  if (payload.date !== undefined) { patch.date = toDateOnly(payload.date); patch.weekday = weekdayOfDateOnly(patch.date); }
  if (payload.originalTeacherId !== undefined) patch.originalTeacherId = payload.originalTeacherId || "";
  if (payload.substituteTeacherId !== undefined) patch.substituteTeacherId = payload.substituteTeacherId;
  if (payload.periodCode !== undefined) patch.periodCode = payload.periodCode || "";
  if (payload.className !== undefined) patch.className = payload.className || "";
  if (payload.subject !== undefined) patch.subject = payload.subject || "";
  if (payload.amount !== undefined) patch.amount = decToStr(Number(payload.amount));
  if (payload.unitPrice !== undefined) patch.unitPrice = decToStr(Number(payload.unitPrice));
  if (payload.note !== undefined) patch.note = payload.note || "";

  var updated = updateRow("SelfFunded", payload.id, patch);
  writeChangeLog("self_funded", payload.id, null,
    JSON.stringify({ amount: existing.amount, date: existing.date }), JSON.stringify({ amount: updated.amount, date: updated.date }),
    payload.updatedBy, "修改自費代課");
  return hydrateSelfFunded(updated);
}

function api_deleteSelfFunded(payload) {
  requireField(payload, "id", "id");
  requireField(payload, "deletedBy", "操作人");
  requireField(payload, "reason", "理由");
  var existing = findById("SelfFunded", payload.id);
  if (!existing) throw new Error("找不到自費代課紀錄");
  assertMonthNotLocked(Number(existing.year), Number(existing.month));

  writeChangeLog("self_funded", payload.id, null, JSON.stringify({ amount: existing.amount, date: existing.date }), null, payload.deletedBy, "刪除自費代課：" + payload.reason);
  deleteRow("SelfFunded", payload.id);
  return { ok: true };
}

function api_listSelfFunded(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "year", "年");
  requireField(payload, "month", "月");
  var rows = readRows("SelfFunded").filter(function (r) {
    return r.semesterId === payload.semesterId && Number(r.year) === Number(payload.year) && Number(r.month) === Number(payload.month);
  });
  rows.sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
  return rows.map(hydrateSelfFunded);
}

function hydrateSelfFunded(r) {
  var o = stripRow(r);
  o.year = Number(o.year);
  o.month = Number(o.month);
  o.amount = decToStr(o.amount);
  o.unitPrice = o.unitPrice === "" ? null : decToStr(o.unitPrice);
  o.originalTeacherId = o.originalTeacherId || null;
  o.originalTeacher = getPersonRef(o.originalTeacherId);
  o.substituteTeacher = getPersonRef(o.substituteTeacherId);
  return o;
}

// ---------- 待處理工作區 ----------
// 把「未配對教師」「規則衝突」「金額算不出來」「匯入錯誤」整合成一個清單，
// 每一筆標示目前狀態：待處理 / 已確認接受（有 IssueAcknowledgement）。
//
// 可追蹤性補強：原本只有 IMPORT_ERROR 這一類完全沒有原教師／代課教師／日期／
// 節次／時數天數等資訊（ImportErrors 分頁本身只存 rowNumber/fieldName/message，
// 沒有存 rawRecordId），使用者只看得到「第34列｜時數天數｜時數天數／待確認…」，
// 完全不知道是哪一位老師。這裡不新增資料表，直接用「同一個 monthlyImportId +
// 同一個 rowNumber」去對照 RawRecords（RawRecords 本來就是每一列 Excel 無條件
// 都會建立一筆，rowNumber 在同一個匯入批次裡不會重複，兩邊用這組複合鍵配對是
// 可靠的），把原教師／代課教師／日期原文／節次原文／時數天數原文／rawRecordId
// 補齊到每一種 issueType，讓使用者一眼就能看到是哪一位老師、也能追溯回原始
// RawRecord。四種 issueType 現在共用同一組欄位形狀，不再是 IMPORT_ERROR 特別稀疏。

function toAckInfo(targetTable, targetId) {
  var ack = getAcknowledgement(targetTable, targetId);
  if (!ack) return { status: "PENDING", acknowledgement: null };
  return { status: "ACKNOWLEDGED", acknowledgement: { reason: ack.reason, acknowledgedBy: ack.acknowledgedBy, acknowledgedAt: ack.acknowledgedAt } };
}

// 依「原教師」姓名分組計算彙總時使用的顯示名稱：完全比照 api_listPendingIssues
// 每一筆 issue 的 originalTeacher 欄位（已配對的人員姓名，或原始 Excel 文字，
// 或都沒有時的 null），一律代換成同一個「未知原教師」字樣，確保彙總結果的
// key 跟畫面上顯示的內容完全一致。
var UNKNOWN_ORIGINAL_TEACHER_LABEL = "未知原教師";

// 把一筆 RawRecord 的「日期原文／節次原文／時數天數原文」補進 issue 物件裡，
// 四種 issueType 共用同一份邏輯，不必各自重複寫一次。
function attachRawRecordFields(issue, raw) {
  issue.rawRecordId = raw ? raw.id : null;
  issue.dateText = raw ? (raw.dateText || null) : null;
  issue.periodText = raw ? (raw.periodText || null) : null;
  issue.hoursOrDaysText = raw ? (raw.hoursOrDaysText || null) : null;
  return issue;
}

function api_listPendingIssues(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "year", "年");
  requireField(payload, "month", "月");
  var year = Number(payload.year), month = Number(payload.month);
  var batchIds = readRows("MonthlyImports")
    .filter(function (b) { return b.semesterId === payload.semesterId && Number(b.year) === year && Number(b.month) === month && b.status === "ACTIVE"; })
    .map(function (b) { return b.id; });
  if (batchIds.length === 0) return [];
  var batchIdSet = {};
  batchIds.forEach(function (id) { batchIdSet[id] = true; });

  // RawRecords 索引：key 是「monthlyImportId::rowNumber」，同一批匯入裡 rowNumber
  // 是 Excel 實際列號、彼此不重複，可以安全當唯一鍵使用，不需要另外存 rawRecordId
  // 在 ImportErrors 上（不新增欄位）。
  var rawByKey = {};
  readRows("RawRecords").filter(function (r) { return batchIdSet[r.monthlyImportId]; }).forEach(function (r) {
    rawByKey[r.monthlyImportId + "::" + String(r.rowNumber)] = r;
  });

  var validPeriodCodes = {};
  readRows("PeriodSlots").forEach(function (p) { validPeriodCodes[p.code] = true; });

  var allRecords = readRows("SubstituteRecords").filter(function (r) { return batchIdSet[r.monthlyImportId] && r.entryType === "EXCEL_IMPORT"; });
  var teacherUnmatched = allRecords.filter(function (r) { return r.classificationMethod === "TEACHER_UNMATCHED"; });
  var conflict = allRecords.filter(function (r) { return r.classificationMethod === "CONFLICT"; });
  var amountMissing = allRecords.filter(function (r) {
    return ["GENERAL", "OVERTIME", "PROJECT"].indexOf(r.fundingSource) !== -1 && (r.amount === "" || r.amount === null || r.amount === undefined);
  });
  var importErrors = readRows("ImportErrors").filter(function (e) { return batchIdSet[e.monthlyImportId]; });

  var rows = [];

  teacherUnmatched.forEach(function (r) {
    var raw = r.rawRecordId ? findById("RawRecords", r.rawRecordId) : null;
    var sub = getPersonRef(r.substituteTeacherId);
    rows.push(Object.assign(attachRawRecordFields({
      issueType: "TEACHER_UNMATCHED", targetTable: "SubstituteRecord", targetId: r.id, date: r.date,
      rowNumber: raw ? Number(raw.rowNumber) || null : null,
      originalTeacher: raw ? raw.originalTeacherText || null : null,
      substituteTeacher: sub ? sub.name : (raw ? raw.substituteTeacherText || null : null),
      periodCode: r.periodCode || null, className: r.className || null, subject: r.subject || null,
      fieldName: "原教師",
      description: "原教師姓名尚未配對到人員資料，無法判斷分類規則",
    }, raw), toAckInfo("SubstituteRecord", r.id)));
  });

  conflict.forEach(function (r) {
    var raw = r.rawRecordId ? findById("RawRecords", r.rawRecordId) : null;
    var orig = getPersonRef(r.originalTeacherId);
    var sub = getPersonRef(r.substituteTeacherId);
    rows.push(Object.assign(attachRawRecordFields({
      issueType: "CONFLICT", targetTable: "SubstituteRecord", targetId: r.id, date: r.date,
      rowNumber: raw ? Number(raw.rowNumber) || null : null,
      originalTeacher: orig ? orig.name : (raw ? raw.originalTeacherText || null : null),
      substituteTeacher: sub ? sub.name : (raw ? raw.substituteTeacherText || null : null),
      periodCode: r.periodCode || null, className: r.className || null, subject: r.subject || null,
      fieldName: "分類規則",
      description: "同時符合多個規則，系統不會自動選一個，需要人工確認",
    }, raw), toAckInfo("SubstituteRecord", r.id)));
  });

  amountMissing.forEach(function (r) {
    var raw = r.rawRecordId ? findById("RawRecords", r.rawRecordId) : null;
    var orig = getPersonRef(r.originalTeacherId);
    var sub = getPersonRef(r.substituteTeacherId);
    // 「時數天數待確認」造成的金額缺漏（見 Import.gs 的 PERIOD_COUNT_PENDING_MARKER／
    // FeeCalculation.gs 的 isPeriodCountPending）跟一般的「沒有費率可退回」是不同
    // 原因，用 note 裡的標記分辨，說明文字要清楚指出是時數天數的問題，不是金額
    // 或費率設定的問題。
    var isPeriodCountPending = Boolean(r.note) && r.note.indexOf(PERIOD_COUNT_PENDING_MARKER) !== -1;
    rows.push(Object.assign(attachRawRecordFields({
      issueType: "AMOUNT_MISSING", targetTable: "SubstituteRecord", targetId: r.id, date: r.date,
      rowNumber: raw ? Number(raw.rowNumber) || null : null,
      originalTeacher: orig ? orig.name : (raw ? raw.originalTeacherText || null : null),
      substituteTeacher: sub ? sub.name : (raw ? raw.substituteTeacherText || null : null),
      periodCode: r.periodCode || null, className: r.className || null, subject: r.subject || null,
      fieldName: isPeriodCountPending ? "時數天數" : "金額",
      description: isPeriodCountPending
        ? "時數天數無法安全解析（原文：" + (raw ? '"' + (raw.hoursOrDaysText || "") + '"' : "未知") + "），計費數量待確認，暫不計算一般代課鐘點費"
        : "已分類為" + r.fundingSource + "，但原始資料沒有金額、也找不到生效中的費率，無法計算",
    }, raw), toAckInfo("SubstituteRecord", r.id)));
  });

  importErrors.forEach(function (e) {
    var raw = rawByKey[e.monthlyImportId + "::" + String(e.rowNumber)] || null;
    var resolvedPeriodCode = null;
    if (raw && raw.periodText) {
      var pr = parsePeriodText(raw.periodText, validPeriodCodes);
      if (!pr.error) resolvedPeriodCode = pr.periodCode;
    }
    rows.push(Object.assign(attachRawRecordFields({
      issueType: "IMPORT_ERROR", targetTable: "MonthlyImportError", targetId: e.id, date: null,
      rowNumber: Number(e.rowNumber) || null,
      originalTeacher: raw ? (raw.originalTeacherText || null) : null,
      substituteTeacher: raw ? (raw.substituteTeacherText || null) : null,
      periodCode: resolvedPeriodCode, className: raw ? (raw.classText || null) : null, subject: raw ? (raw.subjectText || null) : null,
      fieldName: e.fieldName || null,
      description: "第 " + (e.rowNumber || "?") + " 列" + (e.fieldName ? "（" + e.fieldName + "）" : "") + "：" + e.message,
    }, raw), toAckInfo("MonthlyImportError", e.id)));
  });

  // 篩選：原教師／代課教師／問題類型／節次，都只在有傳值時才套用，完全不影響
  // 沒傳篩選條件時的既有回傳內容（一律回傳全部）。
  if (payload.originalTeacher) rows = rows.filter(function (r) { return r.originalTeacher === payload.originalTeacher; });
  if (payload.substituteTeacher) rows = rows.filter(function (r) { return r.substituteTeacher === payload.substituteTeacher; });
  if (payload.issueType) rows = rows.filter(function (r) { return r.issueType === payload.issueType; });
  if (payload.periodCode) rows = rows.filter(function (r) { return r.periodCode === payload.periodCode; });

  return rows;
}

// 依「原教師」彙總待處理問題筆數，供「待處理」頁面顯示「王○○：8筆」這種清單、
// 點選後可以只看該老師的問題。直接彙總 api_listPendingIssues() 的結果，不重新
// 查一次資料、不新增資料表；沒有原教師資訊的問題歸類到「未知原教師」，不會
// 因為缺資料就整筆消失不計。
function api_summarizePendingIssuesByOriginalTeacher(payload) {
  var issues = api_listPendingIssues(payload);
  var byTeacher = {};
  var order = [];
  issues.forEach(function (issue) {
    var name = issue.originalTeacher || UNKNOWN_ORIGINAL_TEACHER_LABEL;
    if (!byTeacher[name]) { byTeacher[name] = 0; order.push(name); }
    byTeacher[name] += 1;
  });
  return order
    .map(function (name) { return { originalTeacher: name, count: byTeacher[name] }; })
    .sort(function (a, b) { return b.count - a.count || String(a.originalTeacher).localeCompare(String(b.originalTeacher), "zh-Hant"); });
}

// ---------- 月結首頁 ----------
// 所有統計都是即時查詢既有資料算出來的，沒有另外存一份快照。

function api_getMonthlyDashboard(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "year", "年");
  requireField(payload, "month", "月");
  var year = Number(payload.year), month = Number(payload.month);

  var allBatches = readRows("MonthlyImports").filter(function (b) { return b.semesterId === payload.semesterId && Number(b.year) === year && Number(b.month) === month; });
  var activeBatches = allBatches.filter(function (b) { return b.status === "ACTIVE"; });
  var activeBatchIds = activeBatches.map(function (b) { return b.id; });
  var activeBatchIdSet = {};
  activeBatchIds.forEach(function (id) { activeBatchIdSet[id] = true; });

  var activeRecords = readRows("SubstituteRecords").filter(function (r) { return activeBatchIdSet[r.monthlyImportId] && r.entryType === "EXCEL_IMPORT"; });

  var classification = { general: 0, overtime: 0, project: 0, conflict: 0, teacherUnmatched: 0 };
  activeRecords.forEach(function (r) {
    if (r.classificationMethod === "CONFLICT") classification.conflict += 1;
    else if (r.classificationMethod === "TEACHER_UNMATCHED") classification.teacherUnmatched += 1;
    else if (r.fundingSource === "GENERAL") classification.general += 1;
    else if (r.fundingSource === "OVERTIME") classification.overtime += 1;
    else if (r.fundingSource === "PROJECT") classification.project += 1;
  });

  var calculatedCount = 0, notCalculatedCount = 0, totalAmount = "0";
  activeRecords.forEach(function (r) {
    if (r.amount !== "" && r.amount !== null && r.amount !== undefined) {
      calculatedCount += 1;
      totalAmount = decAdd(totalAmount, Number(r.amount));
    } else {
      notCalculatedCount += 1;
    }
  });

  var selfFundedRows = readRows("SelfFunded").filter(function (r) { return r.semesterId === payload.semesterId && Number(r.year) === year && Number(r.month) === month; });

  var blocking = getBlockingIssues(payload.semesterId, year, month);
  var acknowledgements = readRows("IssueAcknowledgements").filter(function (a) { return a.semesterId === payload.semesterId && Number(a.year) === year && Number(a.month) === month; });
  var lock = getMonthlyLockStatus(payload.semesterId, year, month);

  return {
    year: year, month: month,
    import: {
      batches: activeBatches.map(function (b) { return { id: b.id, sourceStaffType: b.sourceStaffType, versionNo: Number(b.versionNo), fileName: b.fileName, totalCount: Number(b.totalCount), successCount: Number(b.successCount), errorCount: Number(b.errorCount) }; }),
      hasActiveBatch: activeBatches.length > 0,
      successCount: activeBatches.reduce(function (s, b) { return s + Number(b.successCount); }, 0),
      errorCount: activeBatches.reduce(function (s, b) { return s + Number(b.errorCount); }, 0),
    },
    selfFunded: { exists: selfFundedRows.length > 0, count: selfFundedRows.length },
    classification: classification,
    fee: { calculatedCount: calculatedCount, notCalculatedCount: notCalculatedCount, totalAmount: decToStr(totalAmount) },
    issues: { blocking: blocking, acknowledgedCount: acknowledgements.length },
    lock: lock,
  };
}
