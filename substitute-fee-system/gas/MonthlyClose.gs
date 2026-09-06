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

function toAckInfo(targetTable, targetId) {
  var ack = getAcknowledgement(targetTable, targetId);
  if (!ack) return { status: "PENDING", acknowledgement: null };
  return { status: "ACKNOWLEDGED", acknowledgement: { reason: ack.reason, acknowledgedBy: ack.acknowledgedBy, acknowledgedAt: ack.acknowledgedAt } };
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
    rows.push(Object.assign({
      issueType: "TEACHER_UNMATCHED", targetTable: "SubstituteRecord", targetId: r.id, date: r.date,
      originalTeacher: raw ? raw.originalTeacherText || null : null,
      substituteTeacher: sub ? sub.name : (raw ? raw.substituteTeacherText || null : null),
      periodCode: r.periodCode || null, className: r.className || null, subject: r.subject || null,
      description: "原教師姓名尚未配對到人員資料，無法判斷分類規則",
    }, toAckInfo("SubstituteRecord", r.id)));
  });

  conflict.forEach(function (r) {
    var raw = r.rawRecordId ? findById("RawRecords", r.rawRecordId) : null;
    var orig = getPersonRef(r.originalTeacherId);
    var sub = getPersonRef(r.substituteTeacherId);
    rows.push(Object.assign({
      issueType: "CONFLICT", targetTable: "SubstituteRecord", targetId: r.id, date: r.date,
      originalTeacher: orig ? orig.name : (raw ? raw.originalTeacherText || null : null),
      substituteTeacher: sub ? sub.name : (raw ? raw.substituteTeacherText || null : null),
      periodCode: r.periodCode || null, className: r.className || null, subject: r.subject || null,
      description: "同時符合多個規則，系統不會自動選一個，需要人工確認",
    }, toAckInfo("SubstituteRecord", r.id)));
  });

  amountMissing.forEach(function (r) {
    var raw = r.rawRecordId ? findById("RawRecords", r.rawRecordId) : null;
    var orig = getPersonRef(r.originalTeacherId);
    var sub = getPersonRef(r.substituteTeacherId);
    rows.push(Object.assign({
      issueType: "AMOUNT_MISSING", targetTable: "SubstituteRecord", targetId: r.id, date: r.date,
      originalTeacher: orig ? orig.name : (raw ? raw.originalTeacherText || null : null),
      substituteTeacher: sub ? sub.name : (raw ? raw.substituteTeacherText || null : null),
      periodCode: r.periodCode || null, className: r.className || null, subject: r.subject || null,
      description: "已分類為" + r.fundingSource + "，但原始資料沒有金額、也找不到生效中的費率，無法計算",
    }, toAckInfo("SubstituteRecord", r.id)));
  });

  importErrors.forEach(function (e) {
    rows.push(Object.assign({
      issueType: "IMPORT_ERROR", targetTable: "MonthlyImportError", targetId: e.id, date: null,
      originalTeacher: null, substituteTeacher: null, periodCode: null, className: null, subject: null,
      description: "第 " + (e.rowNumber || "?") + " 列" + (e.fieldName ? "（" + e.fieldName + "）" : "") + "：" + e.message,
    }, toAckInfo("MonthlyImportError", e.id)));
  });

  return rows;
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
