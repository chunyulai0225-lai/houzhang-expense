/**
 * CoreEntities.gs — 學期／人員／節次代碼／專案／每週固定規則／單日例外／費用規則／
 * 學校上課日曆。這些都是「資料維護」性質的 CRUD，邏輯逐一比照 Node 版本對應的
 * service（projectService.ts／specialWeeklyRuleService.ts／specialDateRuleService.ts／
 * feeRuleService.ts／schoolCalendarService.ts／personService.ts）。
 *
 * 回傳形狀刻意跟 Prisma include 出來的巢狀物件一致（例如 rule.person = {id,name}、
 * rule.project = {id,name}），這樣前端 app.js 完全不用改渲染邏輯，只需要換掉
 * fetch 的目標。
 */

// ---------- 學期管理 ----------
// 學期狀態是一個小狀態機，三個值：
//   NOT_STARTED（已建立、尚未被設為目前使用）── 新增學期的預設狀態
//   ACTIVE     （目前使用中）── 只透過 api_setCurrentSemester() 進入，isCurrent 必為 true
//   INACTIVE   （已停用）── 只透過 api_deactivateSemester() 進入，isCurrent 必為 false
// isCurrent 全系統只允許同時一個 true，唯一改變它的地方是 api_setCurrentSemester()；
// 新增／編輯／停用／重新啟用都不會把 isCurrent 改成 true，避免意外造成兩個 current。
// 「重新啟用」只是把 INACTIVE 解除回 NOT_STARTED（可以再被設為目前使用），
// 不會自動變成目前使用中——那一定要管理者另外明確按「設為目前使用」。

function hydrateSemester(r) {
  return {
    id: r.id, schoolYear: Number(r.schoolYear), term: Number(r.term), status: r.status,
    isCurrent: toBool(r.isCurrent),
    startDate: r.startDate, endDate: r.endDate, overtimeMatchMode: r.overtimeMatchMode,
    note: r.note || "", createdAt: r.createdAt, updatedAt: r.updatedAt,
  };
}

function api_listSemesters() {
  var rows = readRows("Semesters").map(stripRow);
  rows.sort(function (a, b) {
    if (a.schoolYear !== b.schoolYear) return Number(b.schoolYear) - Number(a.schoolYear);
    return Number(b.term) - Number(a.term);
  });
  return rows.map(hydrateSemester);
}

function validateSemesterTerm(term) {
  var t = Number(term);
  if (t !== 1 && t !== 2) throw new Error("學期只能是第1學期或第2學期");
  return t;
}

// 對照 Node schema.prisma 的 @@unique([schoolYear, term])：同一個學年度＋學期只能有一筆。
function assertSemesterYearTermUnique(schoolYear, term, excludeId) {
  var dup = findOne("Semesters", function (r) {
    return Number(r.schoolYear) === schoolYear && Number(r.term) === term && r.id !== excludeId;
  });
  if (dup) throw new Error(schoolYear + "學年度第" + term + "學期已經存在，不能重複建立");
}

function validateOvertimeMatchMode(mode) {
  if (mode !== "TEACHER_WEEKDAY_PERIOD" && mode !== "TEACHER_WEEKDAY_PERIOD_SUBJECT") {
    throw new Error("加班比對模式只能是 TEACHER_WEEKDAY_PERIOD 或 TEACHER_WEEKDAY_PERIOD_SUBJECT");
  }
}

function api_createSemester(payload) {
  requireField(payload, "schoolYear", "學年度");
  requireField(payload, "term", "學期");
  requireField(payload, "startDate", "開始日期");
  requireField(payload, "endDate", "結束日期");
  var schoolYear = Number(payload.schoolYear);
  var term = validateSemesterTerm(payload.term);
  assertSemesterYearTermUnique(schoolYear, term, null);
  var startDate = toDateOnly(payload.startDate);
  var endDate = toDateOnly(payload.endDate);
  validateDateRange(startDate, endDate);
  // 目前系統實際在用的預設值是 SUBJECT（比對加班規則時同時比對科目），沒填的話就用這個，
  // 但仍然是可以在新增/編輯時改成 TEACHER_WEEKDAY_PERIOD 的一般欄位，不是寫死的常數。
  var overtimeMatchMode = payload.overtimeMatchMode || "TEACHER_WEEKDAY_PERIOD_SUBJECT";
  validateOvertimeMatchMode(overtimeMatchMode);

  var semester = {
    id: newId(), schoolYear: schoolYear, term: term, status: "NOT_STARTED", isCurrent: false,
    startDate: startDate, endDate: endDate, overtimeMatchMode: overtimeMatchMode,
    note: payload.note || "", createdAt: nowIso(), updatedAt: nowIso(),
  };
  appendRow("Semesters", semester);
  writeChangeLog("semesters", semester.id, null, null,
    schoolYear + "學年度第" + term + "學期（" + startDate + " ~ " + endDate + "）",
    payload.changedBy, payload.reason || "新增學期");
  return hydrateSemester(semester);
}

// 可修改：學年度／學期／開始日期／結束日期／overtimeMatchMode／note。
// 不可修改：id、createdAt。不會動 status／isCurrent —— 就算改的是目前使用中的學期，
// 也不會因為編輯而自動取消 current（跟停用是兩件不同的事，只有 deactivate 會動 isCurrent）。
function api_updateSemester(payload) {
  requireField(payload, "id", "id");
  var existing = findById("Semesters", payload.id);
  if (!existing) throw new Error("找不到這個學期");

  var patch = { updatedAt: nowIso() };
  var nextSchoolYear = existing.schoolYear, nextTerm = existing.term;
  if (payload.schoolYear !== undefined) { nextSchoolYear = Number(payload.schoolYear); patch.schoolYear = nextSchoolYear; }
  if (payload.term !== undefined) { nextTerm = validateSemesterTerm(payload.term); patch.term = nextTerm; }
  if (payload.schoolYear !== undefined || payload.term !== undefined) {
    assertSemesterYearTermUnique(Number(nextSchoolYear), Number(nextTerm), existing.id);
  }

  var nextStartDate = existing.startDate, nextEndDate = existing.endDate;
  if (payload.startDate !== undefined) { nextStartDate = toDateOnly(payload.startDate); patch.startDate = nextStartDate; }
  if (payload.endDate !== undefined) { nextEndDate = toDateOnly(payload.endDate); patch.endDate = nextEndDate; }
  if (payload.startDate !== undefined || payload.endDate !== undefined) {
    validateDateRange(nextStartDate, nextEndDate);
  }

  if (payload.overtimeMatchMode !== undefined) {
    validateOvertimeMatchMode(payload.overtimeMatchMode);
    patch.overtimeMatchMode = payload.overtimeMatchMode;
  }
  if (payload.note !== undefined) patch.note = payload.note || "";

  var updated = updateRow("Semesters", payload.id, patch);
  writeChangeLog("semesters", payload.id, null,
    existing.schoolYear + "學年度第" + existing.term + "學期（" + existing.startDate + " ~ " + existing.endDate + "）",
    updated.schoolYear + "學年度第" + updated.term + "學期（" + updated.startDate + " ~ " + updated.endDate + "）",
    payload.changedBy, payload.reason || "修改學期資料");
  return hydrateSemester(updated);
}

// 設為目前使用：全系統同時只能有一個 isCurrent=true。
// 1) 其他所有目前是 current 的學期，isCurrent 改 false；2) 指定學期 isCurrent 改 true；
// 3) 指定學期 status 改 ACTIVE；4) 兩邊都寫 updatedAt；5) 每一筆被動到的學期都寫 ChangeLog。
function api_setCurrentSemester(payload) {
  requireField(payload, "id", "id");
  var target = findById("Semesters", payload.id);
  if (!target) throw new Error("找不到這個學期");

  var previousCurrents = readRows("Semesters").filter(function (r) { return r.id !== payload.id && toBool(r.isCurrent); });
  previousCurrents.forEach(function (r) {
    updateRow("Semesters", r.id, { isCurrent: false, updatedAt: nowIso() });
    writeChangeLog("semesters", r.id, "isCurrent", "true", "false", payload.changedBy,
      payload.reason || ("改由 " + target.schoolYear + "學年度第" + target.term + "學期 設為目前使用"));
  });

  var updated = updateRow("Semesters", payload.id, { isCurrent: true, status: "ACTIVE", updatedAt: nowIso() });
  writeChangeLog("semesters", payload.id, "isCurrent", "false", "true", payload.changedBy, payload.reason || "設為目前使用學期");
  return hydrateSemester(updated);
}

// 停用：status=INACTIVE、isCurrent=false，不刪除資料，歷史紀錄查詢照常可以查到這筆學期。
// 若停用的正是目前使用中的學期，前端會先跳出確認對話框，這裡仍然照樣執行（後端不重複擋，
// 「確認」是操作流程上的一道關卡，不是資料規則本身）。
function api_deactivateSemester(payload) {
  requireField(payload, "id", "id");
  var existing = findById("Semesters", payload.id);
  if (!existing) throw new Error("找不到這個學期");
  var wasCurrent = toBool(existing.isCurrent);

  var updated = updateRow("Semesters", payload.id, { status: "INACTIVE", isCurrent: false, updatedAt: nowIso() });
  writeChangeLog("semesters", payload.id, "status", existing.status, "INACTIVE", payload.changedBy,
    payload.reason || (wasCurrent ? "停用目前使用中的學期" : "停用學期"));
  if (wasCurrent) {
    writeChangeLog("semesters", payload.id, "isCurrent", "true", "false", payload.changedBy, "停用學期，連帶取消目前使用中狀態");
  }
  return hydrateSemester(updated);
}

// 重新啟用：只解除 INACTIVE 狀態、回到 NOT_STARTED，不會自動變成目前使用中——
// 要設回目前使用中，仍然要另外呼叫 api_setCurrentSemester()。
function api_activateSemester(payload) {
  requireField(payload, "id", "id");
  var existing = findById("Semesters", payload.id);
  if (!existing) throw new Error("找不到這個學期");
  if (existing.status !== "INACTIVE") throw new Error("這個學期目前不是停用狀態，不需要重新啟用");

  var updated = updateRow("Semesters", payload.id, { status: "NOT_STARTED", updatedAt: nowIso() });
  writeChangeLog("semesters", payload.id, "status", "INACTIVE", "NOT_STARTED", payload.changedBy, payload.reason || "重新啟用學期");
  return hydrateSemester(updated);
}

// ---------- 人員／節次代碼 ----------

function getPersonRef(personId) {
  if (!personId) return null;
  var p = findById("Persons", personId);
  return p ? { id: p.id, name: p.name, payrollCode: p.payrollCode } : null;
}

function getProjectRef(projectId) {
  if (!projectId) return null;
  var p = findById("Projects", projectId);
  return p ? { id: p.id, name: p.name } : null;
}

function api_listPersons(payload) {
  var search = payload && payload.search;
  var rows = readRows("Persons").map(stripRow);
  if (search) {
    rows = rows.filter(function (p) { return p.name && p.name.indexOf(search) !== -1; });
  }
  return rows;
}

function api_listPeriodSlots() {
  var rows = readRows("PeriodSlots").map(stripRow);
  rows.sort(function (a, b) { return Number(a.sortOrder) - Number(b.sortOrder); });
  return rows;
}

// ---------- 專案 ----------

// 專案是否已經被其他資料引用。實際檢查過 Setup.gs 的 SHEET_SCHEMAS：真正存有
// projectId 欄位、會參照到 Projects 的只有 WeeklyRules／DateRules／SubstituteRecords
// 這三張表——FeeCalculations／Reconciliation 是附加稽核紀錄（見 Setup.gs 開頭說明），
// 欄位裡沒有 projectId，本來就不會直接參照專案；如果某筆代課紀錄的費用計算被記錄進
// FeeCalculations，那筆代課紀錄本身在 SubstituteRecords 裡的 projectId 已經會被這裡
// 的 SubstituteRecords 檢查涵蓋到，不需要另外重複檢查那兩張稽核表。
//
// 效能筆記（判斷結果完全不變，只是不要為了列出 N 個專案就重複掃 3 張表 N 次）：
// api_listProjects() 原本對每一列專案各自呼叫一次 isProjectInUse()，等於掃了
// WeeklyRules／DateRules／SubstituteRecords 各 N 次（N＝專案數）；SubstituteRecords
// 隨著每個月匯入只會越來越大，這個 N 次重複掃描是「載入專案清單」感覺變慢的主要
// 原因之一。getUsedProjectIdSet() 只掃這三張表各一次、建一份「有被引用的 projectId」
// 對照表，之後不管檢查幾個專案都是 O(1) 查表，跟原本「一個一個查」得到的結果完全一樣。
function getUsedProjectIdSet() {
  var used = {};
  readRows("WeeklyRules").forEach(function (r) { if (r.projectId) used[r.projectId] = true; });
  readRows("DateRules").forEach(function (r) { if (r.projectId) used[r.projectId] = true; });
  readRows("SubstituteRecords").forEach(function (r) { if (r.projectId) used[r.projectId] = true; });
  return used;
}

function isProjectInUse(projectId) {
  return !!getUsedProjectIdSet()[projectId];
}

function api_listProjects(payload) {
  requireField(payload, "semesterId", "學期");
  var rows = readRows("Projects").filter(function (r) { return r.semesterId === payload.semesterId; }).map(stripRow);
  var usedProjectIds = getUsedProjectIdSet(); // 掃一次，下面每一列查表就好，不要各自再掃三張表
  return rows.map(function (r) {
    return {
      id: r.id, semesterId: r.semesterId, name: r.name, isActive: toBool(r.isActive), note: r.note,
      createdAt: r.createdAt, updatedAt: r.updatedAt, isInUse: !!usedProjectIds[r.id],
    };
  });
}

function api_createProject(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "name", "專案名稱");
  var project = {
    id: newId(), semesterId: payload.semesterId, name: payload.name.trim(), isActive: true,
    note: payload.note || "", createdAt: nowIso(), updatedAt: nowIso(),
  };
  appendRow("Projects", project);
  writeChangeLog("projects", project.id, null, null, project.name, payload.changedBy, "新增專案");
  return project;
}

function api_updateProject(payload) {
  requireField(payload, "id", "id");
  var existing = findById("Projects", payload.id);
  if (!existing) throw new Error("找不到專案");
  var patch = { updatedAt: nowIso() };
  if (payload.name !== undefined) patch.name = payload.name;
  if (payload.note !== undefined) patch.note = payload.note;
  var updated = updateRow("Projects", payload.id, patch);
  if (payload.name !== undefined && payload.name !== existing.name) {
    writeChangeLog("projects", payload.id, "name", existing.name, payload.name, payload.changedBy, payload.reason || "修改專案");
  }
  return updated;
}

function api_setProjectActive(payload) {
  requireField(payload, "id", "id");
  var existing = findById("Projects", payload.id);
  if (!existing) throw new Error("找不到專案");
  var isActive = Boolean(payload.isActive);
  var updated = updateRow("Projects", payload.id, { isActive: isActive, updatedAt: nowIso() });
  writeChangeLog("projects", payload.id, "isActive", String(toBool(existing.isActive)), String(isActive), payload.changedBy, payload.reason || "切換啟用狀態");
  return updated;
}

// 刪除是實體刪除（從 Projects 移除那一列），只有完全沒有被引用的專案才允許——
// 一旦被 WeeklyRules／DateRules／SubstituteRecords 任何一筆引用過，一律拒絕，
// 請使用者改用「停用」。不做 cascade delete，也不會動到任何引用它的歷史資料本身
// （這裡只是檢查、拒絕，從來不會去改寫或刪除 SubstituteRecords 等其他表）。
function api_deleteProject(payload) {
  requireField(payload, "id", "id");
  var existing = findById("Projects", payload.id);
  if (!existing) throw new Error("找不到專案");
  if (isProjectInUse(payload.id)) {
    throw new Error("此專案已被資料使用，無法刪除。若不再使用，請改為停用。");
  }
  deleteRow("Projects", payload.id);
  writeChangeLog("projects", payload.id, null, existing.name, null, payload.changedBy, payload.reason || "刪除專案（未被任何資料引用）");
  return { ok: true, id: payload.id };
}

// ---------- 每週固定規則 ----------

function rangesOverlap(aStart, aEnd, bStart, bEnd) {
  var aEndT = aEnd ? new Date(aEnd).getTime() : Infinity;
  var bEndT = bEnd ? new Date(bEnd).getTime() : Infinity;
  return new Date(aStart).getTime() <= bEndT && new Date(bStart).getTime() <= aEndT;
}

function detectWeeklyRuleConflicts(semesterId, personId) {
  var rules = readRows("WeeklyRules").filter(function (r) {
    return r.semesterId === semesterId && (!personId || r.personId === personId);
  });
  var groups = {};
  rules.forEach(function (r) {
    var key = r.personId + "|" + r.weekday + "|" + r.periodCode;
    (groups[key] = groups[key] || []).push(r);
  });
  var conflicts = [];
  Object.keys(groups).forEach(function (key) {
    var list = groups[key];
    if (list.length < 2) return;
    for (var i = 0; i < list.length; i++) {
      for (var j = i + 1; j < list.length; j++) {
        if (rangesOverlap(list[i].effectiveDate, list[i].endDate, list[j].effectiveDate, list[j].endDate)) {
          conflicts.push({ ruleIds: [list[i].id, list[j].id], personId: list[i].personId, weekday: list[i].weekday, periodCode: list[i].periodCode });
        }
      }
    }
  });
  return conflicts;
}

function hydrateWeeklyRule(r) {
  return {
    id: r.id, semesterId: r.semesterId, personId: r.personId, person: getPersonRef(r.personId),
    ruleType: r.ruleType, projectId: r.projectId || null, project: getProjectRef(r.projectId),
    weekday: r.weekday, periodCode: r.periodCode, subject: r.subject || null,
    weeklyPeriods: Number(r.weeklyPeriods), effectiveDate: r.effectiveDate, endDate: r.endDate || null,
    note: r.note, createdAt: r.createdAt, updatedAt: r.updatedAt,
  };
}

function api_listWeeklyRules(payload) {
  requireField(payload, "semesterId", "學期");
  var rows = readRows("WeeklyRules").filter(function (r) {
    if (r.semesterId !== payload.semesterId) return false;
    if (payload.personId && r.personId !== payload.personId) return false;
    if (payload.ruleType && r.ruleType !== payload.ruleType) return false;
    if (payload.weekday && r.weekday !== payload.weekday) return false;
    if (payload.projectId && r.projectId !== payload.projectId) return false;
    return true;
  });
  return rows.map(hydrateWeeklyRule);
}

function api_weeklyRuleConflicts(payload) {
  requireField(payload, "semesterId", "學期");
  return detectWeeklyRuleConflicts(payload.semesterId, payload.personId);
}

function validateRuleTypeProject(ruleType, projectId) {
  if (ruleType === "PROJECT" && !projectId) throw new Error("ruleType 為 PROJECT 時必須指定 projectId");
  if (ruleType === "OVERTIME" && projectId) throw new Error("ruleType 為 OVERTIME 時不應該指定 projectId");
}

function validateDateRange(effectiveDate, endDate) {
  if (endDate && new Date(endDate) < new Date(effectiveDate)) throw new Error("結束日期不能早於生效日期");
}

function api_createWeeklyRule(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "personId", "教師");
  requireField(payload, "ruleType", "規則類型");
  requireField(payload, "weekday", "星期");
  requireField(payload, "periodCode", "節次");
  requireField(payload, "effectiveDate", "生效日期");
  validateRuleTypeProject(payload.ruleType, payload.projectId);
  validateDateRange(payload.effectiveDate, payload.endDate);

  var rule = {
    id: newId(), semesterId: payload.semesterId, personId: payload.personId, ruleType: payload.ruleType,
    projectId: payload.projectId || "", weekday: payload.weekday, periodCode: payload.periodCode,
    subject: payload.subject || "", weeklyPeriods: Number(payload.weeklyPeriods || 0),
    effectiveDate: toDateOnly(payload.effectiveDate), endDate: payload.endDate ? toDateOnly(payload.endDate) : "",
    note: payload.note || "", createdAt: nowIso(), updatedAt: nowIso(),
  };
  appendRow("WeeklyRules", rule);
  writeChangeLog("special_weekly_rules", rule.id, null, null, describeWeeklyRule(rule), payload.changedBy, "新增每週固定規則");
  var conflicts = detectWeeklyRuleConflicts(payload.semesterId, payload.personId);
  return { rule: hydrateWeeklyRule(rule), conflicts: conflicts };
}

function describeWeeklyRule(r) {
  return r.ruleType + " " + r.weekday + " " + r.periodCode + " 每週" + r.weeklyPeriods + "節，自" + r.effectiveDate;
}

function api_updateWeeklyRule(payload) {
  requireField(payload, "id", "id");
  var existing = findById("WeeklyRules", payload.id);
  if (!existing) throw new Error("找不到規則");
  var patch = { updatedAt: nowIso() };
  ["ruleType", "projectId", "weekday", "periodCode", "subject", "note"].forEach(function (k) {
    if (payload[k] !== undefined) patch[k] = payload[k] === "" ? "" : payload[k];
  });
  if (payload.weeklyPeriods !== undefined) patch.weeklyPeriods = Number(payload.weeklyPeriods);
  if (payload.effectiveDate !== undefined) patch.effectiveDate = toDateOnly(payload.effectiveDate);
  if (payload.endDate !== undefined) patch.endDate = payload.endDate ? toDateOnly(payload.endDate) : "";

  var ruleType = patch.ruleType || existing.ruleType;
  var projectId = patch.projectId !== undefined ? patch.projectId : existing.projectId;
  validateRuleTypeProject(ruleType, projectId);
  validateDateRange(patch.effectiveDate || existing.effectiveDate, patch.endDate !== undefined ? patch.endDate : existing.endDate);

  var updated = updateRow("WeeklyRules", payload.id, patch);
  writeChangeLog("special_weekly_rules", payload.id, null, describeWeeklyRule(existing), describeWeeklyRule(updated), payload.changedBy, payload.reason || "修改每週固定規則");
  var conflicts = detectWeeklyRuleConflicts(updated.semesterId, updated.personId);
  return { rule: hydrateWeeklyRule(updated), conflicts: conflicts };
}

function api_deactivateWeeklyRule(payload) {
  requireField(payload, "id", "id");
  requireField(payload, "endDate", "停用日期");
  var existing = findById("WeeklyRules", payload.id);
  if (!existing) throw new Error("找不到規則");
  var updated = updateRow("WeeklyRules", payload.id, { endDate: toDateOnly(payload.endDate), updatedAt: nowIso() });
  writeChangeLog("special_weekly_rules", payload.id, "endDate", existing.endDate, updated.endDate, payload.changedBy, payload.reason || "停用規則");
  return hydrateWeeklyRule(updated);
}

// ---------- 單日例外 ----------

function hydrateDateRule(r) {
  return {
    id: r.id, semesterId: r.semesterId, date: r.date, personId: r.personId, person: getPersonRef(r.personId),
    periodCode: r.periodCode, originalClassificationNote: r.originalClassificationNote || null,
    overrideClassification: r.overrideClassification, projectId: r.projectId || null, project: getProjectRef(r.projectId),
    note: r.note, isCancelled: toBool(r.isCancelled), cancelledAt: r.cancelledAt || null, cancelledBy: r.cancelledBy || null,
    createdAt: r.createdAt, updatedAt: r.updatedAt,
  };
}

function api_listDateRules(payload) {
  requireField(payload, "semesterId", "學期");
  var includeCancelled = payload.includeCancelled === true || payload.includeCancelled === "true";
  var rows = readRows("DateRules").filter(function (r) {
    if (r.semesterId !== payload.semesterId) return false;
    if (payload.personId && r.personId !== payload.personId) return false;
    if (!includeCancelled && toBool(r.isCancelled)) return false;
    return true;
  });
  return rows.map(hydrateDateRule);
}

function api_createDateRule(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "date", "日期");
  requireField(payload, "personId", "教師");
  requireField(payload, "periodCode", "節次");
  requireField(payload, "overrideClassification", "改判分類");
  if (payload.overrideClassification === "PROJECT" && !payload.projectId) {
    throw new Error("改判為 PROJECT 時必須指定 projectId");
  }
  var rule = {
    id: newId(), semesterId: payload.semesterId, date: toDateOnly(payload.date), personId: payload.personId,
    periodCode: payload.periodCode, originalClassificationNote: payload.originalClassificationNote || "",
    overrideClassification: payload.overrideClassification, projectId: payload.projectId || "",
    note: payload.note || "", isCancelled: false, cancelledAt: "", cancelledBy: "",
    createdAt: nowIso(), updatedAt: nowIso(),
  };
  appendRow("DateRules", rule);
  writeChangeLog("special_date_rules", rule.id, null, null, rule.date + " " + rule.periodCode + " -> " + rule.overrideClassification, payload.changedBy, "新增單日例外");
  return hydrateDateRule(rule);
}

function api_cancelDateRule(payload) {
  requireField(payload, "id", "id");
  var existing = findById("DateRules", payload.id);
  if (!existing) throw new Error("找不到單日例外");
  var updated = updateRow("DateRules", payload.id, {
    isCancelled: true, cancelledAt: nowIso(), cancelledBy: payload.changedBy || "", updatedAt: nowIso(),
  });
  writeChangeLog("special_date_rules", payload.id, "isCancelled", "false", "true", payload.changedBy, payload.reason || "取消單日例外");
  return hydrateDateRule(updated);
}

// ---------- 費用規則 ----------

function api_getFeeRuleHistory(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "feeType", "費用類型");
  var rows = readRows("FeeRules").filter(function (r) { return r.semesterId === payload.semesterId && r.feeType === payload.feeType; });
  rows.sort(function (a, b) { return new Date(a.effectiveDate) - new Date(b.effectiveDate); });
  return rows.map(hydrateFeeRule);
}

function hydrateFeeRule(r) {
  return { id: r.id, semesterId: r.semesterId, feeType: r.feeType, amount: decToStr(r.amount), effectiveDate: r.effectiveDate, endDate: r.endDate || null, note: r.note, createdAt: r.createdAt, updatedAt: r.updatedAt };
}

function api_createFeeRule(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "feeType", "費用類型");
  requireField(payload, "amount", "金額");
  requireField(payload, "effectiveDate", "生效日期");
  validateDateRange(payload.effectiveDate, payload.endDate);
  var rule = {
    id: newId(), semesterId: payload.semesterId, feeType: payload.feeType, amount: Number(payload.amount),
    effectiveDate: toDateOnly(payload.effectiveDate), endDate: payload.endDate ? toDateOnly(payload.endDate) : "",
    note: payload.note || "", createdAt: nowIso(), updatedAt: nowIso(),
  };
  appendRow("FeeRules", rule);
  writeChangeLog("fee_rules", rule.id, null, null, rule.feeType + "=" + rule.amount + "(自" + rule.effectiveDate + ")", payload.changedBy, "新增費用規則版本");
  return hydrateFeeRule(rule);
}

function api_deactivateFeeRule(payload) {
  requireField(payload, "id", "id");
  requireField(payload, "endDate", "停用日期");
  var existing = findById("FeeRules", payload.id);
  if (!existing) throw new Error("找不到費用規則");
  var updated = updateRow("FeeRules", payload.id, { endDate: toDateOnly(payload.endDate), updatedAt: nowIso() });
  writeChangeLog("fee_rules", payload.id, "endDate", existing.endDate, updated.endDate, payload.changedBy, payload.reason || "停用費用規則");
  return hydrateFeeRule(updated);
}

// 依日期查出當時生效的費率：effectiveDate <= onDate <= (endDate 或無上限)，
// 有多筆符合時取 effectiveDate 最新的一筆（跟 Node 版本 orderBy desc 一致）。
function getEffectiveFeeRule(semesterId, feeType, onDateOnly) {
  var candidates = readRows("FeeRules").filter(function (r) {
    if (r.semesterId !== semesterId || r.feeType !== feeType) return false;
    if (String(r.effectiveDate) > onDateOnly) return false;
    if (r.endDate && String(r.endDate) < onDateOnly) return false;
    return true;
  });
  if (candidates.length === 0) return null;
  candidates.sort(function (a, b) { return String(b.effectiveDate).localeCompare(String(a.effectiveDate)); });
  return candidates[0];
}

// ---------- 學校上課日曆 ----------

function api_generateSemesterCalendar(payload) {
  requireField(payload, "semesterId", "學期");
  var semester = findById("Semesters", payload.semesterId);
  if (!semester) throw new Error("找不到學期");
  // semester.startDate/endDate 有可能是使用者直接在 Sheets 手動輸入的資料（例如既有的
  // 115 學年度第1學期），讀出來未必是 "YYYY-MM-DD" 字串——先正規化，否則下面的日期
  // 運算會整批變成 Invalid Date，迴圈一次都不會執行，永遠回報「已產生 0 天」
  // （這正是這次要修的 bug 本身，見 toDateOnly() 註解）。
  var startDate = toDateOnly(semester.startDate);
  var endDate = toDateOnly(semester.endDate);
  if (!startDate || !endDate) throw new Error("這個學期沒有設定開始或結束日期，無法產生日曆");

  var existing = readRows("SchoolCalendarDays").filter(function (r) { return r.semesterId === payload.semesterId; });
  var existingDates = {};
  existing.forEach(function (d) { existingDates[d.date] = true; });

  var toCreate = [];
  var cursor = new Date(startDate + "T00:00:00Z");
  var end = new Date(endDate + "T00:00:00Z");
  while (cursor.getTime() <= end.getTime()) {
    var dateOnly = cursor.toISOString().slice(0, 10);
    // 已經存在的日期（不論是先前產生過，還是管理者手動修改過）一律跳過，不覆蓋、
    // 不重複新增——人工設定過的休假日／上課日狀態一定要保留原樣。
    if (!existingDates[dateOnly]) {
      var weekday = weekdayOfDateOnly(dateOnly);
      toCreate.push({
        id: newId(), semesterId: payload.semesterId, date: dateOnly, weekday: weekday,
        isTeachingDay: weekday !== "SAT" && weekday !== "SUN", note: "",
        createdAt: nowIso(), updatedAt: nowIso(),
      });
    }
    cursor.setUTCDate(cursor.getUTCDate() + 1);
  }
  if (toCreate.length > 0) appendRows("SchoolCalendarDays", toCreate);
  writeChangeLog("school_calendar_days", payload.semesterId, null, null,
    "新增 " + toCreate.length + " 天（" + startDate + " ~ " + endDate + "，原本已有 " + existing.length + " 天）",
    payload.changedBy, "批次產生學期日曆");
  return { createdCount: toCreate.length, skippedCount: existing.length, startDate: startDate, endDate: endDate };
}

function api_listCalendarDays(payload) {
  requireField(payload, "semesterId", "學期");
  var rows = readRows("SchoolCalendarDays").filter(function (r) {
    if (r.semesterId !== payload.semesterId) return false;
    if (payload.year && payload.month) {
      var monthStr = String(payload.month).length === 1 ? "0" + payload.month : String(payload.month);
      var prefix = payload.year + "-" + monthStr;
      if (String(r.date).indexOf(prefix) !== 0) return false;
    }
    return true;
  });
  rows.sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
  return rows.map(function (r) {
    return { id: r.id, semesterId: r.semesterId, date: r.date, weekday: r.weekday, isTeachingDay: toBool(r.isTeachingDay), note: r.note, createdAt: r.createdAt, updatedAt: r.updatedAt };
  });
}

function summarizeCalendarDays(days) {
  var weekdays = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"];
  var counts = {};
  var total = 0;
  days.forEach(function (d) {
    if (toBool(d.isTeachingDay)) {
      counts[d.weekday] = (counts[d.weekday] || 0) + 1;
      total += 1;
    }
  });
  return { byWeekday: weekdays.map(function (w) { return { weekday: w, teachingDayCount: counts[w] || 0 }; }), totalTeachingDays: total };
}

function api_calendarSummary(payload) {
  requireField(payload, "semesterId", "學期");
  var days = api_listCalendarDays(payload);
  return summarizeCalendarDays(days);
}

function api_addCalendarDay(payload) {
  requireField(payload, "semesterId", "學期");
  requireField(payload, "date", "日期");
  var dateOnly = toDateOnly(payload.date);
  var dup = findOne("SchoolCalendarDays", function (r) { return r.semesterId === payload.semesterId && r.date === dateOnly; });
  if (dup) throw new Error("這一天已經存在於學校上課日曆");
  var day = {
    id: newId(), semesterId: payload.semesterId, date: dateOnly, weekday: weekdayOfDateOnly(dateOnly),
    isTeachingDay: payload.isTeachingDay !== undefined ? Boolean(payload.isTeachingDay) : true,
    note: payload.note || "", createdAt: nowIso(), updatedAt: nowIso(),
  };
  appendRow("SchoolCalendarDays", day);
  writeChangeLog("school_calendar_days", day.id, null, null, dateOnly, payload.changedBy, "新增單一上課日");
  return day;
}

function api_updateCalendarDay(payload) {
  requireField(payload, "id", "id");
  var existing = findById("SchoolCalendarDays", payload.id);
  if (!existing) throw new Error("找不到日曆資料");
  var patch = { updatedAt: nowIso() };
  if (payload.isTeachingDay !== undefined) patch.isTeachingDay = Boolean(payload.isTeachingDay);
  if (payload.note !== undefined) patch.note = payload.note === "" ? "" : payload.note;
  var updated = updateRow("SchoolCalendarDays", payload.id, patch);
  writeChangeLog("school_calendar_days", payload.id, "isTeachingDay", String(toBool(existing.isTeachingDay)), String(toBool(updated.isTeachingDay)), payload.changedBy, payload.reason || "修改上課日狀態");
  return updated;
}

// ---------- 共用工具 ----------

function toBool(v) {
  return v === true || v === "TRUE" || v === "true" || v === 1;
}

function stripRow(r) {
  var copy = {};
  Object.keys(r).forEach(function (k) { if (k !== "__row") copy[k] = r[k]; });
  return copy;
}
