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

// ---------- 學期／人員／節次代碼 ----------

function api_listSemesters() {
  var rows = readRows("Semesters").map(stripRow);
  rows.sort(function (a, b) {
    if (a.schoolYear !== b.schoolYear) return Number(b.schoolYear) - Number(a.schoolYear);
    return Number(b.term) - Number(a.term);
  });
  return rows.map(function (r) {
    return {
      id: r.id, schoolYear: Number(r.schoolYear), term: Number(r.term), status: r.status,
      isCurrent: r.isCurrent === true || r.isCurrent === "TRUE" || r.isCurrent === "true",
      startDate: r.startDate, endDate: r.endDate, overtimeMatchMode: r.overtimeMatchMode,
      note: r.note, createdAt: r.createdAt, updatedAt: r.updatedAt,
    };
  });
}

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

function api_listProjects(payload) {
  requireField(payload, "semesterId", "學期");
  var rows = readRows("Projects").filter(function (r) { return r.semesterId === payload.semesterId; }).map(stripRow);
  return rows.map(function (r) {
    return { id: r.id, semesterId: r.semesterId, name: r.name, isActive: toBool(r.isActive), note: r.note, createdAt: r.createdAt, updatedAt: r.updatedAt };
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
  var existing = readRows("SchoolCalendarDays").filter(function (r) { return r.semesterId === payload.semesterId; });
  var existingDates = {};
  existing.forEach(function (d) { existingDates[d.date] = true; });

  var toCreate = [];
  var cursor = new Date(semester.startDate + "T00:00:00Z");
  var end = new Date(semester.endDate + "T00:00:00Z");
  while (cursor.getTime() <= end.getTime()) {
    var dateOnly = cursor.toISOString().slice(0, 10);
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
  writeChangeLog("school_calendar_days", payload.semesterId, null, null, "新增 " + toCreate.length + " 天", payload.changedBy, "批次產生學期日曆（略過已存在 " + existing.length + " 天）");
  return { createdCount: toCreate.length, skippedCount: existing.length };
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
