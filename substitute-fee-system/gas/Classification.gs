/**
 * Classification.gs — 自動分類引擎，逐行對照 classificationService.ts 的
 * computeClassification()。優先順序完全不變：
 *   1. SpecialDateRule（單日例外）完全比對 學期+原教師+日期+節次，優先於週規則
 *   2. SpecialWeeklyRule（每週固定規則）比對 學期+原教師+星期+節次，且落在生效期間內
 *   3. 都沒有 → 一般公費（GENERAL）
 * 任一層有多個候選規則都標記 CONFLICT，不自動選一個；原教師未配對標記 TEACHER_UNMATCHED。
 */

var FUNDING_SOURCE_LABEL = { GENERAL: "一般公費", OVERTIME: "超鐘點", PROJECT: "專案", UNDETERMINED: "待確認" };
var WEEKDAY_LABEL = { MON: "星期一", TUE: "星期二", WED: "星期三", THU: "星期四", FRI: "星期五", SAT: "星期六", SUN: "星期日" };

function computeClassification(record, semester) {
  if (!record.originalTeacherId) {
    return { fundingSource: "UNDETERMINED", classificationMethod: "TEACHER_UNMATCHED", ruleId: null, projectId: null, conflict: null };
  }
  if (!record.periodCode) {
    return { fundingSource: "UNDETERMINED", classificationMethod: "TEACHER_UNMATCHED", ruleId: null, projectId: null, conflict: null };
  }

  // 第一優先：單日例外
  var dateRules = readRows("DateRules").filter(function (r) {
    return r.semesterId === semester.id && r.personId === record.originalTeacherId &&
      r.date === record.date && r.periodCode === record.periodCode && !toBool(r.isCancelled);
  });
  if (dateRules.length === 1) {
    var dr = dateRules[0];
    return { fundingSource: dr.overrideClassification, classificationMethod: "DATE_EXCEPTION", ruleId: dr.id, projectId: dr.overrideClassification === "PROJECT" ? dr.projectId : null, conflict: null };
  }
  if (dateRules.length > 1) {
    return {
      fundingSource: "UNDETERMINED", classificationMethod: "CONFLICT", ruleId: null, projectId: null,
      conflict: { source: "DATE_EXCEPTION", candidates: dateRules.map(function (r) { return { ruleId: r.id, wouldBeFundingSource: r.overrideClassification, projectId: r.overrideClassification === "PROJECT" ? r.projectId : null }; }) },
    };
  }

  // 第二優先：每週固定規則（有效期間必須涵蓋代課日期）
  var weeklyRulesRaw = readRows("WeeklyRules").filter(function (r) {
    if (r.semesterId !== semester.id || r.personId !== record.originalTeacherId) return false;
    if (r.weekday !== record.weekday || r.periodCode !== record.periodCode) return false;
    if (String(r.effectiveDate) > record.date) return false;
    if (r.endDate && String(r.endDate) < record.date) return false;
    return true;
  });
  var subjectModeRequired = semester.overtimeMatchMode === "TEACHER_WEEKDAY_PERIOD_SUBJECT";
  var validCandidates = weeklyRulesRaw.filter(function (rule) {
    if (rule.ruleType === "PROJECT" && !rule.projectId) return false;
    if (subjectModeRequired) return rule.subject && record.subject && rule.subject === record.subject;
    return true;
  });

  if (validCandidates.length === 1) {
    var wr = validCandidates[0];
    return { fundingSource: wr.ruleType, classificationMethod: "WEEKLY_RULE", ruleId: wr.id, projectId: wr.ruleType === "PROJECT" ? wr.projectId : null, conflict: null };
  }
  if (validCandidates.length > 1) {
    return {
      fundingSource: "UNDETERMINED", classificationMethod: "CONFLICT", ruleId: null, projectId: null,
      conflict: { source: "WEEKLY_RULE", candidates: validCandidates.map(function (r) { return { ruleId: r.id, wouldBeFundingSource: r.ruleType, projectId: r.ruleType === "PROJECT" ? r.projectId : null }; }) },
    };
  }

  // 第三優先：一般公費
  return { fundingSource: "GENERAL", classificationMethod: "GENERAL_DEFAULT", ruleId: null, projectId: null, conflict: null };
}

function classifySubstituteRecord(recordId, changedBy) {
  var record = findById("SubstituteRecords", recordId);
  if (!record) throw new Error("找不到代課紀錄");
  var monthlyImport = findById("MonthlyImports", record.monthlyImportId);
  assertMonthNotLocked(Number(monthlyImport.year), Number(monthlyImport.month));
  var semester = findById("Semesters", monthlyImport.semesterId);

  var outcome = computeClassification(record, semester);
  var patch = {
    classifiedAt: nowIso(), autoFundingSource: outcome.fundingSource, autoClassificationMethod: outcome.classificationMethod,
    autoClassificationRuleId: outcome.ruleId || "", autoProjectId: outcome.projectId || "",
    conflictCandidatesJson: outcome.conflict ? JSON.stringify(outcome.conflict) : "", updatedAt: nowIso(),
  };
  var wasManual = toBool(record.isManuallyModified);
  if (!wasManual) {
    patch.fundingSource = outcome.fundingSource;
    patch.classificationMethod = outcome.classificationMethod;
    patch.classificationRuleId = outcome.ruleId || "";
    patch.projectId = outcome.projectId || "";
  }
  var updated = updateRow("SubstituteRecords", recordId, patch);

  if (!wasManual && record.fundingSource !== outcome.fundingSource) {
    writeChangeLog("substitute_records", recordId, "fundingSource", record.fundingSource, outcome.fundingSource, changedBy, "自動分類（" + outcome.classificationMethod + "）");
  }
  return hydrateSubstituteRecord(updated);
}

function api_classifyMonthlyImport(payload) {
  requireField(payload, "id", "id");
  var records = readRows("SubstituteRecords").filter(function (r) { return r.monthlyImportId === payload.id; });
  var summary = { total: records.length, general: 0, overtime: 0, project: 0, conflict: 0, teacherUnmatched: 0, manualPreserved: 0 };
  records.forEach(function (r) {
    var updated = classifySubstituteRecord(r.id, payload.changedBy);
    if (updated.isManuallyModified) { summary.manualPreserved += 1; return; }
    if (updated.fundingSource === "GENERAL") summary.general += 1;
    else if (updated.fundingSource === "OVERTIME") summary.overtime += 1;
    else if (updated.fundingSource === "PROJECT") summary.project += 1;
    else if (updated.classificationMethod === "CONFLICT") summary.conflict += 1;
    else if (updated.classificationMethod === "TEACHER_UNMATCHED") summary.teacherUnmatched += 1;
  });
  return summary;
}

function api_overrideClassification(payload) {
  requireField(payload, "id", "id");
  requireField(payload, "fundingSource", "分類結果");
  requireField(payload, "reason", "覆寫原因");
  if (payload.fundingSource === "PROJECT" && !payload.projectId) throw new Error("覆寫為 PROJECT 時必須指定 projectId");
  var existing = findById("SubstituteRecords", payload.id);
  if (!existing) throw new Error("找不到代課紀錄");
  assertRecordMonthNotLocked(payload.id);

  var updated = updateRow("SubstituteRecords", payload.id, {
    fundingSource: payload.fundingSource, projectId: payload.fundingSource === "PROJECT" ? payload.projectId : "",
    classificationMethod: "MANUAL_OVERRIDE", classificationRuleId: "", isManuallyModified: true,
    manualOverrideReason: payload.reason, manualOverrideAt: nowIso(), manualOverrideBy: payload.changedBy || "",
    updatedAt: nowIso(),
  });
  writeChangeLog("substitute_records", payload.id, "fundingSource", existing.fundingSource, payload.fundingSource, payload.changedBy, payload.reason);
  return hydrateSubstituteRecord(updated);
}

function api_revertToAutoClassification(payload) {
  requireField(payload, "id", "id");
  requireField(payload, "reason", "復原原因");
  var record = findById("SubstituteRecords", payload.id);
  if (!record) throw new Error("找不到代課紀錄");
  assertRecordMonthNotLocked(payload.id);
  if (!toBool(record.isManuallyModified)) throw new Error("此筆尚未被人工覆寫，不需要復原");

  var updated = updateRow("SubstituteRecords", payload.id, {
    fundingSource: record.autoFundingSource || "UNDETERMINED", classificationMethod: record.autoClassificationMethod || "GENERAL_DEFAULT",
    classificationRuleId: record.autoClassificationRuleId || "", projectId: record.autoProjectId || "",
    isManuallyModified: false, manualOverrideReason: "", manualOverrideAt: "", manualOverrideBy: "", updatedAt: nowIso(),
  });
  writeChangeLog("substitute_records", payload.id, "isManuallyModified", "true", "false", payload.changedBy, payload.reason);
  return hydrateSubstituteRecord(updated);
}

function api_listClassificationPreview(payload) {
  requireField(payload, "id", "id");
  var records = readRows("SubstituteRecords").filter(function (r) {
    if (r.monthlyImportId !== payload.id) return false;
    if (payload.fundingSource && r.fundingSource !== payload.fundingSource) return false;
    if (payload.classificationMethod && r.classificationMethod !== payload.classificationMethod) return false;
    if (payload.isManuallyModified !== undefined && payload.isManuallyModified !== "" && toBool(r.isManuallyModified) !== (payload.isManuallyModified === true || payload.isManuallyModified === "true")) return false;
    if (payload.staffType && r.staffType !== payload.staffType) return false;
    return true;
  });
  records.sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
  return records.map(function (r) {
    var hydrated = hydrateSubstituteRecord(r);
    hydrated.classificationBasisText = describeClassificationBasis(r);
    return hydrated;
  });
}

// 把分類結果講成人話，逐行對照 classificationService.ts 的 describeClassificationBasis()。
function describeClassificationBasis(record) {
  var fundingLabel = FUNDING_SOURCE_LABEL[record.fundingSource] || record.fundingSource;
  switch (record.classificationMethod) {
    case "WEEKLY_RULE": {
      if (!record.classificationRuleId) return "每週規則 → " + fundingLabel;
      var rule = findById("WeeklyRules", record.classificationRuleId);
      if (!rule) return "每週規則（規則已不存在）→ " + fundingLabel;
      var project = getProjectRef(rule.projectId);
      return "每週規則：" + (WEEKDAY_LABEL[rule.weekday] || rule.weekday) + " " + rule.periodCode + " → " + fundingLabel + (project ? "：" + project.name : "");
    }
    case "DATE_EXCEPTION": {
      if (!record.classificationRuleId) return "單日例外 → " + fundingLabel;
      var dr = findById("DateRules", record.classificationRuleId);
      if (!dr) return "單日例外（規則已不存在）→ " + fundingLabel;
      var proj = getProjectRef(dr.projectId);
      return "單日例外：" + dr.date + " " + dr.periodCode + " → " + fundingLabel + (proj ? "：" + proj.name : "");
    }
    case "GENERAL_DEFAULT":
      return "一般公費：沒有符合特殊規則";
    case "MANUAL_OVERRIDE":
      return "人工覆寫 → " + fundingLabel;
    case "CONFLICT": {
      var count = 0;
      if (record.conflictCandidatesJson) {
        try {
          var parsed = JSON.parse(record.conflictCandidatesJson);
          count = Array.isArray(parsed.candidates) ? parsed.candidates.length : 0;
        } catch (e) { /* ignore */ }
      }
      return "規則衝突：同時符合 " + (count || "多") + " 個規則，需要人工確認";
    }
    case "TEACHER_UNMATCHED":
      return "原教師尚未配對，無法判斷規則";
    default:
      return fundingLabel;
  }
}
