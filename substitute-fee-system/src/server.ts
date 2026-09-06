// Phase 5/6/7：超鐘點／專案設定、學校上課日曆、公費代課 Excel 匯入的基本管理介面
// （後端 API + 靜態頁面）
//
// 這是目前系統唯一的「使用者介面」，刻意保持簡單：一個 Express server
// 把 Phase 1-7 已經寫好的 service 包成 JSON API，前端是不需要建置流程的
// 純 HTML/JS（public/ 目錄）。行政人員好操作優先，不追求美觀。

import express from "express";
import multer from "multer";
import path from "path";
import { prisma } from "./prismaClient";
import { createProject, listProjects, setProjectActive, updateProject } from "./services/projectService";
import { listPersons } from "./services/personService";
import {
  createSpecialDateRule,
  cancelSpecialDateRule,
  listDateRules,
} from "./services/specialDateRuleService";
import {
  createSpecialWeeklyRule,
  deactivateSpecialWeeklyRule,
  detectWeeklyRuleConflicts,
  listWeeklyRules,
  updateSpecialWeeklyRule,
} from "./services/specialWeeklyRuleService";
import {
  addCalendarDay,
  generateSemesterCalendar,
  getMonthlySummary,
  getSemesterSummary,
  listCalendarDays,
  updateCalendarDay,
} from "./services/schoolCalendarService";
import {
  autoApplyUnambiguousTeacherMatches,
  getMonthlyImportDetail,
  importSubstituteExcel,
  listMonthlyImports,
  listSubstituteRecords,
  listUnmatchedTeacherReferences,
  listWorkbookSheets,
  resolveTeacherReference,
} from "./services/excelImportService";
import {
  classifyMonthlyImport,
  listClassificationPreview,
  overrideClassification,
  revertToAutoClassification,
} from "./services/classificationService";
import { createFeeRule, deactivateFeeRule, getFeeRuleHistory } from "./services/feeRuleService";
import { calculateMonthlyImportFees, summarizeTeacherMonthlyFees } from "./services/feeCalculationService";
import { getMonthlyLockStatus, lockMonth, unlockMonth, getBlockingIssues } from "./services/monthlyLockService";
import { acknowledgeIssue, revokeAcknowledgement, listAcknowledgements } from "./services/issueAcknowledgementService";
import { getMonthlyDashboard } from "./services/dashboardService";
import { listPendingIssues } from "./services/pendingIssuesService";
import { createSelfFundedRecord, updateSelfFundedRecord, deleteSelfFundedRecord, listSelfFundedRecords } from "./services/selfFundedService";
import { getChunaSummary, generateChunaExcelBuffer } from "./services/chunaService";
import { parseUploadedChunaWorkbook, reconcile } from "./services/reconciliationService";

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
app.use(express.json());
app.use(express.static(path.join(__dirname, "..", "public")));

function asDate(value: unknown): Date | undefined {
  if (value === undefined || value === null || value === "") return undefined;
  return new Date(String(value));
}

function handleError(res: express.Response, err: unknown) {
  const message = err instanceof Error ? err.message : "未知錯誤";
  res.status(400).json({ error: message });
}

// ---- 基礎資料：學期／人員／節次代碼 ----

app.get("/api/semesters", async (_req, res) => {
  const semesters = await prisma.semester.findMany({
    orderBy: [{ schoolYear: "desc" }, { term: "desc" }],
  });
  res.json(semesters);
});

app.get("/api/persons", async (req, res) => {
  const search = typeof req.query.search === "string" ? req.query.search : undefined;
  const persons = await listPersons("ALL", search);
  res.json(persons);
});

app.get("/api/period-slots", async (_req, res) => {
  const slots = await prisma.periodSlot.findMany({ orderBy: { sortOrder: "asc" } });
  res.json(slots);
});

// ---- 專案 ----

app.get("/api/projects", async (req, res) => {
  const semesterId = String(req.query.semesterId ?? "");
  if (!semesterId) return res.status(400).json({ error: "缺少 semesterId" });
  const projects = await listProjects(semesterId);
  res.json(projects);
});

app.post("/api/projects", async (req, res) => {
  try {
    const { semesterId, name, note, changedBy } = req.body;
    const project = await createProject({ semesterId, name, note }, changedBy);
    res.json(project);
  } catch (err) {
    handleError(res, err);
  }
});

app.patch("/api/projects/:id", async (req, res) => {
  try {
    const { name, note, changedBy, reason } = req.body;
    const project = await updateProject(req.params.id, { name, note }, changedBy, reason);
    res.json(project);
  } catch (err) {
    handleError(res, err);
  }
});

app.patch("/api/projects/:id/active", async (req, res) => {
  try {
    const { isActive, changedBy, reason } = req.body;
    const project = await setProjectActive(req.params.id, Boolean(isActive), changedBy, reason);
    res.json(project);
  } catch (err) {
    handleError(res, err);
  }
});

// ---- 每週固定規則（超鐘點／專案）----

app.get("/api/weekly-rules", async (req, res) => {
  const semesterId = String(req.query.semesterId ?? "");
  if (!semesterId) return res.status(400).json({ error: "缺少 semesterId" });
  const rules = await listWeeklyRules(semesterId, {
    personId: req.query.personId ? String(req.query.personId) : undefined,
    ruleType: req.query.ruleType ? (String(req.query.ruleType) as any) : undefined,
    weekday: req.query.weekday ? (String(req.query.weekday) as any) : undefined,
    projectId: req.query.projectId ? String(req.query.projectId) : undefined,
  });
  res.json(rules);
});

app.get("/api/weekly-rules/conflicts", async (req, res) => {
  const semesterId = String(req.query.semesterId ?? "");
  if (!semesterId) return res.status(400).json({ error: "缺少 semesterId" });
  const conflicts = await detectWeeklyRuleConflicts(
    semesterId,
    req.query.personId ? String(req.query.personId) : undefined
  );
  res.json(conflicts);
});

app.post("/api/weekly-rules", async (req, res) => {
  try {
    const body = req.body;
    const result = await createSpecialWeeklyRule(
      {
        semesterId: body.semesterId,
        personId: body.personId,
        ruleType: body.ruleType,
        projectId: body.projectId || undefined,
        weekday: body.weekday,
        periodCode: body.periodCode,
        subject: body.subject || undefined,
        weeklyPeriods: body.weeklyPeriods,
        effectiveDate: asDate(body.effectiveDate)!,
        endDate: asDate(body.endDate),
        note: body.note || undefined,
      },
      body.changedBy
    );
    res.json(result);
  } catch (err) {
    handleError(res, err);
  }
});

app.patch("/api/weekly-rules/:id", async (req, res) => {
  try {
    const body = req.body;
    const changes: Record<string, unknown> = {};
    for (const key of ["ruleType", "projectId", "weekday", "periodCode", "subject", "weeklyPeriods", "note"]) {
      if (key in body) changes[key] = body[key] === "" ? null : body[key];
    }
    if ("effectiveDate" in body) changes.effectiveDate = asDate(body.effectiveDate);
    if ("endDate" in body) changes.endDate = asDate(body.endDate) ?? null;

    const result = await updateSpecialWeeklyRule(req.params.id, changes, body.changedBy, body.reason);
    res.json(result);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/weekly-rules/:id/deactivate", async (req, res) => {
  try {
    const { endDate, changedBy, reason } = req.body;
    const rule = await deactivateSpecialWeeklyRule(req.params.id, asDate(endDate)!, changedBy, reason);
    res.json(rule);
  } catch (err) {
    handleError(res, err);
  }
});

// ---- 單日例外 ----

app.get("/api/date-rules", async (req, res) => {
  const semesterId = String(req.query.semesterId ?? "");
  if (!semesterId) return res.status(400).json({ error: "缺少 semesterId" });
  const rules = await listDateRules(semesterId, {
    personId: req.query.personId ? String(req.query.personId) : undefined,
    includeCancelled: req.query.includeCancelled === "true",
  });
  res.json(rules);
});

app.post("/api/date-rules", async (req, res) => {
  try {
    const body = req.body;
    const rule = await createSpecialDateRule(
      {
        semesterId: body.semesterId,
        date: asDate(body.date)!,
        personId: body.personId,
        periodCode: body.periodCode,
        overrideClassification: body.overrideClassification,
        projectId: body.projectId || undefined,
        originalClassificationNote: body.originalClassificationNote || undefined,
        note: body.note || undefined,
      },
      body.changedBy
    );
    res.json(rule);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/date-rules/:id/cancel", async (req, res) => {
  try {
    const { changedBy, reason } = req.body;
    const rule = await cancelSpecialDateRule(req.params.id, changedBy, reason);
    res.json(rule);
  } catch (err) {
    handleError(res, err);
  }
});

// ---- 學校上課日曆 ----

app.get("/api/calendar", async (req, res) => {
  const semesterId = String(req.query.semesterId ?? "");
  if (!semesterId) return res.status(400).json({ error: "缺少 semesterId" });
  const year = req.query.year ? Number(req.query.year) : undefined;
  const month = req.query.month ? Number(req.query.month) : undefined;
  const days = await listCalendarDays(semesterId, { year, month });
  res.json(days);
});

app.get("/api/calendar/summary", async (req, res) => {
  const semesterId = String(req.query.semesterId ?? "");
  if (!semesterId) return res.status(400).json({ error: "缺少 semesterId" });
  const year = req.query.year ? Number(req.query.year) : undefined;
  const month = req.query.month ? Number(req.query.month) : undefined;
  const summary = year && month ? await getMonthlySummary(semesterId, year, month) : await getSemesterSummary(semesterId);
  res.json(summary);
});

app.post("/api/calendar/generate", async (req, res) => {
  try {
    const { semesterId, changedBy } = req.body;
    const result = await generateSemesterCalendar(semesterId, changedBy);
    res.json(result);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/calendar", async (req, res) => {
  try {
    const { semesterId, date, isTeachingDay, note, changedBy } = req.body;
    const day = await addCalendarDay({ semesterId, date: asDate(date)!, isTeachingDay, note }, changedBy);
    res.json(day);
  } catch (err) {
    handleError(res, err);
  }
});

app.patch("/api/calendar/:id", async (req, res) => {
  try {
    const { isTeachingDay, note, changedBy, reason } = req.body;
    const changes: Record<string, unknown> = {};
    if ("isTeachingDay" in req.body) changes.isTeachingDay = Boolean(isTeachingDay);
    if ("note" in req.body) changes.note = note === "" ? null : note;
    const day = await updateCalendarDay(req.params.id, changes, changedBy, reason);
    res.json(day);
  } catch (err) {
    handleError(res, err);
  }
});

// ---- 公費代課 Excel 匯入 ----

// 真實檔案通常有多個 Sheet（教師代碼對照、編制內/外明細、給出納彙總……），
// 上傳後先讓管理者看到有哪些 Sheet，選擇要匯入哪一個，而不是直接假設第一個。
app.post("/api/monthly-imports/inspect", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) throw new Error("請選擇要上傳的 Excel 檔案");
    const sheets = await listWorkbookSheets(req.file.buffer);
    // 依 Sheet 名稱粗略猜測 BD／非BD，僅為預先帶入的建議，管理者仍需自行確認/覆蓋
    const withSuggestion = sheets.map((s) => ({
      ...s,
      suggestedStaffType: s.name.includes("非BD") ? "NON_BD" : s.name.includes("BD") ? "BD" : "UNKNOWN",
    }));
    res.json(withSuggestion);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/monthly-imports", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) throw new Error("請選擇要上傳的 Excel 檔案");
    const { semesterId, year, month, changedBy, sheetName, sourceStaffType } = req.body;
    const result = await importSubstituteExcel({
      semesterId,
      year: Number(year),
      month: Number(month),
      fileName: req.file.originalname,
      fileBuffer: req.file.buffer,
      sheetName: sheetName || undefined,
      sourceStaffType: sourceStaffType || undefined,
      importedBy: changedBy || undefined,
    });
    res.json(result);
  } catch (err) {
    handleError(res, err);
  }
});

app.get("/api/monthly-imports", async (req, res) => {
  const semesterId = String(req.query.semesterId ?? "");
  if (!semesterId) return res.status(400).json({ error: "缺少 semesterId" });
  const year = req.query.year ? Number(req.query.year) : undefined;
  const month = req.query.month ? Number(req.query.month) : undefined;
  const batches = await listMonthlyImports(semesterId, { year, month });
  res.json(batches);
});

// 注意：必須放在 "/api/monthly-imports/:id" 之前註冊，否則 Express 會把
// "fee-summary" 當成 :id 吃掉，永遠進不到這個路由。
app.get("/api/monthly-imports/fee-summary", async (req, res) => {
  try {
    const idsParam = req.query.ids;
    const ids = typeof idsParam === "string" ? idsParam.split(",").filter(Boolean) : [];
    if (ids.length === 0) {
      res.status(400).json({ error: "請提供 ids（以逗號分隔的 monthlyImportId 清單）" });
      return;
    }
    const summary = await summarizeTeacherMonthlyFees(ids);
    res.json(summary);
  } catch (err) {
    handleError(res, err);
  }
});

app.get("/api/monthly-imports/:id", async (req, res) => {
  try {
    const detail = await getMonthlyImportDetail(req.params.id);
    res.json(detail);
  } catch (err) {
    handleError(res, err);
  }
});

app.get("/api/monthly-imports/:id/records", async (req, res) => {
  const records = await listSubstituteRecords(req.params.id);
  res.json(records);
});

app.get("/api/monthly-imports/:id/unmatched", async (req, res) => {
  const unmatched = await listUnmatchedTeacherReferences(req.params.id);
  res.json(unmatched);
});

app.post("/api/monthly-imports/:id/auto-apply-matches", async (req, res) => {
  try {
    const result = await autoApplyUnambiguousTeacherMatches(req.params.id, req.body?.changedBy || undefined);
    res.json(result);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/substitute-records/:id/resolve-teacher", async (req, res) => {
  try {
    const { field, personId, changedBy } = req.body;
    const record = await resolveTeacherReference(req.params.id, field, personId, changedBy || undefined);
    res.json(record);
  } catch (err) {
    handleError(res, err);
  }
});

// ---- Phase 8：自動分類（只判斷來源類型，不算金額）----

app.post("/api/monthly-imports/:id/classify", async (req, res) => {
  try {
    const summary = await classifyMonthlyImport(req.params.id, req.body?.changedBy || undefined);
    res.json(summary);
  } catch (err) {
    handleError(res, err);
  }
});

app.get("/api/monthly-imports/:id/classification-preview", async (req, res) => {
  const { fundingSource, classificationMethod, isManuallyModified, staffType } = req.query;
  const records = await listClassificationPreview(req.params.id, {
    fundingSource: fundingSource ? (String(fundingSource) as any) : undefined,
    classificationMethod: classificationMethod ? (String(classificationMethod) as any) : undefined,
    isManuallyModified: isManuallyModified !== undefined ? isManuallyModified === "true" : undefined,
    staffType: staffType ? (String(staffType) as any) : undefined,
  });
  res.json(records);
});

app.post("/api/substitute-records/:id/override-classification", async (req, res) => {
  try {
    const { fundingSource, projectId, changedBy, reason } = req.body;
    const record = await overrideClassification(req.params.id, { fundingSource, projectId }, changedBy || undefined, reason);
    res.json(record);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/substitute-records/:id/revert-classification", async (req, res) => {
  try {
    const { changedBy, reason } = req.body;
    const record = await revertToAutoClassification(req.params.id, changedBy || undefined, reason);
    res.json(record);
  } catch (err) {
    handleError(res, err);
  }
});

// ---- Phase 9 第一階段：費用規則管理 + 節次型費用計算 ----

app.get("/api/fee-rules", async (req, res) => {
  try {
    const { semesterId, feeType } = req.query;
    if (!semesterId || !feeType) {
      res.status(400).json({ error: "semesterId 與 feeType 為必填" });
      return;
    }
    const history = await getFeeRuleHistory(String(semesterId), String(feeType) as any);
    res.json(history);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/fee-rules", async (req, res) => {
  try {
    const { semesterId, feeType, amount, effectiveDate, endDate, note, changedBy } = req.body;
    const rule = await createFeeRule(
      { semesterId, feeType, amount, effectiveDate: asDate(effectiveDate)!, endDate: asDate(endDate), note },
      changedBy || undefined
    );
    res.json(rule);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/fee-rules/:id/deactivate", async (req, res) => {
  try {
    const { endDate, changedBy, reason } = req.body;
    const rule = await deactivateFeeRule(req.params.id, asDate(endDate)!, changedBy || undefined, reason);
    res.json(rule);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/monthly-imports/:id/calculate-fees", async (req, res) => {
  try {
    const results = await calculateMonthlyImportFees(req.params.id, req.body?.changedBy || undefined);
    res.json(results);
  } catch (err) {
    handleError(res, err);
  }
});

// ---- Implementation Batch：月結首頁／待處理／問題確認／月結鎖定／自費代課／給出納／對帳 ----

function requireSemesterYearMonth(req: express.Request, res: express.Response): { semesterId: string; year: number; month: number } | null {
  const { semesterId, year, month } = req.query;
  if (!semesterId || !year || !month) {
    res.status(400).json({ error: "semesterId、year、month 為必填" });
    return null;
  }
  return { semesterId: String(semesterId), year: Number(year), month: Number(month) };
}

app.get("/api/monthly-dashboard", async (req, res) => {
  try {
    const params = requireSemesterYearMonth(req, res);
    if (!params) return;
    const dashboard = await getMonthlyDashboard(params.semesterId, params.year, params.month);
    res.json(dashboard);
  } catch (err) {
    handleError(res, err);
  }
});

app.get("/api/pending-issues", async (req, res) => {
  try {
    const params = requireSemesterYearMonth(req, res);
    if (!params) return;
    const issues = await listPendingIssues(params.semesterId, params.year, params.month);
    res.json(issues);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/issue-acknowledgements", async (req, res) => {
  try {
    const { semesterId, year, month, targetTable, targetId, reason, acknowledgedBy } = req.body;
    const ack = await acknowledgeIssue({ semesterId, year: Number(year), month: Number(month), targetTable, targetId, reason, acknowledgedBy });
    res.json(ack);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/issue-acknowledgements/revoke", async (req, res) => {
  try {
    const { targetTable, targetId, changedBy } = req.body;
    await revokeAcknowledgement(targetTable, targetId, changedBy || undefined);
    res.json({ ok: true });
  } catch (err) {
    handleError(res, err);
  }
});

app.get("/api/monthly-lock", async (req, res) => {
  try {
    const params = requireSemesterYearMonth(req, res);
    if (!params) return;
    const status = await getMonthlyLockStatus(params.semesterId, params.year, params.month);
    const blocking = await getBlockingIssues(params.semesterId, params.year, params.month);
    res.json({ ...status, blocking });
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/monthly-lock/lock", async (req, res) => {
  try {
    const { semesterId, year, month, lockedBy } = req.body;
    const status = await lockMonth(semesterId, Number(year), Number(month), lockedBy);
    res.json(status);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/monthly-lock/unlock", async (req, res) => {
  try {
    const { semesterId, year, month, unlockedBy, reason } = req.body;
    const status = await unlockMonth(semesterId, Number(year), Number(month), unlockedBy, reason);
    res.json(status);
  } catch (err) {
    handleError(res, err);
  }
});

app.get("/api/self-funded", async (req, res) => {
  try {
    const params = requireSemesterYearMonth(req, res);
    if (!params) return;
    const records = await listSelfFundedRecords(params.semesterId, params.year, params.month);
    res.json(records);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/self-funded", async (req, res) => {
  try {
    const body = req.body;
    const record = await createSelfFundedRecord({
      semesterId: body.semesterId,
      year: Number(body.year),
      month: Number(body.month),
      date: new Date(body.date),
      originalTeacherId: body.originalTeacherId || undefined,
      substituteTeacherId: body.substituteTeacherId,
      periodCode: body.periodCode || undefined,
      className: body.className || undefined,
      subject: body.subject || undefined,
      amount: body.amount,
      unitPrice: body.unitPrice || undefined,
      note: body.note || undefined,
      createdBy: body.createdBy,
    });
    res.json(record);
  } catch (err) {
    handleError(res, err);
  }
});

app.patch("/api/self-funded/:id", async (req, res) => {
  try {
    const body = req.body;
    const record = await updateSelfFundedRecord(req.params.id, {
      date: body.date ? new Date(body.date) : undefined,
      originalTeacherId: body.originalTeacherId,
      substituteTeacherId: body.substituteTeacherId,
      periodCode: body.periodCode,
      className: body.className,
      subject: body.subject,
      amount: body.amount,
      unitPrice: body.unitPrice,
      note: body.note,
      updatedBy: body.updatedBy,
    });
    res.json(record);
  } catch (err) {
    handleError(res, err);
  }
});

app.delete("/api/self-funded/:id", async (req, res) => {
  try {
    const { deletedBy, reason } = req.body;
    await deleteSelfFundedRecord(req.params.id, deletedBy, reason);
    res.json({ ok: true });
  } catch (err) {
    handleError(res, err);
  }
});

app.get("/api/chuna-summary", async (req, res) => {
  try {
    const idsParam = req.query.ids;
    const ids = typeof idsParam === "string" ? idsParam.split(",").filter(Boolean) : [];
    if (ids.length === 0) {
      res.status(400).json({ error: "請提供 ids（以逗號分隔的 monthlyImportId 清單）" });
      return;
    }
    const summary = await getChunaSummary(ids);
    res.json(summary);
  } catch (err) {
    handleError(res, err);
  }
});

app.get("/api/chuna-export", async (req, res) => {
  try {
    const idsParam = req.query.ids;
    const ids = typeof idsParam === "string" ? idsParam.split(",").filter(Boolean) : [];
    const year = Number(req.query.year);
    const month = Number(req.query.month);
    if (ids.length === 0 || !year || !month) {
      res.status(400).json({ error: "請提供 ids、year、month" });
      return;
    }
    const buffer = await generateChunaExcelBuffer(ids, year, month);
    // 檔名含中文，HTTP header 只能放 ASCII，中文檔名要用 RFC 5987 的 filename*= 語法，
    // 另外保留一個 ASCII fallback 檔名給不支援 filename* 的舊瀏覽器。
    const encodedFileName = encodeURIComponent(`${year}-${month}-給出納.xlsx`);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${year}-${month}-chuna.xlsx"; filename*=UTF-8''${encodedFileName}`);
    res.send(buffer);
  } catch (err) {
    handleError(res, err);
  }
});

app.post("/api/reconciliation", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      res.status(400).json({ error: "請上傳原始給出納 Excel 檔案" });
      return;
    }
    const idsParam = req.body.ids;
    const ids = typeof idsParam === "string" ? idsParam.split(",").filter(Boolean) : [];
    if (ids.length === 0) {
      res.status(400).json({ error: "請提供 ids（以逗號分隔的 monthlyImportId 清單）" });
      return;
    }
    const uploadedRows = await parseUploadedChunaWorkbook(req.file.buffer);
    const result = await reconcile(ids, uploadedRows);
    res.json(result);
  } catch (err) {
    handleError(res, err);
  }
});

const port = Number(process.env.PORT ?? 3000);
app.listen(port, () => {
  console.log(`管理介面已啟動：http://localhost:${port}`);
});
