// Implementation Batch：月結工作流程（MonthlyIssueAcknowledgement／MonthlyLock／
// 待處理／分類依據／給出納／對帳／自費代課）驗收測試。
// 完全不改 Phase1~9-5 已驗證的分類/計費核心邏輯，這裡只測本次新增的功能本身。
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import { prisma } from "../src/prismaClient";
import { createPerson } from "../src/services/personService";
import { createFeeRule } from "../src/services/feeRuleService";
import { importSubstituteExcel } from "../src/services/excelImportService";
import { classifySubstituteRecord, overrideClassification, describeClassificationBasis } from "../src/services/classificationService";
import { calculateSubstituteRecordFee } from "../src/services/feeCalculationService";
import { acknowledgeIssue, revokeAcknowledgement } from "../src/services/issueAcknowledgementService";
import { getBlockingIssues, lockMonth, unlockMonth, getMonthlyLockStatus, assertMonthNotLocked } from "../src/services/monthlyLockService";
import { getMonthlyDashboard } from "../src/services/dashboardService";
import { listPendingIssues } from "../src/services/pendingIssuesService";
import { createSelfFundedRecord, updateSelfFundedRecord } from "../src/services/selfFundedService";
import { getChunaSummary, generateChunaExcelBuffer } from "../src/services/chunaService";
import { reconcile } from "../src/services/reconciliationService";

const cleanupSemesterIds: string[] = [];
const cleanupPersonIds: string[] = [];
const cleanupLockKeys: { year: number; month: number }[] = [];

async function makeTestSemester(schoolYear: number) {
  const sem = await prisma.semester.create({
    data: { schoolYear, term: 1, startDate: new Date("2030-01-01"), endDate: new Date("2030-12-31") },
  });
  cleanupSemesterIds.push(sem.id);
  return sem;
}

async function makeTeacher(name: string) {
  const result = await createPerson({ name }, { forceCreate: true });
  if (result.status !== "CREATED") throw new Error("unreachable");
  cleanupPersonIds.push(result.person.id);
  return result.person;
}

// 用明顯不會跟真實資料或種子資料衝突的年份（8000+），避免撞到 2026-06（真實驗收月份）
// 或種子資料的示範鎖定（2026-05／2026-09）。
async function makeRecord(
  semesterId: string,
  overrides: {
    date: string;
    year: number;
    month: number;
    fundingSource?: "GENERAL" | "OVERTIME" | "PROJECT" | "UNDETERMINED";
    classificationMethod?: "WEEKLY_RULE" | "DATE_EXCEPTION" | "GENERAL_DEFAULT" | "MANUAL_OVERRIDE" | "CONFLICT" | "TEACHER_UNMATCHED";
    originalTeacherId?: string | null;
    originalTeacherText?: string | null;
    substituteTeacherId?: string | null;
    substituteTeacherText?: string | null;
    unitPrice?: number | null;
    amount?: number | null;
    substitutePeriodFeeText?: string | null;
    sourceStaffType?: "BD" | "NON_BD" | "UNKNOWN";
  }
) {
  cleanupLockKeys.push({ year: overrides.year, month: overrides.month });
  const monthlyImport = await prisma.monthlyImport.create({
    data: {
      semesterId,
      year: overrides.year,
      month: overrides.month,
      fileName: "test.xlsx",
      totalCount: 1,
      successCount: 1,
      sourceStaffType: overrides.sourceStaffType ?? "NON_BD",
    },
  });
  const raw = await prisma.substituteRecordRaw.create({
    data: {
      monthlyImportId: monthlyImport.id,
      rowNumber: 1,
      originalTeacherText: overrides.originalTeacherText ?? "測試教師",
      substituteTeacherText: overrides.substituteTeacherText ?? "測試代課教師",
      dateText: overrides.date,
      periodText: "P2",
      substitutePeriodFeeText: overrides.substitutePeriodFeeText ?? undefined,
    },
  });
  const record = await prisma.substituteRecord.create({
    data: {
      rawRecordId: raw.id,
      monthlyImportId: monthlyImport.id,
      originalTeacherId: overrides.originalTeacherId ?? undefined,
      substituteTeacherId: overrides.substituteTeacherId ?? undefined,
      date: new Date(overrides.date),
      weekday: "TUE",
      periodCode: "P2",
      staffType: overrides.sourceStaffType ?? "NON_BD",
      fundingSource: overrides.fundingSource ?? "GENERAL",
      classificationMethod: overrides.classificationMethod ?? "GENERAL_DEFAULT",
      unitPrice: overrides.unitPrice ?? undefined,
      amount: overrides.amount ?? undefined,
    },
  });
  return { record, monthlyImport };
}

beforeAll(async () => {
  await prisma.$connect();
});

afterAll(async () => {
  const imports = await prisma.monthlyImport.findMany({ where: { semesterId: { in: cleanupSemesterIds } } });
  const importIds = imports.map((i) => i.id);
  const records = await prisma.substituteRecord.findMany({ where: { monthlyImportId: { in: importIds } } });
  await prisma.changeLog.deleteMany({
    where: { OR: [{ recordId: { in: importIds } }, { recordId: { in: cleanupPersonIds } }, { recordId: { in: records.map((r) => r.id) } }] },
  });
  await prisma.monthlyIssueAcknowledgement.deleteMany({ where: { semesterId: { in: cleanupSemesterIds } } });
  await prisma.monthlyImportError.deleteMany({ where: { monthlyImportId: { in: importIds } } });
  await prisma.substituteRecord.deleteMany({ where: { monthlyImportId: { in: importIds } } });
  await prisma.substituteRecordRaw.deleteMany({ where: { monthlyImportId: { in: importIds } } });
  await prisma.monthlyImport.deleteMany({ where: { id: { in: importIds } } });
  await prisma.feeRule.deleteMany({ where: { semesterId: { in: cleanupSemesterIds } } });
  await prisma.specialWeeklyRule.deleteMany({ where: { semesterId: { in: cleanupSemesterIds } } });
  await prisma.specialDateRule.deleteMany({ where: { semesterId: { in: cleanupSemesterIds } } });
  await prisma.person.deleteMany({ where: { id: { in: cleanupPersonIds } } });
  for (const key of cleanupLockKeys) {
    await prisma.monthlyLock.deleteMany({ where: { year: key.year, month: key.month } });
  }
  await prisma.semester.deleteMany({ where: { id: { in: cleanupSemesterIds } } });
  await prisma.$disconnect();
});

describe("1-4. MonthlyIssueAcknowledgement：確認不是消除問題", () => {
  it("1/2/3. 建立確認後，原始問題（TEACHER_UNMATCHED）完全不變，但不再列入阻擋問題", async () => {
    const sem = await makeTestSemester(8901);
    const { record, monthlyImport } = await makeRecord(sem.id, {
      date: "2030-03-10", year: 2030, month: 3,
      fundingSource: "UNDETERMINED", classificationMethod: "TEACHER_UNMATCHED",
    });

    const before = await getBlockingIssues(sem.id, 2030, 3);
    expect(before.teacherUnmatched).toBe(1);

    const ack = await acknowledgeIssue({
      semesterId: sem.id, year: 2030, month: 3,
      targetTable: "SubstituteRecord", targetId: record.id,
      reason: "原始資料本身就無法配對，已跟學校確認", acknowledgedBy: "教學組長",
    });
    expect(ack.reason).toContain("已跟學校確認");

    // 原始問題完全沒有被改動
    const stillThere = await prisma.substituteRecord.findUniqueOrThrow({ where: { id: record.id } });
    expect(stillThere.classificationMethod).toBe("TEACHER_UNMATCHED");
    expect(stillThere.originalTeacherId).toBeNull();
    expect(stillThere.fundingSource).toBe("UNDETERMINED");

    const after = await getBlockingIssues(sem.id, 2030, 3);
    expect(after.teacherUnmatched).toBe(0);
    expect(after.total).toBe(0);

    // 撤銷確認後，問題重新變回待處理
    await revokeAcknowledgement("SubstituteRecord", record.id, "教學組長");
    const afterRevoke = await getBlockingIssues(sem.id, 2030, 3);
    expect(afterRevoke.teacherUnmatched).toBe(1);
  });

  it("不能無條件放行所有問題：只有明確建立確認紀錄的那一筆會被排除", async () => {
    const sem = await makeTestSemester(8902);
    const { record: r1 } = await makeRecord(sem.id, {
      date: "2030-03-10", year: 2030, month: 3,
      fundingSource: "UNDETERMINED", classificationMethod: "CONFLICT",
    });
    const { record: r2 } = await makeRecord(sem.id, {
      date: "2030-03-11", year: 2030, month: 3,
      fundingSource: "UNDETERMINED", classificationMethod: "CONFLICT",
    });

    await acknowledgeIssue({ semesterId: sem.id, year: 2030, month: 3, targetTable: "SubstituteRecord", targetId: r1.id, reason: "已確認", acknowledgedBy: "組長" });

    const issues = await getBlockingIssues(sem.id, 2030, 3);
    expect(issues.conflict).toBe(1); // 只有 r1 被確認，r2 仍然阻擋
  });
});

describe("4-8. MonthlyLock：鎖定條件與保護", () => {
  it("4. 未確認的阻擋問題存在時，鎖定會被拒絕", async () => {
    const sem = await makeTestSemester(8903);
    await makeRecord(sem.id, { date: "2030-04-10", year: 2030, month: 4, fundingSource: "UNDETERMINED", classificationMethod: "TEACHER_UNMATCHED" });

    await expect(lockMonth(sem.id, 2030, 4, "教學組長")).rejects.toThrow(/未確認的阻擋問題/);
  });

  it("5. 所有阻擋問題都解決或確認後，可以成功鎖定", async () => {
    const sem = await makeTestSemester(8904);
    const { record } = await makeRecord(sem.id, { date: "2030-05-10", year: 2030, month: 5, fundingSource: "UNDETERMINED", classificationMethod: "TEACHER_UNMATCHED" });
    await acknowledgeIssue({ semesterId: sem.id, year: 2030, month: 5, targetTable: "SubstituteRecord", targetId: record.id, reason: "已確認無法處理", acknowledgedBy: "組長" });

    const status = await lockMonth(sem.id, 2030, 5, "教學組長");
    expect(status.isLocked).toBe(true);
    expect(status.lockedBy).toBe("教學組長");
  });

  it("6. 鎖定後，匯入／分類覆寫／費用計算／自費代課建立都會被擋下", async () => {
    const sem = await makeTestSemester(8905);
    const teacher = await makeTeacher("鎖定測試教師");
    const sub = await makeTeacher("鎖定測試代課教師");
    await createFeeRule({ semesterId: sem.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2030-01-01") });
    const { record } = await makeRecord(sem.id, {
      date: "2030-06-10", year: 2030, month: 6,
      originalTeacherId: teacher.id, substituteTeacherId: sub.id,
      fundingSource: "GENERAL", classificationMethod: "GENERAL_DEFAULT",
    });
    await calculateSubstituteRecordFee(record.id); // 先算好金額，否則會被「金額無法計算」擋住鎖定，這裡要測的是鎖定之後的保護

    await lockMonth(sem.id, 2030, 6, "教學組長");

    await expect(
      importSubstituteExcel({ semesterId: sem.id, year: 2030, month: 6, fileName: "x.xlsx", fileBuffer: Buffer.from(""), importedBy: "test" })
    ).rejects.toThrow(/已經鎖定/);

    await expect(
      overrideClassification(record.id, { fundingSource: "GENERAL" }, "test", "測試鎖定")
    ).rejects.toThrow(/已經鎖定/);

    await expect(calculateSubstituteRecordFee(record.id)).rejects.toThrow(/已經鎖定/);

    await expect(
      createSelfFundedRecord({
        semesterId: sem.id, year: 2030, month: 6, date: new Date("2030-06-15"),
        substituteTeacherId: sub.id, amount: 500, createdBy: "test",
      })
    ).rejects.toThrow(/已經鎖定/);
  });

  it("7. 解鎖必須填寫理由", async () => {
    const sem = await makeTestSemester(8906);
    cleanupLockKeys.push({ year: 2030, month: 7 });
    await lockMonth(sem.id, 2030, 7, "組長");
    await expect(unlockMonth(sem.id, 2030, 7, "組長", "")).rejects.toThrow(/理由/);
  });

  it("8. 解鎖成功後留下 ChangeLog，且月份恢復可修改", async () => {
    const sem = await makeTestSemester(8907);
    cleanupLockKeys.push({ year: 2030, month: 8 });
    const lockResult = await lockMonth(sem.id, 2030, 8, "組長");

    const unlockResult = await unlockMonth(sem.id, 2030, 8, "教學組長", "學校要求補登一筆遺漏資料");
    expect(unlockResult.isLocked).toBe(false);

    const logs = await prisma.changeLog.findMany({ where: { tableName: "monthly_locks" }, orderBy: { createdAt: "desc" } });
    const unlockLog = logs.find((l) => l.reason?.includes("解鎖") && l.reason?.includes("補登一筆遺漏資料"));
    expect(unlockLog).toBeDefined();

    // 解鎖後恢復可修改，不會再拋出鎖定錯誤
    await expect(assertMonthNotLocked(2030, 8)).resolves.toBeUndefined();
  });
});

describe("9. 月結首頁統計即時計算", () => {
  it("涵蓋匯入／分類／費用／問題／鎖定狀態，且不因重複查詢而改變", async () => {
    const sem = await makeTestSemester(8908);
    await createFeeRule({ semesterId: sem.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2030-01-01") });
    const teacher = await makeTeacher("首頁測試教師");
    const sub = await makeTeacher("首頁測試代課");

    const { record: r1 } = await makeRecord(sem.id, {
      date: "2030-09-01", year: 2030, month: 9, originalTeacherId: teacher.id, substituteTeacherId: sub.id,
      fundingSource: "GENERAL", classificationMethod: "GENERAL_DEFAULT",
    });
    await calculateSubstituteRecordFee(r1.id);
    await makeRecord(sem.id, { date: "2030-09-02", year: 2030, month: 9, fundingSource: "UNDETERMINED", classificationMethod: "TEACHER_UNMATCHED" });

    const dashboard = await getMonthlyDashboard(sem.id, 2030, 9);
    expect(dashboard.classification.general).toBe(1);
    expect(dashboard.classification.teacherUnmatched).toBe(1);
    expect(dashboard.fee.calculatedCount).toBe(1);
    expect(dashboard.fee.notCalculatedCount).toBe(1);
    expect(dashboard.fee.totalAmount).toBe("405");
    expect(dashboard.issues.blocking.teacherUnmatched).toBe(1);
    expect(dashboard.lock.isLocked).toBe(false);
  });
});

describe("10. 待處理統一查詢", () => {
  it("整合 TEACHER_UNMATCHED／CONFLICT／AMOUNT_MISSING／匯入錯誤，並正確反映確認狀態", async () => {
    const sem = await makeTestSemester(8909);
    const { record: unmatched } = await makeRecord(sem.id, { date: "2030-10-01", year: 2030, month: 10, fundingSource: "UNDETERMINED", classificationMethod: "TEACHER_UNMATCHED" });
    const { record: conflict } = await makeRecord(sem.id, { date: "2030-10-02", year: 2030, month: 10, fundingSource: "UNDETERMINED", classificationMethod: "CONFLICT" });
    const { record: amountMissing, monthlyImport } = await makeRecord(sem.id, { date: "2030-10-03", year: 2030, month: 10, fundingSource: "GENERAL", classificationMethod: "GENERAL_DEFAULT", amount: null });
    const importError = await prisma.monthlyImportError.create({ data: { monthlyImportId: monthlyImport.id, rowNumber: 5, fieldName: "節次", message: "無法解析節次：導師時間" } });

    await acknowledgeIssue({ semesterId: sem.id, year: 2030, month: 10, targetTable: "SubstituteRecord", targetId: unmatched.id, reason: "已知問題", acknowledgedBy: "組長" });

    const issues = await listPendingIssues(sem.id, 2030, 10);
    const byId = new Map(issues.map((i) => [i.targetId, i]));

    expect(byId.get(unmatched.id)?.status).toBe("ACKNOWLEDGED");
    expect(byId.get(conflict.id)?.status).toBe("PENDING");
    expect(byId.get(amountMissing.id)?.issueType).toBe("AMOUNT_MISSING");
    expect(byId.get(importError.id)?.issueType).toBe("IMPORT_ERROR");
    expect(byId.get(importError.id)?.description).toContain("導師時間");
  });
});

describe("11. 分類依據人話化", () => {
  it("每週規則／單日例外／一般公費／規則衝突／原教師未配對，都轉成看得懂的文字", async () => {
    const sem = await makeTestSemester(8910);
    const teacher = await makeTeacher("分類依據測試教師");
    const rule = await prisma.specialWeeklyRule.create({
      data: { semesterId: sem.id, personId: teacher.id, ruleType: "OVERTIME", weekday: "TUE", periodCode: "P4", weeklyPeriods: 1, effectiveDate: new Date("2030-01-01") },
    });
    const dateRule = await prisma.specialDateRule.create({
      data: { semesterId: sem.id, date: new Date("2030-06-16"), personId: teacher.id, periodCode: "P4", overrideClassification: "PROJECT" },
    });

    const weeklyText = await describeClassificationBasis({
      classificationMethod: "WEEKLY_RULE", classificationRuleId: rule.id, fundingSource: "OVERTIME", projectId: null, conflictCandidatesJson: null,
    });
    expect(weeklyText).toContain("星期二");
    expect(weeklyText).toContain("P4");
    expect(weeklyText).toContain("超鐘點");

    const dateText = await describeClassificationBasis({
      classificationMethod: "DATE_EXCEPTION", classificationRuleId: dateRule.id, fundingSource: "PROJECT", projectId: null, conflictCandidatesJson: null,
    });
    expect(dateText).toContain("2030-06-16");
    expect(dateText).toContain("專案");

    const generalText = await describeClassificationBasis({
      classificationMethod: "GENERAL_DEFAULT", classificationRuleId: null, fundingSource: "GENERAL", projectId: null, conflictCandidatesJson: null,
    });
    expect(generalText).toContain("沒有符合特殊規則");

    const conflictText = await describeClassificationBasis({
      classificationMethod: "CONFLICT", classificationRuleId: null, fundingSource: "UNDETERMINED", projectId: null,
      conflictCandidatesJson: JSON.stringify({ candidates: [{ ruleId: "a" }, { ruleId: "b" }] }),
    });
    expect(conflictText).toContain("2");
    expect(conflictText).toContain("人工確認");

    const unmatchedText = await describeClassificationBasis({
      classificationMethod: "TEACHER_UNMATCHED", classificationRuleId: null, fundingSource: "UNDETERMINED", projectId: null, conflictCandidatesJson: null,
    });
    expect(unmatchedText).toContain("尚未配對");
  });
});

describe("12-13. 給出納彙總與 Excel 匯出", () => {
  it("12. 依 BD／非BD 分組彙總，並帶出薪資代碼", async () => {
    const sem = await makeTestSemester(8911);
    await createFeeRule({ semesterId: sem.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2030-01-01") });
    const created = await createPerson({ name: "給出納測試代課教師", payrollCode: "T9001" }, { forceCreate: true });
    if (created.status !== "CREATED") throw new Error("unreachable");
    cleanupPersonIds.push(created.person.id);

    const { record: bdRecord, monthlyImport: bdImport } = await makeRecord(sem.id, {
      date: "2030-11-01", year: 2030, month: 11, substituteTeacherId: created.person.id, sourceStaffType: "BD",
      fundingSource: "GENERAL", classificationMethod: "GENERAL_DEFAULT",
    });
    await calculateSubstituteRecordFee(bdRecord.id);

    const summary = await getChunaSummary([bdImport.id]);
    expect(summary.bd.length).toBe(1);
    expect(summary.bd[0].payrollCode).toBe("T9001");
    expect(summary.bd[0].totalAmount).toBe("405");
    expect(summary.nonBd.length).toBe(0);
  });

  it("13. Excel 匯出產生可讀取的檔案，內容與彙總一致", async () => {
    const sem = await makeTestSemester(8912);
    await createFeeRule({ semesterId: sem.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2030-01-01") });
    const sub = await makeTeacher("匯出測試代課教師");
    const { record, monthlyImport } = await makeRecord(sem.id, {
      date: "2030-12-01", year: 2030, month: 12, substituteTeacherId: sub.id, sourceStaffType: "NON_BD",
      fundingSource: "GENERAL", classificationMethod: "GENERAL_DEFAULT",
    });
    await calculateSubstituteRecordFee(record.id);

    const buffer = await generateChunaExcelBuffer([monthlyImport.id], 2030, 12);
    expect(buffer.length).toBeGreaterThan(0);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer as any);
    const ws = wb.worksheets[0];
    let found = false;
    ws.eachRow((row) => {
      const vals = row.values as any[];
      if (vals.includes("匯出測試代課教師")) found = true;
    });
    expect(found).toBe(true);
  });
});

describe("14-15. 對帳", () => {
  it("14. 系統與原始資料完全一致時，狀態為 MATCH，差額為 0", async () => {
    const sem = await makeTestSemester(8913);
    await createFeeRule({ semesterId: sem.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2030-01-01") });
    const sub = await makeTeacher("對帳一致測試教師");
    const { record, monthlyImport } = await makeRecord(sem.id, {
      date: "2031-01-01", year: 2031, month: 1, substituteTeacherId: sub.id,
      fundingSource: "GENERAL", classificationMethod: "GENERAL_DEFAULT",
    });
    await calculateSubstituteRecordFee(record.id);

    const result = await reconcile([monthlyImport.id], [{ name: "對帳一致測試教師", periodCount: 1, amount: 405 }]);
    const row = result.rows.find((r) => r.name === "對帳一致測試教師")!;
    expect(row.status).toBe("MATCH");
    expect(row.amountDiff).toBe(0);
    expect(result.totals.diff).toBe(0);
  });

  it("15. 有差異時列出金額差距與可能原因，且不宣稱已確認原因", async () => {
    const sem = await makeTestSemester(8914);
    await createFeeRule({ semesterId: sem.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2030-01-01") });
    const sub = await makeTeacher("對帳差異測試教師");
    const { record, monthlyImport } = await makeRecord(sem.id, {
      date: "2031-02-01", year: 2031, month: 2, substituteTeacherId: sub.id,
      fundingSource: "GENERAL", classificationMethod: "GENERAL_DEFAULT",
    });
    await calculateSubstituteRecordFee(record.id);

    const result = await reconcile([monthlyImport.id], [{ name: "對帳差異測試教師", periodCount: 3, amount: 1215 }]);
    const row = result.rows.find((r) => r.name === "對帳差異測試教師")!;
    expect(row.status).toBe("SYSTEM_LESS");
    expect(row.amountDiff).toBe(405 - 1215);
    expect(row.possibleReason).toBeTruthy();
    // 只講「可能原因」，不應該出現「已確認」這種字眼
    expect(row.possibleReason).not.toContain("已確認原因");

    const onlyOriginal = await reconcile([monthlyImport.id], [{ name: "根本不存在的教師", periodCount: 2, amount: 810 }]);
    const orphanRow = onlyOriginal.rows.find((r) => r.name === "根本不存在的教師")!;
    expect(orphanRow.status).toBe("ONLY_ORIGINAL");
  });
});

describe("16-20. 自費代課", () => {
  it("16/18. 建立自費代課不需要 rawRecordId，entryType 正確標示來源", async () => {
    const sem = await makeTestSemester(8915);
    const sub = await makeTeacher("自費代課測試教師");
    const record = await createSelfFundedRecord({
      semesterId: sem.id, year: 2031, month: 3, date: new Date("2031-03-15"),
      substituteTeacherId: sub.id, amount: 800, note: "臨時自費請人代課", createdBy: "教學組長",
    });
    cleanupLockKeys.push({ year: 2031, month: 3 });

    expect(record.entryType).toBe("MANUAL_SELF_FUNDED");
    expect(record.rawRecordId).toBeNull();
    expect(record.amount?.toString()).toBe("800");
    expect(record.createdBy).toBe("教學組長");
  });

  it("17. 修改自費代課會更新資料並留下 ChangeLog", async () => {
    const sem = await makeTestSemester(8916);
    const sub = await makeTeacher("自費代課修改測試教師");
    const record = await createSelfFundedRecord({
      semesterId: sem.id, year: 2031, month: 4, date: new Date("2031-04-01"),
      substituteTeacherId: sub.id, amount: 500, createdBy: "組長",
    });
    cleanupLockKeys.push({ year: 2031, month: 4 });

    const updated = await updateSelfFundedRecord(record.id, { amount: 600, updatedBy: "教學組長" });
    expect(updated.amount?.toString()).toBe("600");
    expect(updated.updatedBy).toBe("教學組長");

    const logs = await prisma.changeLog.findMany({ where: { tableName: "substitute_records", recordId: record.id } });
    expect(logs.some((l) => l.reason === "修改自費代課")).toBe(true);
  });

  it("19. 公費 Excel 匯入的紀錄仍然一定要有 rawRecordId", async () => {
    const sem = await makeTestSemester(8917);
    const teacher = await makeTeacher("公費仍需rawRecordId測試教師");
    const { record } = await makeRecord(sem.id, {
      date: "2031-05-01", year: 2031, month: 5, originalTeacherId: teacher.id,
      fundingSource: "GENERAL", classificationMethod: "GENERAL_DEFAULT",
    });
    expect(record.entryType).toBe("EXCEL_IMPORT");
    expect(record.rawRecordId).not.toBeNull();
  });

  it("20. 自費代課不會污染既有公費分類統計與費用計算", async () => {
    const sem = await makeTestSemester(8918);
    await createFeeRule({ semesterId: sem.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2030-01-01") });
    const teacher = await makeTeacher("污染測試原教師");
    const sub1 = await makeTeacher("污染測試公費代課教師");
    const sub2 = await makeTeacher("污染測試自費代課教師");

    const { record: publicRecord, monthlyImport } = await makeRecord(sem.id, {
      date: "2031-06-01", year: 2031, month: 6, originalTeacherId: teacher.id, substituteTeacherId: sub1.id,
      fundingSource: "GENERAL", classificationMethod: "GENERAL_DEFAULT",
    });
    await calculateSubstituteRecordFee(publicRecord.id);

    await createSelfFundedRecord({
      semesterId: sem.id, year: 2031, month: 6, date: new Date("2031-06-02"),
      substituteTeacherId: sub2.id, amount: 1000, createdBy: "組長",
    });

    // 公費批次的計算結果、分類統計完全不受自費代課影響（自費是另一個 monthlyImportId）
    const refreshedPublic = await prisma.substituteRecord.findUniqueOrThrow({ where: { id: publicRecord.id } });
    expect(refreshedPublic.amount?.toString()).toBe("405");
    expect(refreshedPublic.fundingSource).toBe("GENERAL");

    const dashboard = await getMonthlyDashboard(sem.id, 2031, 6);
    // 分類統計只算 entryType=EXCEL_IMPORT，自費那 1 筆不會被算進 GENERAL
    expect(dashboard.classification.general).toBe(1);
    expect(dashboard.selfFunded.count).toBe(1);
    expect(dashboard.fee.totalAmount).toBe("405"); // 公費批次彙總不含自費批次

    // 待處理清單也不會把自費代課當成問題
    const issues = await listPendingIssues(sem.id, 2031, 6);
    expect(issues.length).toBe(0);
  });
});
