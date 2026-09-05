// Phase 9 第一階段：節次型公費代課費用計算引擎
// 只驗證 GENERAL/OVERTIME/PROJECT 三種節次型費用的計算與月彙總，
// 完全不涉及代導師費／日薪／半日薪／自費代課／專案上限。
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { prisma } from "../src/prismaClient";
import { createPerson } from "../src/services/personService";
import { createFeeRule } from "../src/services/feeRuleService";
import { calculateSubstituteRecordFee, calculateMonthlyImportFees, summarizeTeacherMonthlyFees } from "../src/services/feeCalculationService";

const cleanupSemesterIds: string[] = [];
const cleanupPersonIds: string[] = [];

async function makeTestSemester(schoolYear: number) {
  const sem = await prisma.semester.create({
    data: { schoolYear, term: 1, startDate: new Date("2025-01-01"), endDate: new Date("2025-12-31") },
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

async function makeRecord(
  semesterId: string,
  overrides: {
    date: string;
    fundingSource?: "GENERAL" | "OVERTIME" | "PROJECT" | "UNDETERMINED";
    classificationMethod?: "WEEKLY_RULE" | "DATE_EXCEPTION" | "GENERAL_DEFAULT" | "MANUAL_OVERRIDE" | "CONFLICT" | "TEACHER_UNMATCHED";
    substituteTeacherId?: string | null;
    substituteTeacherText?: string | null;
    periodCount?: number | null;
    unitPrice?: number | null;
    amount?: number | null;
  }
) {
  const monthlyImport = await prisma.monthlyImport.create({
    data: { semesterId, year: 2025, month: Number(overrides.date.slice(5, 7)), fileName: "test.xlsx", totalCount: 1, successCount: 1 },
  });
  const raw = await prisma.substituteRecordRaw.create({
    data: {
      monthlyImportId: monthlyImport.id,
      rowNumber: 1,
      originalTeacherText: "測試教師",
      substituteTeacherText: overrides.substituteTeacherText ?? null,
      dateText: overrides.date,
      periodText: "P2",
    },
  });
  const record = await prisma.substituteRecord.create({
    data: {
      rawRecordId: raw.id,
      monthlyImportId: monthlyImport.id,
      substituteTeacherId: overrides.substituteTeacherId ?? null,
      date: new Date(overrides.date),
      weekday: "TUE",
      periodCode: "P2",
      staffType: "UNKNOWN",
      fundingSource: overrides.fundingSource ?? "GENERAL",
      classificationMethod: overrides.classificationMethod ?? "GENERAL_DEFAULT",
      periodCount: overrides.periodCount ?? undefined,
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
  await prisma.substituteRecord.deleteMany({ where: { monthlyImportId: { in: importIds } } });
  await prisma.substituteRecordRaw.deleteMany({ where: { monthlyImportId: { in: importIds } } });
  await prisma.monthlyImport.deleteMany({ where: { id: { in: importIds } } });
  await prisma.feeRule.deleteMany({ where: { semesterId: { in: cleanupSemesterIds } } });
  await prisma.person.deleteMany({ where: { id: { in: cleanupPersonIds } } });
  await prisma.semester.deleteMany({ where: { id: { in: cleanupSemesterIds } } });
  await prisma.$disconnect();
});

describe("Phase 9 第一階段：節次型費用計算", () => {
  it("1/2/3. GENERAL/OVERTIME/PROJECT 各自套用正確費率，都是 405", async () => {
    const semester = await makeTestSemester(9101);
    await createFeeRule({ semesterId: semester.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2025-01-01") });
    await createFeeRule({ semesterId: semester.id, feeType: "OVERTIME_PERIOD", amount: 405, effectiveDate: new Date("2025-01-01") });

    const { record: generalRecord } = await makeRecord(semester.id, { date: "2025-03-10", fundingSource: "GENERAL" });
    const generalResult = await calculateSubstituteRecordFee(generalRecord.id);
    expect(generalResult.amount).toBe("405");
    expect(generalResult.unitPrice).toBe("405");

    const { record: overtimeRecord } = await makeRecord(semester.id, { date: "2025-03-10", fundingSource: "OVERTIME", classificationMethod: "WEEKLY_RULE" });
    const overtimeResult = await calculateSubstituteRecordFee(overtimeRecord.id);
    expect(overtimeResult.amount).toBe("405");

    const { record: projectRecord } = await makeRecord(semester.id, { date: "2025-03-10", fundingSource: "PROJECT", classificationMethod: "WEEKLY_RULE" });
    const projectResult = await calculateSubstituteRecordFee(projectRecord.id);
    expect(projectResult.amount).toBe("405");
  });

  it("4. 一筆 SubstituteRecord 只計 405，不再乘節數（即使 periodCount 被設成其他值）", async () => {
    const semester = await makeTestSemester(9102);
    await createFeeRule({ semesterId: semester.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2025-01-01") });

    const { record } = await makeRecord(semester.id, { date: "2025-03-10", fundingSource: "GENERAL", periodCount: 3 });
    const result = await calculateSubstituteRecordFee(record.id);
    expect(result.amount).toBe("405"); // 不是 1215
  });

  it("5. 同教師多筆節次可以正確累加（月彙總）", async () => {
    const semester = await makeTestSemester(9103);
    await createFeeRule({ semesterId: semester.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2025-01-01") });
    await createFeeRule({ semesterId: semester.id, feeType: "OVERTIME_PERIOD", amount: 405, effectiveDate: new Date("2025-01-01") });
    const teacher = await makeTeacher("彙總測試代課教師");

    const { monthlyImport: imp1 } = await makeRecord(semester.id, { date: "2025-03-01", fundingSource: "GENERAL", substituteTeacherId: teacher.id });
    const { monthlyImport: imp2, record: r2 } = await makeRecord(semester.id, { date: "2025-03-02", fundingSource: "GENERAL", substituteTeacherId: teacher.id });
    const { monthlyImport: imp3, record: r3 } = await makeRecord(semester.id, { date: "2025-03-03", fundingSource: "OVERTIME", classificationMethod: "WEEKLY_RULE", substituteTeacherId: teacher.id });

    const allImportIds = [imp1.id, imp2.id, imp3.id];
    for (const id of allImportIds) {
      await calculateMonthlyImportFees(id);
    }

    const summary = await summarizeTeacherMonthlyFees(allImportIds);
    const row = summary.find((s) => s.substituteTeacherId === teacher.id);
    expect(row).toBeDefined();
    expect(row!.generalCount).toBe(2);
    expect(row!.generalAmount).toBe("810");
    expect(row!.overtimeCount).toBe(1);
    expect(row!.overtimeAmount).toBe("405");
    expect(row!.totalCount).toBe(3);
    expect(row!.totalAmount).toBe("1215");
  });

  it("6/7/8. CONFLICT／TEACHER_UNMATCHED／UNDETERMINED 都不計算，unitPrice/amount 維持 null", async () => {
    const semester = await makeTestSemester(9104);
    await createFeeRule({ semesterId: semester.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2025-01-01") });

    const { record: conflictRecord } = await makeRecord(semester.id, { date: "2025-03-10", fundingSource: "UNDETERMINED", classificationMethod: "CONFLICT" });
    const conflictResult = await calculateSubstituteRecordFee(conflictRecord.id);
    expect(conflictResult.amount).toBeNull();
    expect(conflictResult.unitPrice).toBeNull();
    expect(conflictResult.skippedReason).toContain("UNDETERMINED");

    const { record: unmatchedRecord } = await makeRecord(semester.id, { date: "2025-03-10", fundingSource: "UNDETERMINED", classificationMethod: "TEACHER_UNMATCHED" });
    const unmatchedResult = await calculateSubstituteRecordFee(unmatchedRecord.id);
    expect(unmatchedResult.amount).toBeNull();

    const { record: undeterminedRecord } = await makeRecord(semester.id, { date: "2025-03-10", fundingSource: "UNDETERMINED", classificationMethod: "GENERAL_DEFAULT" });
    const undeterminedResult = await calculateSubstituteRecordFee(undeterminedRecord.id);
    expect(undeterminedResult.amount).toBeNull();

    // 確認資料庫裡真的是 null，不是只有回傳值是 null
    const dbRecord = await prisma.substituteRecord.findUniqueOrThrow({ where: { id: conflictRecord.id } });
    expect(dbRecord.amount).toBeNull();
    expect(dbRecord.unitPrice).toBeNull();
  });

  it("6b. 先前算好金額後，若重新分類變成 CONFLICT，金額會被清成 null，不會留著舊值誤導", async () => {
    const semester = await makeTestSemester(9105);
    await createFeeRule({ semesterId: semester.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2025-01-01") });

    const { record } = await makeRecord(semester.id, { date: "2025-03-10", fundingSource: "GENERAL", unitPrice: 405, amount: 405 });
    // 模擬重新分類後變成衝突
    await prisma.substituteRecord.update({ where: { id: record.id }, data: { fundingSource: "UNDETERMINED", classificationMethod: "CONFLICT" } });

    const result = await calculateSubstituteRecordFee(record.id);
    expect(result.amount).toBeNull();
    const dbRecord = await prisma.substituteRecord.findUniqueOrThrow({ where: { id: record.id } });
    expect(dbRecord.amount).toBeNull();
    expect(dbRecord.unitPrice).toBeNull();
  });

  it("9. FeeRule 生效日期正確：早於生效日期的紀錄找不到費率，不猜", async () => {
    const semester = await makeTestSemester(9106);
    await createFeeRule({ semesterId: semester.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2025-06-01") });

    const { record: beforeRecord } = await makeRecord(semester.id, { date: "2025-03-10", fundingSource: "GENERAL" });
    const beforeResult = await calculateSubstituteRecordFee(beforeRecord.id);
    expect(beforeResult.amount).toBeNull();
    expect(beforeResult.skippedReason).toContain("找不到");

    const { record: afterRecord } = await makeRecord(semester.id, { date: "2025-06-15", fundingSource: "GENERAL" });
    const afterResult = await calculateSubstituteRecordFee(afterRecord.id);
    expect(afterResult.amount).toBe("405");
  });

  it("10. FeeRule 版本切換後，歷史日期仍使用舊費率", async () => {
    const semester = await makeTestSemester(9107);
    const v1 = await createFeeRule({
      semesterId: semester.id,
      feeType: "SUBSTITUTE_PERIOD",
      amount: 405,
      effectiveDate: new Date("2025-01-01"),
      endDate: new Date("2025-05-31"),
    });
    const v2 = await createFeeRule({
      semesterId: semester.id,
      feeType: "SUBSTITUTE_PERIOD",
      amount: 450,
      effectiveDate: new Date("2025-06-01"),
    });

    const { record: oldRecord } = await makeRecord(semester.id, { date: "2025-03-01", fundingSource: "GENERAL" });
    const oldResult = await calculateSubstituteRecordFee(oldRecord.id);
    expect(oldResult.amount).toBe("405");
    expect(oldResult.feeRuleId).toBe(v1.id);

    const { record: newRecord } = await makeRecord(semester.id, { date: "2025-06-15", fundingSource: "GENERAL" });
    const newResult = await calculateSubstituteRecordFee(newRecord.id);
    expect(newResult.amount).toBe("450");
    expect(newResult.feeRuleId).toBe(v2.id);
  });

  it("11. PROJECT 使用 SUBSTITUTE_PERIOD 這張費率表，不是獨立費率", async () => {
    const semester = await makeTestSemester(9108);
    const rule = await createFeeRule({ semesterId: semester.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2025-01-01") });

    const { record } = await makeRecord(semester.id, { date: "2025-03-10", fundingSource: "PROJECT", classificationMethod: "WEEKLY_RULE" });
    const result = await calculateSubstituteRecordFee(record.id);
    expect(result.feeRuleId).toBe(rule.id);
  });

  it("額外：找不到費率時回報 skippedReason，方便管理者知道要先建費率", async () => {
    const semester = await makeTestSemester(9109);
    // 故意不建立任何 FeeRule
    const { record } = await makeRecord(semester.id, { date: "2025-03-10", fundingSource: "OVERTIME", classificationMethod: "WEEKLY_RULE" });
    const result = await calculateSubstituteRecordFee(record.id);
    expect(result.amount).toBeNull();
    expect(result.skippedReason).toContain("OVERTIME_PERIOD");
  });

  it("額外：月彙總不會漏掉代課教師尚未配對的紀錄", async () => {
    const semester = await makeTestSemester(9110);
    await createFeeRule({ semesterId: semester.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2025-01-01") });
    const { monthlyImport, record } = await makeRecord(semester.id, {
      date: "2025-03-10",
      fundingSource: "GENERAL",
      substituteTeacherId: null,
      substituteTeacherText: "未配對的某老師",
    });
    await calculateSubstituteRecordFee(record.id);

    const summary = await summarizeTeacherMonthlyFees([monthlyImport.id]);
    expect(summary.length).toBe(1);
    expect(summary[0].substituteTeacherId).toBeNull();
    expect(summary[0].substituteTeacherName).toContain("未配對");
    expect(summary[0].totalAmount).toBe("405");
  });
});
