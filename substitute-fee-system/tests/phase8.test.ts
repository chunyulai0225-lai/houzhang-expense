// Phase 8（第一階段）驗收測試：公費代課自動分類引擎
// 只驗證「判成哪一種來源、依據哪條規則」，完全不驗證任何金額計算。
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { prisma } from "../src/prismaClient";
import { createPerson } from "../src/services/personService";
import { createProject } from "../src/services/projectService";
import {
  classifyMonthlyImport,
  classifySubstituteRecord,
  computeClassification,
  listClassificationPreview,
  overrideClassification,
  revertToAutoClassification,
} from "../src/services/classificationService";

const cleanupSemesterIds: string[] = [];
const cleanupPersonIds: string[] = [];

async function makeTestSemester(
  schoolYear: number,
  overtimeMatchMode: "TEACHER_WEEKDAY_PERIOD" | "TEACHER_WEEKDAY_PERIOD_SUBJECT" = "TEACHER_WEEKDAY_PERIOD"
) {
  const sem = await prisma.semester.create({
    data: {
      schoolYear,
      term: 1,
      startDate: new Date("2026-08-01"),
      endDate: new Date("2027-01-31"),
      overtimeMatchMode,
    },
  });
  cleanupSemesterIds.push(sem.id);
  return sem;
}

async function makeTeacher(name: string) {
  const result = await createPerson({ name });
  if (result.status !== "CREATED") throw new Error("unreachable");
  cleanupPersonIds.push(result.person.id);
  return result.person;
}

async function makeRecord(
  semesterId: string,
  overrides: {
    originalTeacherId?: string | null;
    date: string;
    weekday: "MON" | "TUE" | "WED" | "THU" | "FRI";
    periodCode?: string;
    subject?: string | null;
    className?: string;
    staffType?: "BD" | "NON_BD" | "UNKNOWN";
  }
) {
  const monthlyImport = await prisma.monthlyImport.create({
    data: { semesterId, year: 2026, month: 9, fileName: "test.xlsx", totalCount: 1, successCount: 1 },
  });
  const raw = await prisma.substituteRecordRaw.create({
    data: {
      monthlyImportId: monthlyImport.id,
      rowNumber: 1,
      originalTeacherText: "測試教師",
      dateText: overrides.date,
      periodText: overrides.periodCode ?? "P2",
    },
  });
  const record = await prisma.substituteRecord.create({
    data: {
      rawRecordId: raw.id,
      monthlyImportId: monthlyImport.id,
      originalTeacherId: overrides.originalTeacherId ?? null,
      date: new Date(overrides.date),
      weekday: overrides.weekday,
      periodCode: overrides.periodCode ?? "P2",
      subject: overrides.subject ?? undefined,
      className: overrides.className,
      staffType: overrides.staffType ?? "UNKNOWN",
    },
  });
  return record;
}

beforeAll(async () => {
  await prisma.$connect();
});

afterAll(async () => {
  const imports = await prisma.monthlyImport.findMany({ where: { semesterId: { in: cleanupSemesterIds } } });
  const importIds = imports.map((i) => i.id);
  await prisma.changeLog.deleteMany({
    where: { OR: [{ recordId: { in: importIds } }, { recordId: { in: cleanupPersonIds } }] },
  });
  const records = await prisma.substituteRecord.findMany({ where: { monthlyImportId: { in: importIds } } });
  await prisma.changeLog.deleteMany({ where: { recordId: { in: records.map((r) => r.id) } } });
  await prisma.substituteRecord.deleteMany({ where: { monthlyImportId: { in: importIds } } });
  await prisma.substituteRecordRaw.deleteMany({ where: { monthlyImportId: { in: importIds } } });
  await prisma.monthlyImport.deleteMany({ where: { id: { in: importIds } } });
  await prisma.specialDateRule.deleteMany({ where: { semesterId: { in: cleanupSemesterIds } } });
  await prisma.specialWeeklyRule.deleteMany({ where: { semesterId: { in: cleanupSemesterIds } } });
  await prisma.project.deleteMany({ where: { semesterId: { in: cleanupSemesterIds } } });
  await prisma.person.deleteMany({ where: { id: { in: cleanupPersonIds } } });
  await prisma.semester.deleteMany({ where: { id: { in: cleanupSemesterIds } } });
  await prisma.$disconnect();
});

describe("1. 沒有特殊規則 → GENERAL", () => {
  it("沒有任何規則時分類為一般公費", async () => {
    const sem = await makeTestSemester(701);
    const teacher = await makeTeacher("測試分類教師甲");
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });

    const updated = await classifySubstituteRecord(record.id);
    expect(updated.fundingSource).toBe("GENERAL");
    expect(updated.classificationMethod).toBe("GENERAL_DEFAULT");
    expect(updated.classificationRuleId).toBeNull();
  });
});

describe("2/3. 週規則：OVERTIME／PROJECT", () => {
  it("符合 OVERTIME 週規則時分類為超鐘點", async () => {
    const sem = await makeTestSemester(702);
    const teacher = await makeTeacher("測試分類教師乙");
    const rule = await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacher.id,
        ruleType: "OVERTIME",
        weekday: "TUE",
        periodCode: "P2",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-01"),
      },
    });
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });

    const updated = await classifySubstituteRecord(record.id);
    expect(updated.fundingSource).toBe("OVERTIME");
    expect(updated.classificationMethod).toBe("WEEKLY_RULE");
    expect(updated.classificationRuleId).toBe(rule.id);
    expect(updated.projectId).toBeNull();
  });

  it("符合 PROJECT 週規則時分類為專案，並帶出 projectId", async () => {
    const sem = await makeTestSemester(703);
    const teacher = await makeTeacher("測試分類教師丙");
    const project = await createProject({ semesterId: sem.id, name: "測試分類專案" });
    const rule = await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacher.id,
        ruleType: "PROJECT",
        projectId: project.id,
        weekday: "WED",
        periodCode: "P5",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-01"),
      },
    });
    const record = await makeRecord(sem.id, {
      originalTeacherId: teacher.id,
      date: "2026-09-02",
      weekday: "WED",
      periodCode: "P5",
    });

    const updated = await classifySubstituteRecord(record.id);
    expect(updated.fundingSource).toBe("PROJECT");
    expect(updated.classificationMethod).toBe("WEEKLY_RULE");
    expect(updated.classificationRuleId).toBe(rule.id);
    expect(updated.projectId).toBe(project.id);
  });
});

describe("4. SpecialDateRule 覆蓋 WeeklyRule", () => {
  it("單日例外優先於每週固定規則", async () => {
    const sem = await makeTestSemester(704);
    const teacher = await makeTeacher("測試分類教師丁");
    await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacher.id,
        ruleType: "OVERTIME",
        weekday: "TUE",
        periodCode: "P2",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-01"),
      },
    });
    const dateRule = await prisma.specialDateRule.create({
      data: {
        semesterId: sem.id,
        date: new Date("2026-09-15"),
        personId: teacher.id,
        periodCode: "P2",
        overrideClassification: "GENERAL",
      },
    });
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-15", weekday: "TUE" });

    const updated = await classifySubstituteRecord(record.id);
    expect(updated.fundingSource).toBe("GENERAL");
    expect(updated.classificationMethod).toBe("DATE_EXCEPTION");
    expect(updated.classificationRuleId).toBe(dateRule.id);

    // 其他星期二（沒有單日例外）仍然套用週規則
    const otherRecord = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-22", weekday: "TUE" });
    const otherUpdated = await classifySubstituteRecord(otherRecord.id);
    expect(otherUpdated.fundingSource).toBe("OVERTIME");
    expect(otherUpdated.classificationMethod).toBe("WEEKLY_RULE");
  });
});

describe("5/6. 有效日期範圍", () => {
  it("effectiveDate 尚未開始時不套用週規則，退回一般公費", async () => {
    const sem = await makeTestSemester(705);
    const teacher = await makeTeacher("測試分類教師戊");
    await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacher.id,
        ruleType: "OVERTIME",
        weekday: "TUE",
        periodCode: "P2",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-10-01"), // 生效日在代課日之後
      },
    });
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });
    const updated = await classifySubstituteRecord(record.id);
    expect(updated.fundingSource).toBe("GENERAL");
  });

  it("endDate 已結束時不套用週規則，退回一般公費", async () => {
    const sem = await makeTestSemester(706);
    const teacher = await makeTeacher("測試分類教師己");
    await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacher.id,
        ruleType: "OVERTIME",
        weekday: "TUE",
        periodCode: "P2",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-01"),
        endDate: new Date("2026-08-31"), // 已於代課日之前結束
      },
    });
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });
    const updated = await classifySubstituteRecord(record.id);
    expect(updated.fundingSource).toBe("GENERAL");
  });
});

describe("7/8. Subject Match 模式", () => {
  it("要求科目比對時，科目相符才套用規則，不符則退回一般公費", async () => {
    const sem = await makeTestSemester(707, "TEACHER_WEEKDAY_PERIOD_SUBJECT");
    const teacher = await makeTeacher("測試分類教師庚");
    await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacher.id,
        ruleType: "OVERTIME",
        weekday: "TUE",
        periodCode: "P2",
        subject: "自然",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-01"),
      },
    });

    const matching = await makeRecord(sem.id, {
      originalTeacherId: teacher.id,
      date: "2026-09-01",
      weekday: "TUE",
      subject: "自然",
    });
    const matchingResult = await classifySubstituteRecord(matching.id);
    expect(matchingResult.fundingSource).toBe("OVERTIME");

    const nonMatching = await makeRecord(sem.id, {
      originalTeacherId: teacher.id,
      date: "2026-09-08",
      weekday: "TUE",
      subject: "數學",
    });
    const nonMatchingResult = await classifySubstituteRecord(nonMatching.id);
    expect(nonMatchingResult.fundingSource).toBe("GENERAL");
  });

  it("不要求科目比對時，科目不同也套用規則", async () => {
    const sem = await makeTestSemester(708, "TEACHER_WEEKDAY_PERIOD");
    const teacher = await makeTeacher("測試分類教師辛");
    await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacher.id,
        ruleType: "OVERTIME",
        weekday: "TUE",
        periodCode: "P2",
        subject: "自然",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-01"),
      },
    });
    const record = await makeRecord(sem.id, {
      originalTeacherId: teacher.id,
      date: "2026-09-01",
      weekday: "TUE",
      subject: "數學", // 與規則設定的科目不同，但此模式不比對科目
    });
    const updated = await classifySubstituteRecord(record.id);
    expect(updated.fundingSource).toBe("OVERTIME");
  });
});

describe("9. 多規則衝突 → CONFLICT", () => {
  it("同一教師同星期同節次期間重疊的多個規則，標記 CONFLICT 並保留候選", async () => {
    const sem = await makeTestSemester(709);
    const teacher = await makeTeacher("測試分類教師壬");
    const project = await createProject({ semesterId: sem.id, name: "測試衝突專案" });
    const ruleA = await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacher.id,
        ruleType: "OVERTIME",
        weekday: "TUE",
        periodCode: "P2",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-01"),
      },
    });
    const ruleB = await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacher.id,
        ruleType: "PROJECT",
        projectId: project.id,
        weekday: "TUE",
        periodCode: "P2",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-01"),
      },
    });
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });

    const updated = await classifySubstituteRecord(record.id);
    expect(updated.fundingSource).toBe("UNDETERMINED");
    expect(updated.classificationMethod).toBe("CONFLICT");
    expect(updated.classificationRuleId).toBeNull();
    const conflict = JSON.parse(updated.conflictCandidatesJson!);
    expect(conflict.source).toBe("WEEKLY_RULE");
    expect(conflict.candidates.map((c: { ruleId: string }) => c.ruleId).sort()).toEqual([ruleA.id, ruleB.id].sort());
  });
});

describe("10. PROJECT 沒有 projectId 不形成有效 PROJECT", () => {
  it("資料異常的 PROJECT 規則（projectId 為空）被視為無效，不參與比對", async () => {
    const sem = await makeTestSemester(710);
    const teacher = await makeTeacher("測試分類教師癸");
    // 直接用 Prisma Client 建立異常資料（略過 specialWeeklyRuleService 原本會擋下這種輸入的驗證），
    // 模擬資料庫裡萬一出現 ruleType=PROJECT 卻沒有 projectId 的情況。
    await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacher.id,
        ruleType: "PROJECT",
        weekday: "TUE",
        periodCode: "P2",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-01"),
      },
    });
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });

    const updated = await classifySubstituteRecord(record.id);
    expect(updated.fundingSource).toBe("GENERAL"); // 無效規則被排除，沒有其他候選，退回一般公費
  });
});

describe("11/12/13/14. 人工覆寫", () => {
  it("可以人工覆寫為 GENERAL，並記錄原因與 ChangeLog", async () => {
    const sem = await makeTestSemester(711);
    const teacher = await makeTeacher("測試覆寫教師甲");
    await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacher.id,
        ruleType: "OVERTIME",
        weekday: "TUE",
        periodCode: "P2",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-01"),
      },
    });
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });
    await classifySubstituteRecord(record.id);

    const overridden = await overrideClassification(record.id, { fundingSource: "GENERAL" }, "測試管理者", "行政確認應為一般公費");
    expect(overridden.fundingSource).toBe("GENERAL");
    expect(overridden.classificationMethod).toBe("MANUAL_OVERRIDE");
    expect(overridden.isManuallyModified).toBe(true);
    expect(overridden.manualOverrideReason).toBe("行政確認應為一般公費");

    const logs = await prisma.changeLog.findMany({ where: { tableName: "substitute_records", recordId: record.id } });
    expect(logs.some((l) => l.reason === "行政確認應為一般公費" && l.newValue === "GENERAL")).toBe(true);
  });

  it("可以人工覆寫為 OVERTIME", async () => {
    const sem = await makeTestSemester(712);
    const teacher = await makeTeacher("測試覆寫教師乙");
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });
    await classifySubstituteRecord(record.id);

    const overridden = await overrideClassification(record.id, { fundingSource: "OVERTIME" }, "測試管理者", "確認為合理員額超鐘點");
    expect(overridden.fundingSource).toBe("OVERTIME");
    expect(overridden.projectId).toBeNull();
  });

  it("可以人工覆寫為 PROJECT，且必須指定 projectId", async () => {
    const sem = await makeTestSemester(713);
    const teacher = await makeTeacher("測試覆寫教師丙");
    const project = await createProject({ semesterId: sem.id, name: "測試覆寫專案" });
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });
    await classifySubstituteRecord(record.id);

    await expect(overrideClassification(record.id, { fundingSource: "PROJECT" }, "測試管理者", "應歸屬專案")).rejects.toThrow();

    const overridden = await overrideClassification(
      record.id,
      { fundingSource: "PROJECT", projectId: project.id },
      "測試管理者",
      "應歸屬專案"
    );
    expect(overridden.fundingSource).toBe("PROJECT");
    expect(overridden.projectId).toBe(project.id);
  });
});

describe("15. 人工覆寫不應被重新自動分類洗掉", () => {
  it("重新執行分類後，人工覆寫的結果維持不變，但 auto* 欄位會更新", async () => {
    const sem = await makeTestSemester(714);
    const teacher = await makeTeacher("測試保留教師");
    await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacher.id,
        ruleType: "OVERTIME",
        weekday: "TUE",
        periodCode: "P2",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-01"),
      },
    });
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });
    const firstAuto = await classifySubstituteRecord(record.id);
    expect(firstAuto.fundingSource).toBe("OVERTIME");

    const overridden = await overrideClassification(record.id, { fundingSource: "GENERAL" }, "測試管理者", "人工確認為一般公費");
    expect(overridden.fundingSource).toBe("GENERAL");

    // 重新執行分類（模擬管理者要求「重新自動分類」，或月批次重跑）
    const reclassified = await classifySubstituteRecord(record.id);
    expect(reclassified.fundingSource).toBe("GENERAL"); // 人工覆寫結果沒有被洗掉
    expect(reclassified.isManuallyModified).toBe(true);
    expect(reclassified.autoFundingSource).toBe("OVERTIME"); // auto 參考欄位仍然更新，讓管理者知道系統現在會怎麼判
  });

  it("可以復原成自動分類結果", async () => {
    const sem = await makeTestSemester(715);
    const teacher = await makeTeacher("測試復原教師");
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });
    await classifySubstituteRecord(record.id); // GENERAL
    await overrideClassification(record.id, { fundingSource: "OVERTIME" }, "測試管理者", "先誤改為超鐘點");

    const reverted = await revertToAutoClassification(record.id, "測試管理者", "誤改，復原");
    expect(reverted.isManuallyModified).toBe(false);
    expect(reverted.fundingSource).toBe("GENERAL");
  });
});

describe("17. BD／非BD 不互相污染", () => {
  it("分類流程完全不改變 staffType，BD 與非BD 各自獨立", async () => {
    const sem = await makeTestSemester(717);
    const teacher = await makeTeacher("測試BD教師");
    const bdRecord = await makeRecord(sem.id, {
      originalTeacherId: teacher.id,
      date: "2026-09-01",
      weekday: "TUE",
      staffType: "BD",
    });
    const nonBdRecord = await makeRecord(sem.id, {
      originalTeacherId: teacher.id,
      date: "2026-09-02",
      weekday: "WED",
      periodCode: "P3",
      staffType: "NON_BD",
    });

    await classifySubstituteRecord(bdRecord.id);
    await classifySubstituteRecord(nonBdRecord.id);

    const bdAfter = await prisma.substituteRecord.findUniqueOrThrow({ where: { id: bdRecord.id } });
    const nonBdAfter = await prisma.substituteRecord.findUniqueOrThrow({ where: { id: nonBdRecord.id } });
    expect(bdAfter.staffType).toBe("BD");
    expect(nonBdAfter.staffType).toBe("NON_BD");
  });
});

describe("18. Raw Data 不被分類流程修改", () => {
  it("SubstituteRecordRaw 在分類前後完全相同", async () => {
    const sem = await makeTestSemester(718);
    const teacher = await makeTeacher("測試Raw不變教師");
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });
    const rawBefore = await prisma.substituteRecordRaw.findUniqueOrThrow({ where: { id: record.rawRecordId! } });

    await classifySubstituteRecord(record.id);

    const rawAfter = await prisma.substituteRecordRaw.findUniqueOrThrow({ where: { id: record.rawRecordId! } });
    expect(rawAfter).toEqual(rawBefore);
  });
});

describe("批次分類與預覽查詢", () => {
  it("classifyMonthlyImport 回傳正確統計，listClassificationPreview 可依條件篩選", async () => {
    const sem = await makeTestSemester(719);
    const teacherA = await makeTeacher("批次教師A");
    const teacherB = await makeTeacher("批次教師B");
    await prisma.specialWeeklyRule.create({
      data: {
        semesterId: sem.id,
        personId: teacherA.id,
        ruleType: "OVERTIME",
        weekday: "TUE",
        periodCode: "P2",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-01"),
      },
    });

    const monthlyImport = await prisma.monthlyImport.create({
      data: { semesterId: sem.id, year: 2026, month: 9, fileName: "batch.xlsx", totalCount: 2, successCount: 2 },
    });
    const raw1 = await prisma.substituteRecordRaw.create({
      data: { monthlyImportId: monthlyImport.id, rowNumber: 1, originalTeacherText: "批次教師A", dateText: "x", periodText: "P2" },
    });
    const rec1 = await prisma.substituteRecord.create({
      data: {
        rawRecordId: raw1.id,
        monthlyImportId: monthlyImport.id,
        originalTeacherId: teacherA.id,
        date: new Date("2026-09-01"),
        weekday: "TUE",
        periodCode: "P2",
      },
    });
    const raw2 = await prisma.substituteRecordRaw.create({
      data: { monthlyImportId: monthlyImport.id, rowNumber: 2, originalTeacherText: "批次教師B", dateText: "x", periodText: "P1" },
    });
    const rec2 = await prisma.substituteRecord.create({
      data: {
        rawRecordId: raw2.id,
        monthlyImportId: monthlyImport.id,
        originalTeacherId: teacherB.id,
        date: new Date("2026-09-01"),
        weekday: "TUE",
        periodCode: "P1",
      },
    });

    const summary = await classifyMonthlyImport(monthlyImport.id, "測試管理者");
    expect(summary.total).toBe(2);
    expect(summary.overtime).toBe(1);
    expect(summary.general).toBe(1);

    const overtimeOnly = await listClassificationPreview(monthlyImport.id, { fundingSource: "OVERTIME" });
    expect(overtimeOnly.map((r) => r.id)).toEqual([rec1.id]);

    const generalOnly = await listClassificationPreview(monthlyImport.id, { fundingSource: "GENERAL" });
    expect(generalOnly.map((r) => r.id)).toEqual([rec2.id]);
  });
});

describe("SchoolCalendarDay 完全不受分類流程影響", () => {
  it("執行分類前後 SchoolCalendarDay 資料不變", async () => {
    const sem = await makeTestSemester(720);
    const day = await prisma.schoolCalendarDay.create({
      data: { semesterId: sem.id, date: new Date("2026-09-01"), weekday: "TUE", isTeachingDay: true },
    });
    const teacher = await makeTeacher("測試日曆不變教師");
    const record = await makeRecord(sem.id, { originalTeacherId: teacher.id, date: "2026-09-01", weekday: "TUE" });

    await classifySubstituteRecord(record.id);

    const dayAfter = await prisma.schoolCalendarDay.findUniqueOrThrow({ where: { id: day.id } });
    expect(dayAfter).toEqual(day);

    await prisma.schoolCalendarDay.delete({ where: { id: day.id } });
  });
});

describe("computeClassification 純函式行為", () => {
  it("原教師尚未配對時回傳 TEACHER_UNMATCHED", async () => {
    const sem = await makeTestSemester(721);
    const outcome = await computeClassification(
      { originalTeacherId: null, date: new Date("2026-09-01"), weekday: "TUE", periodCode: "P2", subject: null },
      sem
    );
    expect(outcome.fundingSource).toBe("UNDETERMINED");
    expect(outcome.classificationMethod).toBe("TEACHER_UNMATCHED");
  });
});
