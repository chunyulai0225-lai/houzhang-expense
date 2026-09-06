// Phase 5 驗收測試：超鐘點／專案設定（規則設定與管理，不含自動分類與費用計算）
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { prisma } from "../src/prismaClient";
import { createPerson } from "../src/services/personService";
import { createProject } from "../src/services/projectService";
import {
  createSpecialDateRule,
  cancelSpecialDateRule,
  listDateRules,
  resolveClassificationForDate,
} from "../src/services/specialDateRuleService";
import {
  createSpecialWeeklyRule,
  deactivateSpecialWeeklyRule,
  deleteSpecialWeeklyRule,
  detectWeeklyRuleConflicts,
  listWeeklyRules,
  updateSpecialWeeklyRule,
} from "../src/services/specialWeeklyRuleService";

const cleanupPersonIds: string[] = [];
const cleanupProjectIds: string[] = [];
const cleanupWeeklyRuleIds: string[] = [];
const cleanupDateRuleIds: string[] = [];

beforeAll(async () => {
  await prisma.$connect();
});

afterAll(async () => {
  await prisma.changeLog.deleteMany({
    where: {
      OR: [
        { recordId: { in: cleanupPersonIds } },
        { recordId: { in: cleanupProjectIds } },
        { recordId: { in: cleanupWeeklyRuleIds } },
        { recordId: { in: cleanupDateRuleIds } },
      ],
    },
  });
  await prisma.specialDateRule.deleteMany({ where: { id: { in: cleanupDateRuleIds } } });
  await prisma.specialWeeklyRule.deleteMany({ where: { id: { in: cleanupWeeklyRuleIds } } });
  await prisma.project.deleteMany({ where: { id: { in: cleanupProjectIds } } });
  await prisma.person.deleteMany({ where: { id: { in: cleanupPersonIds } } });
  await prisma.$disconnect();
});

async function makeTestPerson(name: string) {
  const result = await createPerson({ name });
  if (result.status !== "CREATED") throw new Error("unreachable");
  cleanupPersonIds.push(result.person.id);
  return result.person;
}

describe("1. 新增超鐘點規則", () => {
  it("可以新增 ruleType=OVERTIME 的每週固定規則", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const person = await makeTestPerson("測試超鐘點教師甲");

    const { rule, conflicts } = await createSpecialWeeklyRule(
      {
        semesterId: sem.id,
        personId: person.id,
        ruleType: "OVERTIME",
        weekday: "WED",
        periodCode: "P3",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-31"),
      },
      "測試管理者"
    );
    cleanupWeeklyRuleIds.push(rule.id);

    expect(rule.ruleType).toBe("OVERTIME");
    expect(rule.projectId).toBeNull();
    expect(conflicts).toEqual([]);
  });

  it("ruleType=OVERTIME 卻指定 projectId 會被拒絕", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const person = await makeTestPerson("測試超鐘點教師乙");
    const project = await createProject({ semesterId: sem.id, name: "測試不應使用的專案" });
    cleanupProjectIds.push(project.id);

    await expect(
      createSpecialWeeklyRule({
        semesterId: sem.id,
        personId: person.id,
        ruleType: "OVERTIME",
        projectId: project.id,
        weekday: "THU",
        periodCode: "P1",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-31"),
      })
    ).rejects.toThrow();
  });
});

describe("2. 新增專案", () => {
  it("可以新增專案，名稱不寫死，預設啟用", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const project = await createProject({ semesterId: sem.id, name: "測試專案：課後陪讀計畫" }, "測試管理者");
    cleanupProjectIds.push(project.id);

    expect(project.name).toBe("測試專案：課後陪讀計畫");
    expect(project.isActive).toBe(true);

    const log = await prisma.changeLog.findFirst({ where: { tableName: "projects", recordId: project.id } });
    expect(log?.reason).toBe("新增專案");
  });
});

describe("3. 新增專案減課規則", () => {
  it("可以表達「某教師、某星期、每週N節、某專案」", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const project = await createProject({ semesterId: sem.id, name: "測試專案：輔導團減課" });
    cleanupProjectIds.push(project.id);
    const person = await makeTestPerson("測試專案教師甲");

    const { rule } = await createSpecialWeeklyRule({
      semesterId: sem.id,
      personId: person.id,
      ruleType: "PROJECT",
      projectId: project.id,
      weekday: "WED",
      periodCode: "P5",
      weeklyPeriods: 1,
      effectiveDate: new Date("2026-08-31"),
    });
    cleanupWeeklyRuleIds.push(rule.id);

    expect(rule.ruleType).toBe("PROJECT");
    expect(rule.projectId).toBe(project.id);
  });

  it("ruleType=PROJECT 沒有指定 projectId 會被拒絕", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const person = await makeTestPerson("測試專案教師乙");
    await expect(
      createSpecialWeeklyRule({
        semesterId: sem.id,
        personId: person.id,
        ruleType: "PROJECT",
        weekday: "WED",
        periodCode: "P5",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-31"),
      })
    ).rejects.toThrow();
  });
});

describe("4/5. 生效日期／結束日期正確", () => {
  it("規則的生效日期正確保存，asOfDate 篩選能反映生效狀態", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const person = await makeTestPerson("測試日期教師甲");

    const { rule } = await createSpecialWeeklyRule({
      semesterId: sem.id,
      personId: person.id,
      ruleType: "OVERTIME",
      weekday: "FRI",
      periodCode: "P4",
      weeklyPeriods: 1,
      effectiveDate: new Date("2026-10-01"),
      endDate: new Date("2026-12-31"),
    });
    cleanupWeeklyRuleIds.push(rule.id);

    expect(rule.effectiveDate.toISOString().slice(0, 10)).toBe("2026-10-01");
    expect(rule.endDate?.toISOString().slice(0, 10)).toBe("2026-12-31");

    const beforeEffective = await listWeeklyRules(sem.id, { personId: person.id, asOfDate: new Date("2026-09-01") });
    expect(beforeEffective.some((r) => r.id === rule.id)).toBe(false);

    const withinRange = await listWeeklyRules(sem.id, { personId: person.id, asOfDate: new Date("2026-11-01") });
    expect(withinRange.some((r) => r.id === rule.id)).toBe(true);

    const afterEnd = await listWeeklyRules(sem.id, { personId: person.id, asOfDate: new Date("2027-01-01") });
    expect(afterEnd.some((r) => r.id === rule.id)).toBe(false);
  });

  it("結束日期早於生效日期會被拒絕", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const person = await makeTestPerson("測試日期教師乙");
    await expect(
      createSpecialWeeklyRule({
        semesterId: sem.id,
        personId: person.id,
        ruleType: "OVERTIME",
        weekday: "FRI",
        periodCode: "P4",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-12-31"),
        endDate: new Date("2026-10-01"),
      })
    ).rejects.toThrow();
  });
});

describe("6. 同一教師不同期間可以有不同規則", () => {
  it("不重疊的期間允許共存，且不視為衝突", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const person = await makeTestPerson("測試多期間教師");
    const projectA = await createProject({ semesterId: sem.id, name: "測試專案A" });
    const projectB = await createProject({ semesterId: sem.id, name: "測試專案B" });
    cleanupProjectIds.push(projectA.id, projectB.id);

    const ruleA = await createSpecialWeeklyRule({
      semesterId: sem.id,
      personId: person.id,
      ruleType: "PROJECT",
      projectId: projectA.id,
      weekday: "WED",
      periodCode: "P5",
      weeklyPeriods: 1,
      effectiveDate: new Date("2026-08-31"),
      endDate: new Date("2026-10-31"),
    });
    const ruleB = await createSpecialWeeklyRule({
      semesterId: sem.id,
      personId: person.id,
      ruleType: "PROJECT",
      projectId: projectB.id,
      weekday: "WED",
      periodCode: "P5",
      weeklyPeriods: 1,
      effectiveDate: new Date("2026-11-01"),
      endDate: new Date("2027-01-20"),
    });
    cleanupWeeklyRuleIds.push(ruleA.rule.id, ruleB.rule.id);

    expect(ruleA.conflicts).toEqual([]);
    expect(ruleB.conflicts).toEqual([]);

    const conflicts = await detectWeeklyRuleConflicts(sem.id, person.id);
    expect(conflicts).toEqual([]);
  });
});

describe("7. 規則衝突會被偵測", () => {
  it("同一教師同一星期同一節次，重疊期間的兩筆規則會被標示為衝突", async () => {
    const chen = await prisma.person.findFirstOrThrow({ where: { name: "陳心啓" } });
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const project = await createProject({ semesterId: sem.id, name: "測試專案：與超鐘點衝突" });
    cleanupProjectIds.push(project.id);

    // 陳心啓在種子資料裡已經有「星期二第2節、超鐘點、2026-08-31~2027-01-20」的規則，
    // 這裡刻意在重疊期間內新增另一筆專案規則，應該被偵測為衝突。
    const { rule, conflicts } = await createSpecialWeeklyRule({
      semesterId: sem.id,
      personId: chen.id,
      ruleType: "PROJECT",
      projectId: project.id,
      weekday: "TUE",
      periodCode: "P2",
      weeklyPeriods: 1,
      effectiveDate: new Date("2026-09-01"),
      endDate: new Date("2026-12-01"),
    });
    cleanupWeeklyRuleIds.push(rule.id);

    expect(conflicts.length).toBeGreaterThanOrEqual(1);
    expect(conflicts.some((c) => c.ruleIds.includes(rule.id))).toBe(true);

    const allConflicts = await detectWeeklyRuleConflicts(sem.id, chen.id);
    expect(allConflicts.some((c) => c.ruleIds.includes(rule.id))).toBe(true);
  });
});

describe("8/9. 單日例外：建立與優先權", () => {
  it("可以建立單日例外", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const person = await makeTestPerson("測試單日例外教師甲");

    const rule = await createSpecialDateRule(
      {
        semesterId: sem.id,
        date: new Date("2026-09-15"),
        personId: person.id,
        periodCode: "P2",
        overrideClassification: "OVERTIME",
        note: "臨時安排",
      },
      "測試管理者"
    );
    cleanupDateRuleIds.push(rule.id);

    expect(rule.overrideClassification).toBe("OVERTIME");
    const log = await prisma.changeLog.findFirst({ where: { tableName: "special_date_rules", recordId: rule.id } });
    expect(log?.reason).toBe("新增單日例外");
  });

  it("單日例外優先於每週固定規則", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const person = await makeTestPerson("測試優先權教師");

    const { rule: weeklyRule } = await createSpecialWeeklyRule({
      semesterId: sem.id,
      personId: person.id,
      ruleType: "OVERTIME",
      weekday: "THU",
      periodCode: "P3",
      weeklyPeriods: 1,
      effectiveDate: new Date("2026-08-31"),
      endDate: new Date("2027-01-20"),
    });
    cleanupWeeklyRuleIds.push(weeklyRule.id);

    // 星期四對應的其中一天：2026/09/17，先驗證沒有單日例外時，回傳每週固定規則
    const withoutException = await resolveClassificationForDate({
      semesterId: sem.id,
      personId: person.id,
      date: new Date("2026-09-17"),
      weekday: "THU",
      periodCode: "P3",
    });
    expect(withoutException.source).toBe("WEEKLY_RULE");
    expect(withoutException.weeklyRule?.ruleType).toBe("OVERTIME");

    // 針對同一天新增單日例外，改判為一般公費
    const dateRule = await createSpecialDateRule({
      semesterId: sem.id,
      date: new Date("2026-09-17"),
      personId: person.id,
      periodCode: "P3",
      originalClassificationNote: "OVERTIME",
      overrideClassification: "GENERAL",
      note: "本日特殊安排，改回一般公費",
    });
    cleanupDateRuleIds.push(dateRule.id);

    const withException = await resolveClassificationForDate({
      semesterId: sem.id,
      personId: person.id,
      date: new Date("2026-09-17"),
      weekday: "THU",
      periodCode: "P3",
    });
    expect(withException.source).toBe("DATE_EXCEPTION");
    expect(withException.overrideClassification).toBe("GENERAL");

    // 其他星期四（沒有設定單日例外）仍然回到每週固定規則
    const otherThursday = await resolveClassificationForDate({
      semesterId: sem.id,
      personId: person.id,
      date: new Date("2026-09-24"),
      weekday: "THU",
      periodCode: "P3",
    });
    expect(otherThursday.source).toBe("WEEKLY_RULE");
  });

  it("取消的單日例外不會影響優先權判斷", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const person = await makeTestPerson("測試取消例外教師");

    const dateRule = await createSpecialDateRule({
      semesterId: sem.id,
      date: new Date("2026-10-06"),
      personId: person.id,
      periodCode: "P1",
      overrideClassification: "GENERAL",
    });
    cleanupDateRuleIds.push(dateRule.id);

    await cancelSpecialDateRule(dateRule.id, "測試管理者", "誤植，取消");

    const resolved = await resolveClassificationForDate({
      semesterId: sem.id,
      personId: person.id,
      date: new Date("2026-10-06"),
      weekday: "TUE",
      periodCode: "P1",
    });
    expect(resolved.source).toBe("GENERAL_DEFAULT");

    const cancelled = await prisma.specialDateRule.findUniqueOrThrow({ where: { id: dateRule.id } });
    expect(cancelled.isCancelled).toBe(true);

    const activeList = await listDateRules(sem.id, { personId: person.id });
    expect(activeList.some((r) => r.id === dateRule.id)).toBe(false);
    const allList = await listDateRules(sem.id, { personId: person.id, includeCancelled: true });
    expect(allList.some((r) => r.id === dateRule.id)).toBe(true);
  });
});

describe("10. 停用規則不刪除歷史", () => {
  it("停用每週固定規則後，資料列仍存在，只是不再視為生效中", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const person = await makeTestPerson("測試停用教師");

    const { rule } = await createSpecialWeeklyRule({
      semesterId: sem.id,
      personId: person.id,
      ruleType: "OVERTIME",
      weekday: "MON",
      periodCode: "P1",
      weeklyPeriods: 1,
      effectiveDate: new Date("2026-08-31"),
    });
    cleanupWeeklyRuleIds.push(rule.id);

    await deactivateSpecialWeeklyRule(rule.id, new Date("2026-10-31"), "測試管理者", "測試停用");

    const stillExists = await prisma.specialWeeklyRule.findUnique({ where: { id: rule.id } });
    expect(stillExists).not.toBeNull();
    expect(stillExists?.endDate?.toISOString().slice(0, 10)).toBe("2026-10-31");

    const stillEffectiveBeforeEnd = await listWeeklyRules(sem.id, {
      personId: person.id,
      asOfDate: new Date("2026-10-01"),
    });
    expect(stillEffectiveBeforeEnd.some((r) => r.id === rule.id)).toBe(true);

    const notEffectiveAfterEnd = await listWeeklyRules(sem.id, {
      personId: person.id,
      asOfDate: new Date("2026-11-01"),
    });
    expect(notEffectiveAfterEnd.some((r) => r.id === rule.id)).toBe(false);
  });
});

describe("11. ChangeLog 正常", () => {
  it("新增、修改、刪除每週固定規則都會留下 ChangeLog", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const person = await makeTestPerson("測試ChangeLog教師");

    const { rule } = await createSpecialWeeklyRule(
      {
        semesterId: sem.id,
        personId: person.id,
        ruleType: "OVERTIME",
        weekday: "MON",
        periodCode: "P2",
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-08-31"),
      },
      "測試管理者"
    );

    await updateSpecialWeeklyRule(rule.id, { weeklyPeriods: 2 }, "測試管理者", "更正每週節數");
    await deleteSpecialWeeklyRule(rule.id, "測試管理者", "測試刪除");

    const logs = await prisma.changeLog.findMany({
      where: { tableName: "special_weekly_rules", recordId: rule.id },
      orderBy: { createdAt: "asc" },
    });
    expect(logs.map((l) => l.reason)).toEqual(["新增每週固定規則", "更正每週節數", "測試刪除"]);
    expect(logs[1].fieldName).toBe("weeklyPeriods");
    expect(logs[1].oldValue).toBe("1");
    expect(logs[1].newValue).toBe("2");
  });
});

describe("12. subject 可以保存但不是必要判斷條件", () => {
  it("有無 subject 都可以正常建立與查詢", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const personWithSubject = await makeTestPerson("測試科目教師甲");
    const personWithoutSubject = await makeTestPerson("測試科目教師乙");

    const withSubject = await createSpecialWeeklyRule({
      semesterId: sem.id,
      personId: personWithSubject.id,
      ruleType: "OVERTIME",
      weekday: "FRI",
      periodCode: "P6",
      subject: "自然",
      weeklyPeriods: 1,
      effectiveDate: new Date("2026-08-31"),
    });
    const withoutSubject = await createSpecialWeeklyRule({
      semesterId: sem.id,
      personId: personWithoutSubject.id,
      ruleType: "OVERTIME",
      weekday: "FRI",
      periodCode: "P6",
      weeklyPeriods: 1,
      effectiveDate: new Date("2026-08-31"),
    });
    cleanupWeeklyRuleIds.push(withSubject.rule.id, withoutSubject.rule.id);

    expect(withSubject.rule.subject).toBe("自然");
    expect(withoutSubject.rule.subject).toBeNull();

    // 查詢預設不會因為 subject 不同而被排除或要求一定要相符
    const listed = await listWeeklyRules(sem.id, { weekday: "FRI", periodCode: "P6" });
    expect(listed.some((r) => r.id === withSubject.rule.id)).toBe(true);
    expect(listed.some((r) => r.id === withoutSubject.rule.id)).toBe(true);
  });
});

describe("13. overtimeMatchMode 可供後續計算模組使用", () => {
  it("Semester.overtimeMatchMode 可以被讀取與切換", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 116, term: 1 } });
    expect(["TEACHER_WEEKDAY_PERIOD", "TEACHER_WEEKDAY_PERIOD_SUBJECT"]).toContain(sem.overtimeMatchMode);

    const original = sem.overtimeMatchMode;
    const toggled = original === "TEACHER_WEEKDAY_PERIOD" ? "TEACHER_WEEKDAY_PERIOD_SUBJECT" : "TEACHER_WEEKDAY_PERIOD";

    const updated = await prisma.semester.update({ where: { id: sem.id }, data: { overtimeMatchMode: toggled } });
    expect(updated.overtimeMatchMode).toBe(toggled);

    // 還原，避免影響其他測試或種子資料的預期狀態
    await prisma.semester.update({ where: { id: sem.id }, data: { overtimeMatchMode: original } });
  });
});
