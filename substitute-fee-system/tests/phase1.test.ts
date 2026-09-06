// Phase 1 驗收測試：驗證資料模型是否符合 PRD 第五節「重要資料設計原則」，
// 而不是驗證任何計算或分類邏輯（那是 Phase 8 之後的工作）。
import { PrismaClient } from "@prisma/client";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

const prisma = new PrismaClient();

beforeAll(async () => {
  await prisma.$connect();
});

afterAll(async () => {
  await prisma.$disconnect();
});

describe("學年度／學期不寫死", () => {
  it("可以建立任意未曾出現過的學年度（例如 999），不需修改程式或 schema", async () => {
    const created = await prisma.semester.create({
      data: {
        schoolYear: 999,
        term: 1,
        startDate: new Date("2999-08-31"),
        endDate: new Date("3000-01-20"),
      },
    });
    expect(created.schoolYear).toBe(999);

    // 同一學年度／學期不可重複建立
    await expect(
      prisma.semester.create({
        data: {
          schoolYear: 999,
          term: 1,
          startDate: new Date("2999-08-31"),
          endDate: new Date("3000-01-20"),
        },
      })
    ).rejects.toThrow();

    await prisma.semester.delete({ where: { id: created.id } });
  });

  it("種子資料涵蓋跨學年度的多個學期（115-1、115-2、116-1）", async () => {
    const semesters = await prisma.semester.findMany({ orderBy: [{ schoolYear: "asc" }, { term: "asc" }] });
    const keys = semesters.map((s) => `${s.schoolYear}-${s.term}`);
    expect(keys).toEqual(expect.arrayContaining(["115-1", "115-2", "116-1"]));
  });
});

describe("費用規則以資料列表示，不寫死金額", () => {
  it("同一費用類型可以有多個版本，新版本不會覆蓋舊版本", async () => {
    // 明確限定在種子資料的兩個學期，避免跟其他測試檔案（各自建立自己的
    // FeeRule）平行執行時互相干擾——這裡本來就只是驗證種子資料本身的兩筆版本。
    const rules = await prisma.feeRule.findMany({
      where: { feeType: "SUBSTITUTE_PERIOD", semester: { schoolYear: 115, term: { in: [1, 2] } } },
      orderBy: { effectiveDate: "asc" },
    });
    expect(rules.length).toBe(2);
    expect(rules[0].amount.toNumber()).toBe(405);
    expect(rules[1].amount.toNumber()).toBe(420);
    // 兩筆都還在，代表調整費用不是用 UPDATE 覆蓋，而是新增資料列
    expect(rules[0].id).not.toBe(rules[1].id);
  });
});

describe("編制內外判斷規則以資料表維護", () => {
  it("BD 分類是查表結果，不是程式常數", async () => {
    const rules = await prisma.bdClassificationRule.findMany();
    const map = Object.fromEntries(rules.map((r) => [r.codeValue, r.isBd]));
    expect(map["B"]).toBe(true);
    expect(map["D"]).toBe(true);
  });
});

describe("原始匯入資料與系統判斷後資料分離", () => {
  it("原始資料維持匯入當下的文字，不因判斷結果而改變", async () => {
    const raw = await prisma.substituteRecordRaw.findFirst({
      where: { originalTeacherText: "陳心啓" },
      include: { processedRecord: true },
    });
    expect(raw).not.toBeNull();
    expect(raw!.dateText).toBe("06-30(二) 09:10 ~ 10:00");
    expect(raw!.substituteTeacherText).toBe("徐碧苓");

    const processed = raw!.processedRecord!;
    expect(processed.fundingSource).toBe("OVERTIME");
    expect(processed.amount?.toNumber()).toBe(405);
  });
});

describe("單日例外優先於每週固定規則（資料面）", () => {
  it("同一教師同一節次，單日例外與每週規則可以並存且可分別查得", async () => {
    const chen = await prisma.person.findFirstOrThrow({ where: { name: "陳心啓" } });

    const weeklyRule = await prisma.specialWeeklyRule.findFirst({
      where: { personId: chen.id, weekday: "TUE", periodCode: "P2" },
    });
    expect(weeklyRule?.ruleType).toBe("OVERTIME");

    const dateRule = await prisma.specialDateRule.findFirst({
      where: { personId: chen.id, date: new Date("2026-06-30"), periodCode: "P2" },
    });
    expect(dateRule?.overrideClassification).toBe("OVERTIME");

    const record = await prisma.substituteRecord.findFirst({
      where: { originalTeacherId: chen.id, date: new Date("2026-06-30"), periodCode: "P2" },
    });
    expect(record?.classificationMethod).toBe("DATE_EXCEPTION");
  });
});

describe("歷史資料不因人員目前狀態而消失", () => {
  it("已離校教師的學期職務歷史紀錄仍可查詢", async () => {
    const leftPerson = await prisma.person.findFirstOrThrow({ where: { name: "王○○" } });
    expect(leftPerson.enrollmentStatus).toBe("NOT_ENROLLED");

    const role = await prisma.semesterRole.findFirst({
      where: { personId: leftPerson.id },
      include: { semester: true },
    });
    expect(role).not.toBeNull();
    expect(role!.semester.schoolYear).toBe(115);
    expect(role!.roleDetail).toBe("六年甲班導師");
  });
});

describe("每月鎖定狀態彼此獨立，歷史月份鎖定不影響其他月份", () => {
  // 注意：示範用的「已鎖定歷史月份」故意避開 2026年6月——那是全案用來驗證真實
  // 「114學年2026.06月代課(公費).xlsx」的月份，MonthlyLock 用 (year, month) 當
  // 唯一鍵、不分學期，如果種子資料把它鎖住，會連帶擋住所有真實資料回歸測試。
  it("5月已鎖定、9月待確認可以同時存在", async () => {
    const may = await prisma.monthlyLock.findUnique({ where: { year_month: { year: 2026, month: 5 } } });
    const sept = await prisma.monthlyLock.findUnique({ where: { year_month: { year: 2026, month: 9 } } });
    expect(may?.status).toBe("LOCKED");
    expect(may?.lockedBy).toBe("教學組長");
    expect(sept?.status).toBe("PENDING_REVIEW");
    expect(sept?.lockedAt).toBeNull();
  });

  it("同一年月不可重複建立鎖定紀錄", async () => {
    const existingMay = await prisma.monthlyLock.findUniqueOrThrow({ where: { year_month: { year: 2026, month: 5 } } });
    await expect(
      prisma.monthlyLock.create({
        data: { year: 2026, month: 5, status: "IMPORTED", semesterId: existingMay.semesterId },
      })
    ).rejects.toThrow();
  });
});

describe("人工修正留下可追溯的修改紀錄", () => {
  it("ChangeLog 紀錄修改前、修改後與修改原因", async () => {
    const record = await prisma.substituteRecord.findFirstOrThrow({
      where: { note: { contains: "6/30 單日例外" } },
    });
    const logs = await prisma.changeLog.findMany({
      where: { tableName: "substitute_records", recordId: record.id },
    });
    expect(logs.length).toBeGreaterThanOrEqual(1);
    expect(logs[0].oldValue).toBe("GENERAL");
    expect(logs[0].newValue).toBe("OVERTIME");
    expect(logs[0].reason).toContain("單日例外");
  });
});
