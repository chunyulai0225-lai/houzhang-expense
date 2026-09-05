// Phase 4 驗收測試：費用規則管理
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { prisma } from "../src/prismaClient";
import {
  createFeeRule,
  deactivateFeeRule,
  getEffectiveFeeRule,
  getFeeRuleHistory,
  updateFeeRuleNote,
} from "../src/services/feeRuleService";

const cleanupRuleIds: string[] = [];

beforeAll(async () => {
  await prisma.$connect();
});

afterAll(async () => {
  await prisma.changeLog.deleteMany({ where: { recordId: { in: cleanupRuleIds } } });
  await prisma.feeRule.deleteMany({ where: { id: { in: cleanupRuleIds } } });
  await prisma.$disconnect();
});

describe("費率版本：新增不覆蓋舊資料", () => {
  it("同一費用類型可以有多個版本，舊版本 id/金額不變", async () => {
    const sem115_2 = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 2 } });
    const before = await getFeeRuleHistory(sem115_2.id, "SUBSTITUTE_PERIOD");
    expect(before.length).toBe(1);
    expect(before[0].amount.toNumber()).toBe(420);
    const oldRuleId = before[0].id;

    const newRule = await createFeeRule(
      { semesterId: sem115_2.id, feeType: "SUBSTITUTE_PERIOD", amount: 430, effectiveDate: new Date("2027-05-01") },
      "測試管理者"
    );
    cleanupRuleIds.push(newRule.id);

    const after = await getFeeRuleHistory(sem115_2.id, "SUBSTITUTE_PERIOD");
    expect(after.length).toBe(2);
    expect(after.map((r) => r.id)).toContain(oldRuleId);
    expect(after[0].amount.toNumber()).toBe(420); // 舊版本仍在，且排序在前（依生效日期）
    expect(after[1].amount.toNumber()).toBe(430);

    const oldRuleStillIntact = await prisma.feeRule.findUniqueOrThrow({ where: { id: oldRuleId } });
    expect(oldRuleStillIntact.amount.toNumber()).toBe(420);

    const log = await prisma.changeLog.findFirst({ where: { tableName: "fee_rules", recordId: newRule.id } });
    expect(log?.reason).toBe("新增費用規則版本");
  });

  it("結束日期早於生效日期會被拒絕", async () => {
    const sem115_1 = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    await expect(
      createFeeRule({
        semesterId: sem115_1.id,
        feeType: "OTHER",
        amount: 1,
        effectiveDate: new Date("2026-09-10"),
        endDate: new Date("2026-09-01"),
      })
    ).rejects.toThrow();
  });
});

describe("依日期查詢當時應適用的費率：舊月份不受後來調價影響", () => {
  it("115-2 的代課鐘點費在調漲前後查詢結果不同，且是各自固定的", async () => {
    const sem115_2 = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 2 } });
    // 上一個 it 已經新增 430（2027-05-01 生效），這裡再次確認查詢邏輯本身正確
    const before = await getEffectiveFeeRule({
      semesterId: sem115_2.id,
      feeType: "SUBSTITUTE_PERIOD",
      onDate: new Date("2027-03-01"),
    });
    expect(before?.amount.toNumber()).toBe(420);

    const after = await getEffectiveFeeRule({
      semesterId: sem115_2.id,
      feeType: "SUBSTITUTE_PERIOD",
      onDate: new Date("2027-06-01"),
    });
    expect(after?.amount.toNumber()).toBe(430);
    expect(after?.id).not.toBe(before?.id);
  });

  it("查詢早於任何生效日期的日期，應找不到適用費率", async () => {
    const sem115_1 = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const result = await getEffectiveFeeRule({
      semesterId: sem115_1.id,
      feeType: "SUBSTITUTE_PERIOD",
      onDate: new Date("2020-01-01"),
    });
    expect(result).toBeNull();
  });
});

describe("停用費率", () => {
  it("停用後，停用日之後查不到、之前仍查得到", async () => {
    const sem115_1 = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const rule = await createFeeRule({
      semesterId: sem115_1.id,
      feeType: "OTHER",
      amount: 50,
      effectiveDate: new Date("2026-09-01"),
    });
    cleanupRuleIds.push(rule.id);

    const stillActive = await getEffectiveFeeRule({
      semesterId: sem115_1.id,
      feeType: "OTHER",
      onDate: new Date("2026-09-05"),
    });
    expect(stillActive?.id).toBe(rule.id);

    await deactivateFeeRule(rule.id, new Date("2026-09-10"), "測試管理者", "測試停用");

    const withinRange = await getEffectiveFeeRule({
      semesterId: sem115_1.id,
      feeType: "OTHER",
      onDate: new Date("2026-09-08"),
    });
    expect(withinRange?.id).toBe(rule.id);

    const afterEnd = await getEffectiveFeeRule({
      semesterId: sem115_1.id,
      feeType: "OTHER",
      onDate: new Date("2026-09-15"),
    });
    expect(afterEnd).toBeNull();

    const log = await prisma.changeLog.findFirst({
      where: { tableName: "fee_rules", recordId: rule.id, fieldName: "endDate" },
    });
    expect(log?.reason).toBe("測試停用");
  });

  it("停用日期早於生效日期會被拒絕", async () => {
    const sem115_1 = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const rule = await createFeeRule({
      semesterId: sem115_1.id,
      feeType: "OTHER",
      amount: 99,
      effectiveDate: new Date("2026-10-01"),
    });
    cleanupRuleIds.push(rule.id);

    await expect(deactivateFeeRule(rule.id, new Date("2026-09-01"))).rejects.toThrow();
  });
});

describe("修改：只能改備註，金額與生效日期不可變更", () => {
  it("updateFeeRuleNote 只會改變 note，金額與生效日期維持不變", async () => {
    const sem115_1 = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const rule = await createFeeRule({
      semesterId: sem115_1.id,
      feeType: "OTHER",
      amount: 77,
      effectiveDate: new Date("2026-11-01"),
      note: "初始備註",
    });
    cleanupRuleIds.push(rule.id);

    const updated = await updateFeeRuleNote(rule.id, "更正後的備註", "測試管理者", "補充說明");
    expect(updated.note).toBe("更正後的備註");
    expect(updated.amount.toNumber()).toBe(77);
    expect(updated.effectiveDate.toISOString()).toBe(rule.effectiveDate.toISOString());

    const log = await prisma.changeLog.findFirst({
      where: { tableName: "fee_rules", recordId: rule.id, fieldName: "note" },
    });
    expect(log?.oldValue).toBe("初始備註");
    expect(log?.newValue).toBe("更正後的備註");
  });
});

describe("Phase 1 種子資料的費用類型維持可查詢", () => {
  it("代導師費／日薪／半日薪的初始版本仍在", async () => {
    const sem115_1 = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const homeroom = await getFeeRuleHistory(sem115_1.id, "HOMEROOM_SUBSTITUTE");
    const daily = await getFeeRuleHistory(sem115_1.id, "DAILY_WAGE");
    const halfDay = await getFeeRuleHistory(sem115_1.id, "HALF_DAY_WAGE");
    expect(homeroom[0]?.amount.toNumber()).toBe(133);
    expect(daily[0]?.amount.toNumber()).toBe(1399);
    expect(halfDay[0]?.amount.toNumber()).toBe(700);
  });
});
