// Phase 4：費用規則管理
//
// 資料模型沿用 Phase 1 的 FeeRule（semesterId + feeType + amount + effectiveDate + endDate），
// 沒有新增欄位或資料表。這裡只做「費率資料管理」：新增版本／修改備註／停用，
// 以及「依日期查出當時應該適用哪一筆費率」的查詢。
//
// 刻意不做的事（尚未確認，留給後續 Phase）：
// - 代導師費／日薪／半日薪何時產生、彼此是否互斥
// - 專案是否使用不同費率、BD/非BD是否影響費率
// 這些都是「計算邏輯」，不是「費率資料管理」，Phase 4 不處理。
//
// 關鍵設計：金額（amount）、費用類型（feeType）、所屬學期（semesterId）在建立後
// 就不能再被任何 API 修改——要調整金額只能新增一筆新版本（新的 effectiveDate）。
// 這樣可以保證舊月份的計算結果不會因為之後調價而改變：只要計算時是用
// getEffectiveFeeRule() 依「當時的日期」去查，查到的永遠是當時生效的那一筆。

import type { FeeRule, FeeType } from "@prisma/client";
import { prisma } from "../prismaClient";

export interface CreateFeeRuleInput {
  semesterId: string;
  feeType: FeeType;
  amount: number | string;
  effectiveDate: Date;
  endDate?: Date;
  note?: string;
}

// 1. 新增費率版本。同一 feeType 可以在同一學期內有多筆，依 effectiveDate 區分版本，
//    絕對不會覆蓋既有資料列。
export async function createFeeRule(input: CreateFeeRuleInput, changedBy?: string): Promise<FeeRule> {
  if (input.endDate && input.endDate < input.effectiveDate) {
    throw new Error("結束日期不能早於生效日期");
  }

  const rule = await prisma.feeRule.create({
    data: {
      semesterId: input.semesterId,
      feeType: input.feeType,
      amount: input.amount,
      effectiveDate: input.effectiveDate,
      endDate: input.endDate,
      note: input.note,
    },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "fee_rules",
      recordId: rule.id,
      newValue: describeFeeRule(rule),
      changedBy,
      reason: "新增費用規則版本",
    },
  });

  return rule;
}

// 2. 修改：只允許修改備註，金額／費用類型／生效日期一律不可變更，
//    避免有 API 可以「就地」改掉已經存在的費率金額。
export async function updateFeeRuleNote(
  ruleId: string,
  note: string | null,
  changedBy?: string,
  reason?: string
): Promise<FeeRule> {
  const existing = await prisma.feeRule.findUniqueOrThrow({ where: { id: ruleId } });
  const updated = await prisma.feeRule.update({ where: { id: ruleId }, data: { note } });

  if (existing.note !== updated.note) {
    await prisma.changeLog.create({
      data: {
        tableName: "fee_rules",
        recordId: ruleId,
        fieldName: "note",
        oldValue: existing.note,
        newValue: updated.note,
        changedBy,
        reason,
      },
    });
  }

  return updated;
}

// 3. 停用：設定 endDate，代表此版本自某日起不再適用。不刪除資料列，歷史仍可查詢。
export async function deactivateFeeRule(
  ruleId: string,
  endDate: Date,
  changedBy?: string,
  reason?: string
): Promise<FeeRule> {
  const existing = await prisma.feeRule.findUniqueOrThrow({ where: { id: ruleId } });
  if (endDate < existing.effectiveDate) {
    throw new Error("停用日期不能早於生效日期");
  }

  const updated = await prisma.feeRule.update({ where: { id: ruleId }, data: { endDate } });

  await prisma.changeLog.create({
    data: {
      tableName: "fee_rules",
      recordId: ruleId,
      fieldName: "endDate",
      oldValue: existing.endDate ? existing.endDate.toISOString() : null,
      newValue: endDate.toISOString(),
      changedBy,
      reason: reason ?? "停用費用規則",
    },
  });

  return updated;
}

// 4. 查詢：某學期、某費用類型的完整歷史版本（舊到新）
export async function getFeeRuleHistory(semesterId: string, feeType: FeeType): Promise<FeeRule[]> {
  return prisma.feeRule.findMany({
    where: { semesterId, feeType },
    orderBy: { effectiveDate: "asc" },
  });
}

// 5. 查詢：在指定日期，某學期、某費用類型「當時」應該適用的那一筆費率。
//    只取 effectiveDate <= onDate 中最新的一筆，且尚未停用（endDate 為空或 >= onDate）。
//    這是確保「舊月份計算不受後來調價影響」的核心：計算舊資料時一律用當時的日期查詢，
//    永遠查不到日後才生效的新費率。
export async function getEffectiveFeeRule(params: {
  semesterId: string;
  feeType: FeeType;
  onDate: Date;
}): Promise<FeeRule | null> {
  return prisma.feeRule.findFirst({
    where: {
      semesterId: params.semesterId,
      feeType: params.feeType,
      effectiveDate: { lte: params.onDate },
      OR: [{ endDate: null }, { endDate: { gte: params.onDate } }],
    },
    orderBy: { effectiveDate: "desc" },
  });
}

function describeFeeRule(rule: Pick<FeeRule, "feeType" | "amount" | "effectiveDate">): string {
  return `${rule.feeType}=${rule.amount.toString()}(自${rule.effectiveDate.toISOString().slice(0, 10)})`;
}
