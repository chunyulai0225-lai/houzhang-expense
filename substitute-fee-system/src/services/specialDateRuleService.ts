// Phase 5：單日例外管理
//
// 只處理「規則本身」的新增／取消／查詢，以及「規則優先權」的查詢輔助函式
// （單日例外 > 每週固定規則 > 一般公費，資料層級的規則解析）。
// 不處理實際代課紀錄的自動分類、金額計算——那是 Phase 8 讀取每月匯入資料
// 後才會做的事，這裡只驗證「規則本身」的優先權設計是否正確。

import type { OverrideClassification, SpecialDateRule, SpecialWeeklyRule, Weekday } from "@prisma/client";
import { prisma } from "../prismaClient";

export interface CreateSpecialDateRuleInput {
  semesterId: string;
  date: Date;
  personId: string;
  periodCode: string;
  originalClassificationNote?: string;
  overrideClassification: OverrideClassification;
  projectId?: string;
  note?: string;
}

function validateClassificationProject(
  overrideClassification: OverrideClassification,
  projectId: string | null | undefined
) {
  if (overrideClassification === "PROJECT" && !projectId) {
    throw new Error("overrideClassification 為 PROJECT 時必須指定 projectId");
  }
  if (overrideClassification !== "PROJECT" && projectId) {
    throw new Error("只有 overrideClassification 為 PROJECT 時才能指定 projectId");
  }
}

// 1. 新增單日例外
export async function createSpecialDateRule(
  input: CreateSpecialDateRuleInput,
  changedBy?: string
): Promise<SpecialDateRule> {
  validateClassificationProject(input.overrideClassification, input.projectId);

  const rule = await prisma.specialDateRule.create({
    data: {
      semesterId: input.semesterId,
      date: input.date,
      personId: input.personId,
      periodCode: input.periodCode,
      originalClassificationNote: input.originalClassificationNote,
      overrideClassification: input.overrideClassification,
      projectId: input.projectId,
      note: input.note,
    },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "special_date_rules",
      recordId: rule.id,
      newValue: describeDateRule(rule),
      changedBy,
      reason: "新增單日例外",
    },
  });

  return rule;
}

// 2. 取消單日例外（生效／取消狀態）。不刪除資料列，只標記取消，歷史仍可查詢。
export async function cancelSpecialDateRule(
  ruleId: string,
  changedBy?: string,
  reason?: string
): Promise<SpecialDateRule> {
  const existing = await prisma.specialDateRule.findUniqueOrThrow({ where: { id: ruleId } });
  if (existing.isCancelled) {
    throw new Error("此單日例外已經是取消狀態");
  }

  const updated = await prisma.specialDateRule.update({
    where: { id: ruleId },
    data: { isCancelled: true, cancelledAt: new Date(), cancelledBy: changedBy },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "special_date_rules",
      recordId: ruleId,
      fieldName: "isCancelled",
      oldValue: "false",
      newValue: "true",
      changedBy,
      reason: reason ?? "取消單日例外",
    },
  });

  return updated;
}

export interface ListDateRulesFilter {
  personId?: string;
  date?: Date;
  includeCancelled?: boolean;
}

// 3. 查詢：某學期的單日例外列表
export async function listDateRules(semesterId: string, filter: ListDateRulesFilter = {}) {
  return prisma.specialDateRule.findMany({
    where: {
      semesterId,
      ...(filter.personId ? { personId: filter.personId } : {}),
      ...(filter.date ? { date: filter.date } : {}),
      ...(filter.includeCancelled ? {} : { isCancelled: false }),
    },
    include: { person: true, project: true },
    orderBy: [{ date: "asc" }],
  });
}

export type ClassificationSource = "DATE_EXCEPTION" | "WEEKLY_RULE" | "GENERAL_DEFAULT";

export interface ResolvedClassification {
  source: ClassificationSource;
  overrideClassification?: OverrideClassification; // 當 source 為 DATE_EXCEPTION
  dateRule?: SpecialDateRule;
  weeklyRule?: SpecialWeeklyRule;
}

// 4. 規則優先權查詢：單日例外 > 每週固定規則 > 一般公費（無特殊設定）。
//    這只回答「規則怎麼設定」，不回答「這一節該付多少錢」——金額計算留給 Phase 8。
export async function resolveClassificationForDate(params: {
  semesterId: string;
  personId: string;
  date: Date;
  weekday: Weekday;
  periodCode: string;
}): Promise<ResolvedClassification> {
  const dateRule = await prisma.specialDateRule.findFirst({
    where: {
      semesterId: params.semesterId,
      personId: params.personId,
      date: params.date,
      periodCode: params.periodCode,
      isCancelled: false,
    },
  });
  if (dateRule) {
    return { source: "DATE_EXCEPTION", overrideClassification: dateRule.overrideClassification, dateRule };
  }

  const weeklyRule = await prisma.specialWeeklyRule.findFirst({
    where: {
      semesterId: params.semesterId,
      personId: params.personId,
      weekday: params.weekday,
      periodCode: params.periodCode,
      effectiveDate: { lte: params.date },
      OR: [{ endDate: null }, { endDate: { gte: params.date } }],
    },
    orderBy: { effectiveDate: "desc" },
  });
  if (weeklyRule) {
    return { source: "WEEKLY_RULE", weeklyRule };
  }

  return { source: "GENERAL_DEFAULT" };
}

function describeDateRule(rule: Pick<SpecialDateRule, "date" | "periodCode" | "overrideClassification">): string {
  return `${rule.date.toISOString().slice(0, 10)} ${rule.periodCode} -> ${rule.overrideClassification}`;
}
