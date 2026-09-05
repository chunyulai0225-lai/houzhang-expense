// Phase 5：每週固定超鐘點／專案規則管理
//
// 超鐘點與專案共用 SpecialWeeklyRule（Phase 1 已建立），以 ruleType 區分。
// 這裡只做規則的新增／修改／停用／刪除／查詢，以及衝突偵測。
// 完全不處理：當月應有節數、上班日計算、代課節數扣除、費用換算——
// 那些是 Phase 8 之後的計算邏輯，不在這裡決定。

import type { Prisma, RuleType, SpecialWeeklyRule, Weekday } from "@prisma/client";
import { prisma } from "../prismaClient";

export interface CreateSpecialWeeklyRuleInput {
  semesterId: string;
  personId: string;
  ruleType: RuleType;
  projectId?: string;
  weekday: Weekday;
  periodCode: string;
  subject?: string; // 保留欄位；是否納入判斷交由 Semester.overtimeMatchMode 決定，這裡不強制
  weeklyPeriods: number | string;
  effectiveDate: Date;
  endDate?: Date;
  note?: string;
}

function validateRuleTypeProject(ruleType: RuleType, projectId: string | null | undefined) {
  if (ruleType === "PROJECT" && !projectId) {
    throw new Error("ruleType 為 PROJECT 時必須指定 projectId");
  }
  if (ruleType === "OVERTIME" && projectId) {
    throw new Error("ruleType 為 OVERTIME 時不應該指定 projectId");
  }
}

function validateDateRange(effectiveDate: Date, endDate: Date | null | undefined) {
  if (endDate && endDate < effectiveDate) {
    throw new Error("結束日期不能早於生效日期");
  }
}

export interface WeeklyRuleConflict {
  ruleIds: [string, string];
  personId: string;
  weekday: Weekday;
  periodCode: string;
}

function rangesOverlap(aStart: Date, aEnd: Date | null, bStart: Date, bEnd: Date | null): boolean {
  const aEndTime = aEnd ? aEnd.getTime() : Number.POSITIVE_INFINITY;
  const bEndTime = bEnd ? bEnd.getTime() : Number.POSITIVE_INFINITY;
  return aStart.getTime() <= bEndTime && bStart.getTime() <= aEndTime;
}

// 衝突偵測：同一教師＋同一星期＋同一節次，只要有兩筆（或以上）規則的有效期間互相重疊，
// 一律視為衝突——不論 ruleType 是否相同、是否指向同一個專案，都不自動判斷該用哪一筆，
// 交由管理者處理（調整期間或刪除其中一筆）。
export async function detectWeeklyRuleConflicts(
  semesterId: string,
  personId?: string
): Promise<WeeklyRuleConflict[]> {
  const rules = await prisma.specialWeeklyRule.findMany({
    where: { semesterId, ...(personId ? { personId } : {}) },
  });

  const groups = new Map<string, SpecialWeeklyRule[]>();
  for (const rule of rules) {
    const key = `${rule.personId}|${rule.weekday}|${rule.periodCode}`;
    const list = groups.get(key);
    if (list) {
      list.push(rule);
    } else {
      groups.set(key, [rule]);
    }
  }

  const conflicts: WeeklyRuleConflict[] = [];
  for (const groupRules of groups.values()) {
    if (groupRules.length < 2) continue;
    for (let i = 0; i < groupRules.length; i++) {
      for (let j = i + 1; j < groupRules.length; j++) {
        const a = groupRules[i];
        const b = groupRules[j];
        if (rangesOverlap(a.effectiveDate, a.endDate, b.effectiveDate, b.endDate)) {
          conflicts.push({
            ruleIds: [a.id, b.id],
            personId: a.personId,
            weekday: a.weekday,
            periodCode: a.periodCode,
          });
        }
      }
    }
  }
  return conflicts;
}

export interface WeeklyRuleWriteResult {
  rule: SpecialWeeklyRule;
  conflicts: WeeklyRuleConflict[];
}

// 1/3. 新增每週固定規則（超鐘點或專案）。系統不會拒絕建立衝突規則——
//     而是建立後立刻回傳偵測到的衝突清單，由呼叫端（管理介面）明確標示
//     「⚠️ 規則衝突」要求管理者處理，不會默默接受、隨便選一個。
export async function createSpecialWeeklyRule(
  input: CreateSpecialWeeklyRuleInput,
  changedBy?: string
): Promise<WeeklyRuleWriteResult> {
  validateRuleTypeProject(input.ruleType, input.projectId);
  validateDateRange(input.effectiveDate, input.endDate);

  const rule = await prisma.specialWeeklyRule.create({
    data: {
      semesterId: input.semesterId,
      personId: input.personId,
      ruleType: input.ruleType,
      projectId: input.projectId,
      weekday: input.weekday,
      periodCode: input.periodCode,
      subject: input.subject,
      weeklyPeriods: input.weeklyPeriods,
      effectiveDate: input.effectiveDate,
      endDate: input.endDate,
      note: input.note,
    },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "special_weekly_rules",
      recordId: rule.id,
      newValue: describeWeeklyRule(rule),
      changedBy,
      reason: "新增每週固定規則",
    },
  });

  const allConflicts = await detectWeeklyRuleConflicts(input.semesterId, input.personId);
  const conflicts = allConflicts.filter((c) => c.ruleIds.includes(rule.id));

  return { rule, conflicts };
}

export interface UpdateSpecialWeeklyRuleInput {
  ruleType?: RuleType;
  projectId?: string | null;
  weekday?: Weekday;
  periodCode?: string;
  subject?: string | null;
  weeklyPeriods?: number | string;
  effectiveDate?: Date;
  endDate?: Date | null;
  note?: string | null;
}

// 2. 修改每週固定規則。所有變更逐欄位記錄，並重新檢查衝突。
export async function updateSpecialWeeklyRule(
  ruleId: string,
  changes: UpdateSpecialWeeklyRuleInput,
  changedBy?: string,
  reason?: string
): Promise<WeeklyRuleWriteResult> {
  const existing = await prisma.specialWeeklyRule.findUniqueOrThrow({ where: { id: ruleId } });

  const nextRuleType = changes.ruleType ?? existing.ruleType;
  const nextProjectId = "projectId" in changes ? changes.projectId : existing.projectId;
  validateRuleTypeProject(nextRuleType, nextProjectId);

  const nextEffectiveDate = changes.effectiveDate ?? existing.effectiveDate;
  const nextEndDate = "endDate" in changes ? changes.endDate ?? null : existing.endDate;
  validateDateRange(nextEffectiveDate, nextEndDate);

  const updated = await prisma.specialWeeklyRule.update({
    where: { id: ruleId },
    data: changes as Prisma.SpecialWeeklyRuleUpdateInput,
  });

  const fields = Object.keys(changes) as (keyof UpdateSpecialWeeklyRuleInput)[];
  for (const field of fields) {
    const oldValue = existing[field as keyof SpecialWeeklyRule];
    const newValue = updated[field as keyof SpecialWeeklyRule];
    if (stringifyValue(oldValue) !== stringifyValue(newValue)) {
      await prisma.changeLog.create({
        data: {
          tableName: "special_weekly_rules",
          recordId: ruleId,
          fieldName: field,
          oldValue: stringifyValue(oldValue),
          newValue: stringifyValue(newValue),
          changedBy,
          reason,
        },
      });
    }
  }

  const allConflicts = await detectWeeklyRuleConflicts(updated.semesterId, updated.personId);
  const conflicts = allConflicts.filter((c) => c.ruleIds.includes(ruleId));

  return { rule: updated, conflicts };
}

// 3. 停用（設定結束日期）——優先於刪除，歷史規則仍可查詢。
export async function deactivateSpecialWeeklyRule(
  ruleId: string,
  endDate: Date,
  changedBy?: string,
  reason?: string
): Promise<SpecialWeeklyRule> {
  const existing = await prisma.specialWeeklyRule.findUniqueOrThrow({ where: { id: ruleId } });
  validateDateRange(existing.effectiveDate, endDate);

  const updated = await prisma.specialWeeklyRule.update({ where: { id: ruleId }, data: { endDate } });

  await prisma.changeLog.create({
    data: {
      tableName: "special_weekly_rules",
      recordId: ruleId,
      fieldName: "endDate",
      oldValue: existing.endDate ? existing.endDate.toISOString() : null,
      newValue: endDate.toISOString(),
      changedBy,
      reason: reason ?? "停用每週固定規則",
    },
  });

  return updated;
}

// 刪除：僅建議用於從未生效過或建立錯誤的規則。若規則已經生效一段期間，
// 請優先使用 deactivateSpecialWeeklyRule 設定結束日期，而不是直接刪除。
export async function deleteSpecialWeeklyRule(ruleId: string, changedBy?: string, reason?: string): Promise<void> {
  const existing = await prisma.specialWeeklyRule.findUniqueOrThrow({ where: { id: ruleId } });
  await prisma.specialWeeklyRule.delete({ where: { id: ruleId } });

  await prisma.changeLog.create({
    data: {
      tableName: "special_weekly_rules",
      recordId: ruleId,
      oldValue: describeWeeklyRule(existing),
      changedBy,
      reason: reason ?? "刪除每週固定規則",
    },
  });
}

export interface ListWeeklyRulesFilter {
  personId?: string;
  ruleType?: RuleType;
  weekday?: Weekday;
  periodCode?: string;
  projectId?: string;
  asOfDate?: Date; // 只列出在此日期仍然有效的規則（effectiveDate <= date <= endDate 或 endDate 為空）
}

// 4. 查詢：供管理介面搜尋教師／篩選星期／篩選專案
export async function listWeeklyRules(semesterId: string, filter: ListWeeklyRulesFilter = {}) {
  return prisma.specialWeeklyRule.findMany({
    where: {
      semesterId,
      ...(filter.personId ? { personId: filter.personId } : {}),
      ...(filter.ruleType ? { ruleType: filter.ruleType } : {}),
      ...(filter.weekday ? { weekday: filter.weekday } : {}),
      ...(filter.periodCode ? { periodCode: filter.periodCode } : {}),
      ...(filter.projectId ? { projectId: filter.projectId } : {}),
      ...(filter.asOfDate
        ? {
            effectiveDate: { lte: filter.asOfDate },
            OR: [{ endDate: null }, { endDate: { gte: filter.asOfDate } }],
          }
        : {}),
    },
    include: { person: true, project: true },
    orderBy: [{ weekday: "asc" }, { periodCode: "asc" }, { person: { name: "asc" } }],
  });
}

function describeWeeklyRule(rule: Pick<SpecialWeeklyRule, "ruleType" | "weekday" | "periodCode" | "weeklyPeriods">): string {
  return `${rule.ruleType} ${rule.weekday} ${rule.periodCode} x${rule.weeklyPeriods.toString()}`;
}

function stringifyValue(value: unknown): string | null {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return value.toISOString();
  return String(value);
}
