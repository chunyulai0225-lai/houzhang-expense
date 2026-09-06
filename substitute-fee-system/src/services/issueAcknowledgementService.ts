// Implementation Batch：月結問題確認
//
// 「已確認／接受」不是消除錯誤：這裡只記錄「管理者已經看過這筆問題、知道原因、
// 同意讓它不再阻擋這個月的月結」這件事本身。原始問題（SubstituteRecord 或
// MonthlyImportError）完全不會被修改，永遠可以查詢到原樣。
import type { MonthlyIssueAcknowledgement } from "@prisma/client";
import { prisma } from "../prismaClient";

export type AcknowledgeableTargetTable = "SubstituteRecord" | "MonthlyImportError";

export interface AcknowledgeIssueInput {
  semesterId: string;
  year: number;
  month: number;
  targetTable: AcknowledgeableTargetTable;
  targetId: string;
  reason: string;
  acknowledgedBy: string;
}

// 確認一筆問題「不再阻擋月結」。同一筆問題只保留一筆有效確認紀錄，
// 重複呼叫（例如更新理由）用 upsert，不會產生一堆重複資料列。
export async function acknowledgeIssue(input: AcknowledgeIssueInput): Promise<MonthlyIssueAcknowledgement> {
  if (!input.reason || !input.reason.trim()) {
    throw new Error("確認問題必須填寫理由");
  }
  if (!input.acknowledgedBy || !input.acknowledgedBy.trim()) {
    throw new Error("確認問題必須填寫確認人");
  }

  // 確認目標真的存在，不接受確認一筆不存在的問題（避免打錯 id 留下垃圾資料）。
  if (input.targetTable === "SubstituteRecord") {
    await prisma.substituteRecord.findUniqueOrThrow({ where: { id: input.targetId } });
  } else {
    await prisma.monthlyImportError.findUniqueOrThrow({ where: { id: input.targetId } });
  }

  const existing = await prisma.monthlyIssueAcknowledgement.findUnique({
    where: { targetTable_targetId: { targetTable: input.targetTable, targetId: input.targetId } },
  });

  const ack = await prisma.monthlyIssueAcknowledgement.upsert({
    where: { targetTable_targetId: { targetTable: input.targetTable, targetId: input.targetId } },
    create: {
      semesterId: input.semesterId,
      year: input.year,
      month: input.month,
      targetTable: input.targetTable,
      targetId: input.targetId,
      reason: input.reason,
      acknowledgedBy: input.acknowledgedBy,
    },
    update: {
      reason: input.reason,
      acknowledgedBy: input.acknowledgedBy,
      acknowledgedAt: new Date(),
    },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "monthly_issue_acknowledgements",
      recordId: ack.id,
      fieldName: existing ? "reason" : null,
      oldValue: existing?.reason ?? null,
      newValue: input.reason,
      changedBy: input.acknowledgedBy,
      reason: `確認接受問題（${input.targetTable} ${input.targetId}）：${input.reason}`,
    },
  });

  return ack;
}

// 撤銷確認：問題重新變回「待處理」。原始問題本身還是完全沒有被動過。
export async function revokeAcknowledgement(targetTable: AcknowledgeableTargetTable, targetId: string, changedBy?: string): Promise<void> {
  const existing = await prisma.monthlyIssueAcknowledgement.findUnique({
    where: { targetTable_targetId: { targetTable, targetId } },
  });
  if (!existing) return;

  await prisma.monthlyIssueAcknowledgement.delete({ where: { id: existing.id } });

  await prisma.changeLog.create({
    data: {
      tableName: "monthly_issue_acknowledgements",
      recordId: existing.id,
      fieldName: null,
      oldValue: existing.reason,
      newValue: null,
      changedBy,
      reason: `撤銷確認（${targetTable} ${targetId}），問題重新列為待處理`,
    },
  });
}

export async function listAcknowledgements(semesterId: string, year: number, month: number): Promise<MonthlyIssueAcknowledgement[]> {
  return prisma.monthlyIssueAcknowledgement.findMany({
    where: { semesterId, year, month },
    orderBy: { acknowledgedAt: "desc" },
  });
}

export async function getAcknowledgement(
  targetTable: AcknowledgeableTargetTable,
  targetId: string
): Promise<MonthlyIssueAcknowledgement | null> {
  return prisma.monthlyIssueAcknowledgement.findUnique({
    where: { targetTable_targetId: { targetTable, targetId } },
  });
}
