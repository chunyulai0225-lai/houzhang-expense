// Implementation Batch：待處理工作區
//
// 把原本分散在「公費代課匯入」（未配對教師、匯入錯誤）跟「分類預覽」（規則衝突）
// 兩個地方的問題，整合成一個清單。每一筆都標示目前狀態：
// 待處理 / 已解決（原始問題已經不存在了）/ 已確認接受（有 MonthlyIssueAcknowledgement）。
import { prisma } from "../prismaClient";
import { getAcknowledgement } from "./issueAcknowledgementService";

export type PendingIssueType = "TEACHER_UNMATCHED" | "CONFLICT" | "AMOUNT_MISSING" | "IMPORT_ERROR";
export type PendingIssueStatus = "PENDING" | "ACKNOWLEDGED";

export interface PendingIssueRow {
  issueType: PendingIssueType;
  targetTable: "SubstituteRecord" | "MonthlyImportError";
  targetId: string;
  date: string | null;
  originalTeacher: string | null;
  substituteTeacher: string | null;
  periodCode: string | null;
  className: string | null;
  subject: string | null;
  description: string;
  status: PendingIssueStatus;
  acknowledgement: { reason: string; acknowledgedBy: string; acknowledgedAt: string } | null;
}

async function toAckInfo(targetTable: "SubstituteRecord" | "MonthlyImportError", targetId: string) {
  const ack = await getAcknowledgement(targetTable, targetId);
  if (!ack) return { status: "PENDING" as const, acknowledgement: null };
  return {
    status: "ACKNOWLEDGED" as const,
    acknowledgement: { reason: ack.reason, acknowledgedBy: ack.acknowledgedBy, acknowledgedAt: ack.acknowledgedAt.toISOString() },
  };
}

export async function listPendingIssues(semesterId: string, year: number, month: number): Promise<PendingIssueRow[]> {
  const batches = await prisma.monthlyImport.findMany({ where: { semesterId, year, month, status: "ACTIVE" } });
  const batchIds = batches.map((b) => b.id);
  if (batchIds.length === 0) return [];

  const [teacherUnmatched, conflict, amountMissing, importErrors] = await Promise.all([
    prisma.substituteRecord.findMany({
      where: { monthlyImportId: { in: batchIds }, entryType: "EXCEL_IMPORT", classificationMethod: "TEACHER_UNMATCHED" },
      include: { rawRecord: true, substituteTeacher: true },
    }),
    prisma.substituteRecord.findMany({
      where: { monthlyImportId: { in: batchIds }, entryType: "EXCEL_IMPORT", classificationMethod: "CONFLICT" },
      include: { rawRecord: true, originalTeacher: true, substituteTeacher: true },
    }),
    prisma.substituteRecord.findMany({
      where: {
        monthlyImportId: { in: batchIds },
        entryType: "EXCEL_IMPORT",
        fundingSource: { in: ["GENERAL", "OVERTIME", "PROJECT"] },
        amount: null,
      },
      include: { rawRecord: true, originalTeacher: true, substituteTeacher: true },
    }),
    prisma.monthlyImportError.findMany({ where: { monthlyImportId: { in: batchIds } } }),
  ]);

  const rows: PendingIssueRow[] = [];

  for (const r of teacherUnmatched) {
    const ack = await toAckInfo("SubstituteRecord", r.id);
    rows.push({
      issueType: "TEACHER_UNMATCHED",
      targetTable: "SubstituteRecord",
      targetId: r.id,
      date: r.date.toISOString().slice(0, 10),
      originalTeacher: r.rawRecord?.originalTeacherText ?? null,
      substituteTeacher: r.substituteTeacher?.name ?? r.rawRecord?.substituteTeacherText ?? null,
      periodCode: r.periodCode,
      className: r.className,
      subject: r.subject,
      description: "原教師姓名尚未配對到人員資料，無法判斷分類規則",
      ...ack,
    });
  }

  for (const r of conflict) {
    const ack = await toAckInfo("SubstituteRecord", r.id);
    rows.push({
      issueType: "CONFLICT",
      targetTable: "SubstituteRecord",
      targetId: r.id,
      date: r.date.toISOString().slice(0, 10),
      originalTeacher: r.originalTeacher?.name ?? r.rawRecord?.originalTeacherText ?? null,
      substituteTeacher: r.substituteTeacher?.name ?? r.rawRecord?.substituteTeacherText ?? null,
      periodCode: r.periodCode,
      className: r.className,
      subject: r.subject,
      description: "同時符合多個規則，系統不會自動選一個，需要人工確認",
      ...ack,
    });
  }

  for (const r of amountMissing) {
    const ack = await toAckInfo("SubstituteRecord", r.id);
    rows.push({
      issueType: "AMOUNT_MISSING",
      targetTable: "SubstituteRecord",
      targetId: r.id,
      date: r.date.toISOString().slice(0, 10),
      originalTeacher: r.originalTeacher?.name ?? r.rawRecord?.originalTeacherText ?? null,
      substituteTeacher: r.substituteTeacher?.name ?? r.rawRecord?.substituteTeacherText ?? null,
      periodCode: r.periodCode,
      className: r.className,
      subject: r.subject,
      description: `已分類為${r.fundingSource}，但原始資料沒有金額、也找不到生效中的費率，無法計算`,
      ...ack,
    });
  }

  for (const e of importErrors) {
    const ack = await toAckInfo("MonthlyImportError", e.id);
    rows.push({
      issueType: "IMPORT_ERROR",
      targetTable: "MonthlyImportError",
      targetId: e.id,
      date: null,
      originalTeacher: null,
      substituteTeacher: null,
      periodCode: null,
      className: null,
      subject: null,
      description: `第 ${e.rowNumber ?? "?"} 列${e.fieldName ? `（${e.fieldName}）` : ""}：${e.message}`,
      ...ack,
    });
  }

  return rows;
}
