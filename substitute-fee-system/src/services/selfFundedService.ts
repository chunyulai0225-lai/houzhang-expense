// Implementation Batch：自費代課
//
// 自費代課不是 Excel 匯入資料，沒有對應的 SubstituteRecordRaw，entryType 一律是
// MANUAL_SELF_FUNDED。刻意不設定 fundingSource（維持預設 UNDETERMINED）、
// 不跑分類引擎、不跑 Phase9-5 的費用計算——金額由管理者在建立當下直接輸入，
// 這樣才不會污染 GENERAL/OVERTIME/PROJECT/CONFLICT 這一整套公費分類邏輯，
// 也不會被「待處理」清單誤判成需要處理的公費異常。
//
// 為了不需要讓 SubstituteRecord.monthlyImportId 也變成可為 null（多動一個既有必填欄位），
// 每個學期＋年月的自費代課，共用同一筆「虛擬」MonthlyImport 容器（sourceStaffType=UNKNOWN，
// 用固定的 fileName 標記辨識），這樣月結鎖定、月份彙總查詢都能沿用現有以
// monthlyImportId 為主的查詢方式，不用額外多開一條路徑。
import type { SubstituteRecord, Weekday } from "@prisma/client";
import { prisma } from "../prismaClient";
import { assertMonthNotLocked, assertRecordMonthNotLocked } from "./monthlyLockService";

const SELF_FUNDED_BATCH_FILE_NAME = "（自費代課手動登錄）";

const WEEKDAY_BY_JS_INDEX: Weekday[] = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"];
function weekdayOf(date: Date): Weekday {
  return WEEKDAY_BY_JS_INDEX[date.getUTCDay()];
}

async function getOrCreateSelfFundedBatch(semesterId: string, year: number, month: number) {
  const existing = await prisma.monthlyImport.findFirst({
    where: { semesterId, year, month, sourceStaffType: "UNKNOWN", fileName: SELF_FUNDED_BATCH_FILE_NAME },
  });
  if (existing) return existing;
  return prisma.monthlyImport.create({
    data: {
      semesterId,
      year,
      month,
      fileName: SELF_FUNDED_BATCH_FILE_NAME,
      sourceStaffType: "UNKNOWN",
      status: "ACTIVE",
      totalCount: 0,
      successCount: 0,
    },
  });
}

export interface CreateSelfFundedRecordInput {
  semesterId: string;
  year: number;
  month: number;
  date: Date;
  originalTeacherId?: string;
  substituteTeacherId: string;
  periodCode?: string;
  className?: string;
  subject?: string;
  amount: number | string;
  unitPrice?: number | string;
  note?: string;
  createdBy: string;
}

export async function createSelfFundedRecord(input: CreateSelfFundedRecordInput): Promise<SubstituteRecord> {
  if (!input.createdBy || !input.createdBy.trim()) {
    throw new Error("建立自費代課必須填寫建立人");
  }
  if (!input.substituteTeacherId) {
    throw new Error("自費代課必須指定代課教師");
  }
  await assertMonthNotLocked(input.year, input.month);

  const batch = await getOrCreateSelfFundedBatch(input.semesterId, input.year, input.month);

  const record = await prisma.substituteRecord.create({
    data: {
      monthlyImportId: batch.id,
      rawRecordId: null,
      entryType: "MANUAL_SELF_FUNDED",
      date: input.date,
      weekday: weekdayOf(input.date),
      originalTeacherId: input.originalTeacherId,
      substituteTeacherId: input.substituteTeacherId,
      periodCode: input.periodCode,
      className: input.className,
      subject: input.subject,
      amount: input.amount,
      unitPrice: input.unitPrice ?? input.amount,
      note: input.note,
      createdBy: input.createdBy,
    },
  });

  await prisma.monthlyImport.update({
    where: { id: batch.id },
    data: { totalCount: { increment: 1 }, successCount: { increment: 1 } },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "substitute_records",
      recordId: record.id,
      fieldName: null,
      oldValue: null,
      newValue: `建立自費代課：${input.date.toISOString().slice(0, 10)} 金額 ${record.amount?.toString()}`,
      changedBy: input.createdBy,
      reason: input.note ?? "手動建立自費代課",
    },
  });

  return record;
}

export interface UpdateSelfFundedRecordInput {
  date?: Date;
  originalTeacherId?: string | null;
  substituteTeacherId?: string;
  periodCode?: string | null;
  className?: string | null;
  subject?: string | null;
  amount?: number | string;
  unitPrice?: number | string;
  note?: string | null;
  updatedBy: string;
}

export async function updateSelfFundedRecord(recordId: string, input: UpdateSelfFundedRecordInput): Promise<SubstituteRecord> {
  if (!input.updatedBy || !input.updatedBy.trim()) {
    throw new Error("修改自費代課必須填寫操作人");
  }
  const existing = await prisma.substituteRecord.findUniqueOrThrow({ where: { id: recordId } });
  if (existing.entryType !== "MANUAL_SELF_FUNDED") {
    throw new Error("這筆不是自費代課紀錄，不能用這個功能修改");
  }
  await assertRecordMonthNotLocked(recordId);

  const data: Record<string, unknown> = { updatedBy: input.updatedBy };
  if (input.date !== undefined) {
    data.date = input.date;
    data.weekday = weekdayOf(input.date);
  }
  if (input.originalTeacherId !== undefined) data.originalTeacherId = input.originalTeacherId;
  if (input.substituteTeacherId !== undefined) data.substituteTeacherId = input.substituteTeacherId;
  if (input.periodCode !== undefined) data.periodCode = input.periodCode;
  if (input.className !== undefined) data.className = input.className;
  if (input.subject !== undefined) data.subject = input.subject;
  if (input.amount !== undefined) data.amount = input.amount;
  if (input.unitPrice !== undefined) data.unitPrice = input.unitPrice;
  if (input.note !== undefined) data.note = input.note;

  const updated = await prisma.substituteRecord.update({ where: { id: recordId }, data });

  await prisma.changeLog.create({
    data: {
      tableName: "substitute_records",
      recordId,
      fieldName: null,
      oldValue: JSON.stringify({ amount: existing.amount?.toString(), date: existing.date }),
      newValue: JSON.stringify({ amount: updated.amount?.toString(), date: updated.date }),
      changedBy: input.updatedBy,
      reason: "修改自費代課",
    },
  });

  return updated;
}

export async function deleteSelfFundedRecord(recordId: string, deletedBy: string, reason: string): Promise<void> {
  if (!deletedBy || !deletedBy.trim()) throw new Error("刪除自費代課必須填寫操作人");
  if (!reason || !reason.trim()) throw new Error("刪除自費代課必須填寫理由");

  const existing = await prisma.substituteRecord.findUniqueOrThrow({ where: { id: recordId } });
  if (existing.entryType !== "MANUAL_SELF_FUNDED") {
    throw new Error("這筆不是自費代課紀錄，不能用這個功能刪除");
  }
  await assertRecordMonthNotLocked(recordId);

  await prisma.changeLog.create({
    data: {
      tableName: "substitute_records",
      recordId,
      fieldName: null,
      oldValue: JSON.stringify({ amount: existing.amount?.toString(), date: existing.date }),
      newValue: null,
      changedBy: deletedBy,
      reason: `刪除自費代課：${reason}`,
    },
  });

  await prisma.substituteRecord.delete({ where: { id: recordId } });
  await prisma.monthlyImport.update({
    where: { id: existing.monthlyImportId },
    data: { totalCount: { decrement: 1 }, successCount: { decrement: 1 } },
  });
}

export async function listSelfFundedRecords(semesterId: string, year: number, month: number) {
  const batch = await prisma.monthlyImport.findFirst({
    where: { semesterId, year, month, sourceStaffType: "UNKNOWN", fileName: SELF_FUNDED_BATCH_FILE_NAME },
  });
  if (!batch) return [];
  return prisma.substituteRecord.findMany({
    where: { monthlyImportId: batch.id, entryType: "MANUAL_SELF_FUNDED" },
    include: { originalTeacher: true, substituteTeacher: true },
    orderBy: { date: "asc" },
  });
}
