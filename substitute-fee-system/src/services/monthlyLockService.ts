// Implementation Batch：月結鎖定
//
// MonthlyLock 的 unique key 是 (year, month)，不含 semesterId（Phase 1 schema 本來就這樣設計，
// 這裡沿用、不修改）。實務上一個 (year, month) 只會落在一個學期裡，所以鎖定狀態本身不需要
// 知道是哪個學期；但檢查「這個月還有哪些阻擋問題」時，需要 semesterId 去找出對應的
// MonthlyImport 批次（因為 MonthlyImport 是用 semesterId+year+month+sourceStaffType 找的）。
//
// 「已確認／接受」不會消除問題，只是讓它不再阻擋鎖定——所以這裡的「blocking count」
// 一律是「尚未被 MonthlyIssueAcknowledgement 認領」的數量，原始問題本身永遠查得到。
import { MonthlyStatus } from "@prisma/client";
import { prisma } from "../prismaClient";

export interface BlockingIssueSummary {
  teacherUnmatched: number;
  conflict: number;
  amountMissing: number; // 已分類為 GENERAL/OVERTIME/PROJECT，但金額算不出來（Raw缺漏又沒有FeeRule）
  importErrors: number; // 匯入錯誤（日期/節次無法解析等），尚未被確認接受的
  total: number;
}

async function getActiveMonthlyImportIds(semesterId: string, year: number, month: number): Promise<string[]> {
  const imports = await prisma.monthlyImport.findMany({
    where: { semesterId, year, month, status: "ACTIVE" },
    select: { id: true },
  });
  return imports.map((i) => i.id);
}

async function getAcknowledgedIds(targetTable: "SubstituteRecord" | "MonthlyImportError", targetIds: string[]): Promise<Set<string>> {
  if (targetIds.length === 0) return new Set();
  const acks = await prisma.monthlyIssueAcknowledgement.findMany({
    where: { targetTable, targetId: { in: targetIds } },
    select: { targetId: true },
  });
  return new Set(acks.map((a) => a.targetId));
}

// 計算某學期某年月「尚未被確認接受」的阻擋性問題數量。
// 只看 entryType=EXCEL_IMPORT 的紀錄——自費代課（MANUAL_SELF_FUNDED）建立時就已經是
// 管理者確認過的資料，不會落入公費分類/自動配對的判斷邏輯，不應該被當成「待處理」。
export async function getBlockingIssues(semesterId: string, year: number, month: number): Promise<BlockingIssueSummary> {
  const importIds = await getActiveMonthlyImportIds(semesterId, year, month);
  if (importIds.length === 0) {
    return { teacherUnmatched: 0, conflict: 0, amountMissing: 0, importErrors: 0, total: 0 };
  }

  const [teacherUnmatchedRows, conflictRows, amountMissingRows, importErrorRows] = await Promise.all([
    prisma.substituteRecord.findMany({
      where: { monthlyImportId: { in: importIds }, entryType: "EXCEL_IMPORT", classificationMethod: "TEACHER_UNMATCHED" },
      select: { id: true },
    }),
    prisma.substituteRecord.findMany({
      where: { monthlyImportId: { in: importIds }, entryType: "EXCEL_IMPORT", classificationMethod: "CONFLICT" },
      select: { id: true },
    }),
    prisma.substituteRecord.findMany({
      where: {
        monthlyImportId: { in: importIds },
        entryType: "EXCEL_IMPORT",
        fundingSource: { in: ["GENERAL", "OVERTIME", "PROJECT"] },
        amount: null,
      },
      select: { id: true },
    }),
    prisma.monthlyImportError.findMany({
      where: { monthlyImportId: { in: importIds } },
      select: { id: true },
    }),
  ]);

  const [ackTeacherUnmatched, ackConflict, ackAmountMissing, ackImportErrors] = await Promise.all([
    getAcknowledgedIds("SubstituteRecord", teacherUnmatchedRows.map((r) => r.id)),
    getAcknowledgedIds("SubstituteRecord", conflictRows.map((r) => r.id)),
    getAcknowledgedIds("SubstituteRecord", amountMissingRows.map((r) => r.id)),
    getAcknowledgedIds("MonthlyImportError", importErrorRows.map((r) => r.id)),
  ]);

  const teacherUnmatched = teacherUnmatchedRows.filter((r) => !ackTeacherUnmatched.has(r.id)).length;
  const conflict = conflictRows.filter((r) => !ackConflict.has(r.id)).length;
  const amountMissing = amountMissingRows.filter((r) => !ackAmountMissing.has(r.id)).length;
  const importErrors = importErrorRows.filter((r) => !ackImportErrors.has(r.id)).length;

  return {
    teacherUnmatched,
    conflict,
    amountMissing,
    importErrors,
    total: teacherUnmatched + conflict + amountMissing + importErrors,
  };
}

export interface MonthlyLockStatusResult {
  year: number;
  month: number;
  status: MonthlyStatus;
  isLocked: boolean;
  lockedAt: Date | null;
  lockedBy: string | null;
  unlockedAt: Date | null;
  unlockedBy: string | null;
}

async function getOrCreateLockRow(semesterId: string, year: number, month: number) {
  const existing = await prisma.monthlyLock.findUnique({ where: { year_month: { year, month } } });
  if (existing) return existing;
  return prisma.monthlyLock.create({ data: { semesterId, year, month } });
}

export async function getMonthlyLockStatus(semesterId: string, year: number, month: number): Promise<MonthlyLockStatusResult> {
  const lock = await getOrCreateLockRow(semesterId, year, month);
  return {
    year: lock.year,
    month: lock.month,
    status: lock.status,
    isLocked: lock.status === "LOCKED",
    lockedAt: lock.lockedAt,
    lockedBy: lock.lockedBy,
    unlockedAt: lock.unlockedAt,
    unlockedBy: lock.unlockedBy,
  };
}

// 給其他 service 用的守門檢查：只看 (year, month) 是否已鎖定，不需要知道 semesterId。
// 找不到 MonthlyLock 資料列時視為「尚未鎖定」（NOT_IMPORTED 是預設狀態，本來就可以修改）。
export async function assertMonthNotLocked(year: number, month: number): Promise<void> {
  const lock = await prisma.monthlyLock.findUnique({ where: { year_month: { year, month } } });
  if (lock?.status === "LOCKED") {
    throw new Error(`${year}年${month}月已經鎖定，不能修改。如需修改請先由管理者解除鎖定並留下理由。`);
  }
}

// 給只知道 monthlyImportId 的呼叫端用（匯入、費用計算等批次操作）。
export async function assertImportMonthNotLocked(monthlyImportId: string): Promise<void> {
  const monthlyImport = await prisma.monthlyImport.findUniqueOrThrow({
    where: { id: monthlyImportId },
    select: { year: true, month: true },
  });
  await assertMonthNotLocked(monthlyImport.year, monthlyImport.month);
}

// 給只知道 substituteRecordId 的呼叫端用（分類覆寫/復原、單筆費用計算等）。
export async function assertRecordMonthNotLocked(substituteRecordId: string): Promise<void> {
  const record = await prisma.substituteRecord.findUniqueOrThrow({
    where: { id: substituteRecordId },
    select: { monthlyImport: { select: { year: true, month: true } } },
  });
  await assertMonthNotLocked(record.monthlyImport.year, record.monthlyImport.month);
}

export async function lockMonth(semesterId: string, year: number, month: number, lockedBy: string): Promise<MonthlyLockStatusResult> {
  if (!lockedBy || !lockedBy.trim()) {
    throw new Error("鎖定必須填寫操作人");
  }
  const issues = await getBlockingIssues(semesterId, year, month);
  if (issues.total > 0) {
    throw new Error(
      `尚有 ${issues.total} 筆未確認的阻擋問題無法鎖定：原教師未配對 ${issues.teacherUnmatched} 筆、規則衝突 ${issues.conflict} 筆、` +
        `金額無法計算 ${issues.amountMissing} 筆、匯入錯誤 ${issues.importErrors} 筆。請先處理或在「待處理」頁面確認接受。`
    );
  }

  const lock = await getOrCreateLockRow(semesterId, year, month);
  const updated = await prisma.monthlyLock.update({
    where: { id: lock.id },
    data: { status: "LOCKED", lockedAt: new Date(), lockedBy, unlockedAt: null, unlockedBy: null },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "monthly_locks",
      recordId: updated.id,
      fieldName: "status",
      oldValue: lock.status,
      newValue: "LOCKED",
      changedBy: lockedBy,
      reason: `鎖定 ${year}年${month}月`,
    },
  });

  return getMonthlyLockStatus(semesterId, year, month);
}

export async function unlockMonth(
  semesterId: string,
  year: number,
  month: number,
  unlockedBy: string,
  reason: string
): Promise<MonthlyLockStatusResult> {
  if (!unlockedBy || !unlockedBy.trim()) {
    throw new Error("解鎖必須填寫操作人");
  }
  if (!reason || !reason.trim()) {
    throw new Error("解鎖必須填寫理由");
  }

  const lock = await prisma.monthlyLock.findUnique({ where: { year_month: { year, month } } });
  if (!lock || lock.status !== "LOCKED") {
    throw new Error(`${year}年${month}月目前不是鎖定狀態，不需要解鎖`);
  }

  const updated = await prisma.monthlyLock.update({
    where: { id: lock.id },
    data: { status: "NOT_IMPORTED", unlockedAt: new Date(), unlockedBy },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "monthly_locks",
      recordId: updated.id,
      fieldName: "status",
      oldValue: "LOCKED",
      newValue: "NOT_IMPORTED",
      changedBy: unlockedBy,
      reason: `解鎖 ${year}年${month}月：${reason}`,
    },
  });

  return getMonthlyLockStatus(semesterId, year, month);
}
