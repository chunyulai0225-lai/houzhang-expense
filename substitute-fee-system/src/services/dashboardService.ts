// Implementation Batch：月結首頁
//
// 所有統計都是即時查詢既有資料算出來的，沒有另外存一份快照——
// 這裡只是把已經存在的資料（MonthlyImport／SubstituteRecord／MonthlyImportError／
// MonthlyLock）組合成一次回應，方便前端一次顯示整個月的狀態。
import { prisma } from "../prismaClient";
import { getBlockingIssues, getMonthlyLockStatus } from "./monthlyLockService";
import { listAcknowledgements } from "./issueAcknowledgementService";

export interface MonthlyDashboard {
  year: number;
  month: number;
  import: {
    batches: { id: string; sourceStaffType: string; versionNo: number; fileName: string; totalCount: number; successCount: number; errorCount: number }[];
    hasActiveBatch: boolean;
    successCount: number;
    errorCount: number;
  };
  selfFunded: {
    exists: boolean;
    count: number;
  };
  classification: {
    general: number;
    overtime: number;
    project: number;
    conflict: number;
    teacherUnmatched: number;
  };
  fee: {
    calculatedCount: number;
    notCalculatedCount: number;
    totalAmount: string;
  };
  issues: {
    blocking: Awaited<ReturnType<typeof getBlockingIssues>>;
    acknowledgedCount: number;
  };
  lock: Awaited<ReturnType<typeof getMonthlyLockStatus>>;
}

export async function getMonthlyDashboard(semesterId: string, year: number, month: number): Promise<MonthlyDashboard> {
  const allBatches = await prisma.monthlyImport.findMany({
    where: { semesterId, year, month },
    orderBy: [{ sourceStaffType: "asc" }, { versionNo: "desc" }],
  });
  const activeBatches = allBatches.filter((b) => b.status === "ACTIVE");
  const excelActiveBatches = activeBatches.filter((b) => b.fileName !== "（自費代課手動登錄）");
  const selfFundedBatch = activeBatches.find((b) => b.fileName === "（自費代課手動登錄）");
  const activeBatchIds = activeBatches.map((b) => b.id);

  const byFunding = activeBatchIds.length
    ? await prisma.substituteRecord.groupBy({
        by: ["fundingSource", "classificationMethod"],
        where: { monthlyImportId: { in: activeBatchIds }, entryType: "EXCEL_IMPORT" },
        _count: true,
      })
    : [];

  const classification = { general: 0, overtime: 0, project: 0, conflict: 0, teacherUnmatched: 0 };
  for (const row of byFunding) {
    if (row.classificationMethod === "CONFLICT") classification.conflict += row._count;
    else if (row.classificationMethod === "TEACHER_UNMATCHED") classification.teacherUnmatched += row._count;
    else if (row.fundingSource === "GENERAL") classification.general += row._count;
    else if (row.fundingSource === "OVERTIME") classification.overtime += row._count;
    else if (row.fundingSource === "PROJECT") classification.project += row._count;
  }

  // 「費用狀態」只看公費（Excel 匯入）批次，自費代課有自己獨立的區塊顯示，
  // 不要混在同一個總額裡，否則會誤導使用者以為公費支出比實際還多。
  const excelActiveBatchIds = excelActiveBatches.map((b) => b.id);
  const [calculatedCount, notCalculatedCount, amountAgg, selfFundedCount] = excelActiveBatchIds.length
    ? await Promise.all([
        prisma.substituteRecord.count({ where: { monthlyImportId: { in: excelActiveBatchIds }, amount: { not: null } } }),
        prisma.substituteRecord.count({ where: { monthlyImportId: { in: excelActiveBatchIds }, amount: null } }),
        prisma.substituteRecord.aggregate({ where: { monthlyImportId: { in: excelActiveBatchIds }, amount: { not: null } }, _sum: { amount: true } }),
        selfFundedBatch
          ? prisma.substituteRecord.count({ where: { monthlyImportId: selfFundedBatch.id, entryType: "MANUAL_SELF_FUNDED" } })
          : Promise.resolve(0),
      ])
    : [0, 0, { _sum: { amount: null } }, 0];

  const [blocking, acknowledgements, lock] = await Promise.all([
    getBlockingIssues(semesterId, year, month),
    listAcknowledgements(semesterId, year, month),
    getMonthlyLockStatus(semesterId, year, month),
  ]);

  return {
    year,
    month,
    import: {
      batches: excelActiveBatches.map((b) => ({
        id: b.id,
        sourceStaffType: b.sourceStaffType,
        versionNo: b.versionNo,
        fileName: b.fileName,
        totalCount: b.totalCount,
        successCount: b.successCount,
        errorCount: b.errorCount,
      })),
      hasActiveBatch: excelActiveBatches.length > 0,
      successCount: excelActiveBatches.reduce((s, b) => s + b.successCount, 0),
      errorCount: excelActiveBatches.reduce((s, b) => s + b.errorCount, 0),
    },
    selfFunded: { exists: !!selfFundedBatch, count: selfFundedCount },
    classification,
    fee: {
      calculatedCount,
      notCalculatedCount,
      totalAmount: (amountAgg._sum.amount ?? 0).toString(),
    },
    issues: { blocking, acknowledgedCount: acknowledgements.length },
    lock,
  };
}
