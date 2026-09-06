// Phase 2：教職員工代號 PDF 匯入架構
//
// 重要：PDF 的實際欄位配置尚未取得（49 頁教職員工代號 PDF），因此這裡刻意把
// 「PDF 文字擷取」與「一列資料如何處理」分開——本檔案只負責後者：
// 呼叫端把 PDF 解析出的列（員工代號／姓名／部門名稱）傳進來，交由管理者
// 逐列確認後才會建立 PersonCode，絕不自動把 PDF 每一列當成一個新 Person。
//
// 等拿到實際 PDF 樣本、確認其表格結構後，再另外實作「PDF → StaffDirectoryRawRow[]」
// 的擷取步驟，銜接到 startStaffDirectoryImport()，不需要更動這裡的資料模型或流程。
//
// categoryCode（供未來 BD/非BD 判斷使用）目前刻意不處理：
// 不自行從部門名稱推導，也不要求管理者逐列手動輸入。這裡只先保存
// 員工代號／姓名／部門名稱／匯入批次／配對狀態。等拿到真實 PDF、
// 確認「員工代號／部門名稱 → 身分類別」的實際對應規則後，再另外做
// 一個可以批次套用的解析規則（不會是逐筆手填）。

import type { ImportRowMatchStatus, Person } from "@prisma/client";
import { prisma } from "../prismaClient";

export interface StaffDirectoryRawRow {
  rowNumber?: number;
  employeeCode: string; // 員工代號
  name: string; // 姓名
  departmentName?: string; // 部門名稱（身分類別文字，例如「代理教師」）
  schoolYear?: number; // 此份 PDF 對應的學年度，由匯入時人工指定
}

export async function startStaffDirectoryImport(params: {
  fileName: string;
  importedBy?: string;
  rows: StaffDirectoryRawRow[];
}) {
  const batch = await prisma.staffDirectoryImport.create({
    data: {
      fileName: params.fileName,
      importedBy: params.importedBy,
      totalCount: params.rows.length,
    },
  });

  for (const row of params.rows) {
    const normalizedName = row.name.trim();
    await prisma.staffDirectoryImportRow.create({
      data: {
        importId: batch.id,
        rowNumber: row.rowNumber,
        employeeCode: row.employeeCode,
        name: normalizedName,
        departmentName: row.departmentName,
        schoolYear: row.schoolYear,
      },
    });
  }

  return prisma.staffDirectoryImport.findUniqueOrThrow({
    where: { id: batch.id },
    include: { rows: true },
  });
}

// 候選人清單一律「即時」查詢，不使用匯入當下快取的建議，
// 這樣同一批匯入中，前面幾列已經確認建立的新人，後面列也能立刻查到並可選來配對，
// 但仍然只是候選清單，不會自動配對。
export async function getCandidatePersonsForRow(rowId: string): Promise<Person[]> {
  const row = await prisma.staffDirectoryImportRow.findUniqueOrThrow({ where: { id: rowId } });
  return prisma.person.findMany({ where: { name: row.name }, orderBy: { createdAt: "asc" } });
}

export interface ImportRowWithCandidates {
  row: Awaited<ReturnType<typeof prisma.staffDirectoryImportRow.findFirstOrThrow>>;
  candidates: Person[];
}

export async function listImportRows(
  importId: string,
  matchStatus?: ImportRowMatchStatus
): Promise<ImportRowWithCandidates[]> {
  const rows = await prisma.staffDirectoryImportRow.findMany({
    where: { importId, ...(matchStatus ? { matchStatus } : {}) },
    include: { matchedPerson: true },
    orderBy: { rowNumber: "asc" },
  });

  return Promise.all(
    rows.map(async (row) => ({
      row,
      candidates: await prisma.person.findMany({ where: { name: row.name }, orderBy: { createdAt: "asc" } }),
    }))
  );
}

export type ResolveImportRowInput =
  | { action: "MATCH_EXISTING"; personId: string }
  | { action: "CREATE_NEW"; personOverrides?: { name?: string; payrollCode?: string } }
  | { action: "IGNORE" };

// 管理者對單一 PDF 列做出的最終決定：配對既有人員／建立新人員／略過。
// 只有在這一步之後，才會真正寫入 PersonCode。categoryCode 暫不處理，見檔案頂端說明。
export async function resolveImportRow(rowId: string, input: ResolveImportRowInput, resolvedBy?: string) {
  const row = await prisma.staffDirectoryImportRow.findUniqueOrThrow({ where: { id: rowId } });

  if (row.matchStatus !== "PENDING") {
    throw new Error(`此列已處理過（狀態：${row.matchStatus}），不可重複處理`);
  }

  if (input.action === "IGNORE") {
    return prisma.staffDirectoryImportRow.update({
      where: { id: rowId },
      data: { matchStatus: "IGNORED", resolvedAt: new Date(), resolvedBy },
    });
  }

  let personId: string;
  let matchStatus: Extract<ImportRowMatchStatus, "MATCHED_EXISTING" | "CREATED_NEW">;

  if (input.action === "MATCH_EXISTING") {
    personId = input.personId;
    matchStatus = "MATCHED_EXISTING";
  } else {
    const created = await prisma.person.create({
      data: {
        name: input.personOverrides?.name ?? row.name,
        payrollCode: input.personOverrides?.payrollCode,
      },
    });
    personId = created.id;
    matchStatus = "CREATED_NEW";
  }

  const personCode = await prisma.personCode.create({
    data: {
      personId,
      schoolYear: row.schoolYear ?? undefined,
      categoryName: row.departmentName ?? "未分類",
      originalStaffCode: row.employeeCode,
    },
  });

  return prisma.staffDirectoryImportRow.update({
    where: { id: rowId },
    data: {
      matchStatus,
      matchedPersonId: personId,
      createdPersonCodeId: personCode.id,
      resolvedAt: new Date(),
      resolvedBy,
    },
  });
}
