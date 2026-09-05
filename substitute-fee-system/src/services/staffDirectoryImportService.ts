// Phase 2：教職員工代號 PDF 匯入架構
//
// 重要：PDF 的實際欄位配置尚未取得（49 頁教職員工代號 PDF），因此這裡刻意把
// 「PDF 文字擷取」與「一列資料如何處理」分開——本檔案只負責後者：
// 呼叫端把 PDF 解析出的列（員工代號／姓名／部門名稱）傳進來，
// 由這裡依姓名找出「可能的既有人員」作為建議，但一律交由管理者確認後
// 才會建立 PersonCode，絕不自動把 PDF 每一列當成一個新 Person。
//
// 等拿到實際 PDF 樣本、確認其表格結構後，再另外實作「PDF → StaffDirectoryRawRow[]」
// 的擷取步驟，銜接到 startStaffDirectoryImport()，不需要更動這裡的資料模型或流程。

import type { ImportRowMatchStatus } from "@prisma/client";
import { prisma } from "../prismaClient";

export interface StaffDirectoryRawRow {
  rowNumber?: number;
  employeeCode: string; // 員工代號
  name: string; // 姓名
  departmentName?: string; // 部門名稱（身分類別文字，例如「代理教師」）
  categoryCode?: string; // 若 PDF 另有可辨識的分類代碼（例如 B/D），先保留欄位、不強制
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
    // 僅作為「建議」，不會自動配對——即使剛好找到唯一同名人員，仍須由管理者確認。
    const suggested = await prisma.person.findFirst({ where: { name: normalizedName } });

    await prisma.staffDirectoryImportRow.create({
      data: {
        importId: batch.id,
        rowNumber: row.rowNumber,
        employeeCode: row.employeeCode,
        name: normalizedName,
        departmentName: row.departmentName,
        schoolYear: row.schoolYear,
        suggestedPersonId: suggested?.id,
      },
    });
  }

  return prisma.staffDirectoryImport.findUniqueOrThrow({
    where: { id: batch.id },
    include: { rows: { include: { suggestedPerson: true } } },
  });
}

export async function listImportRows(importId: string, matchStatus?: ImportRowMatchStatus) {
  return prisma.staffDirectoryImportRow.findMany({
    where: { importId, ...(matchStatus ? { matchStatus } : {}) },
    include: { suggestedPerson: true, matchedPerson: true },
    orderBy: { rowNumber: "asc" },
  });
}

export type ResolveImportRowInput =
  | { action: "MATCH_EXISTING"; personId: string; categoryCode?: string }
  | { action: "CREATE_NEW"; personOverrides?: { name?: string; payrollCode?: string }; categoryCode?: string }
  | { action: "IGNORE" };

// 管理者對單一 PDF 列做出的最終決定：配對既有人員／建立新人員／略過。
// 只有在這一步之後，才會真正寫入 PersonCode。
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

  const categoryCode = "categoryCode" in input ? input.categoryCode ?? row.categoryCode : row.categoryCode;

  const personCode = await prisma.personCode.create({
    data: {
      personId,
      schoolYear: row.schoolYear ?? undefined,
      categoryName: row.departmentName ?? "未分類",
      categoryCode: categoryCode ?? undefined,
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
