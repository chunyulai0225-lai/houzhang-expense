// Implementation Batch：給出納彙總與 Excel 匯出
//
// 完全不重新計算金額，直接沿用 Phase9-5 已驗證的 summarizeTeacherMonthlyFees。
// 這裡只是把結果依 BD／非BD 分組、補上薪資代碼，並提供 Excel 匯出格式，
// 排版盡量貼近真實「給出納」檔案（代課教師代碼／代課教師／節數／金額／備註）。
import ExcelJS from "exceljs";
import { prisma } from "../prismaClient";
import { summarizeTeacherMonthlyFees, TeacherMonthlyFeeSummaryRow } from "./feeCalculationService";
import { SELF_FUNDED_BATCH_FILE_NAME } from "./selfFundedService";

export interface ChunaRow extends TeacherMonthlyFeeSummaryRow {
  payrollCode: string | null;
}

export interface ChunaSummary {
  bd: ChunaRow[];
  nonBd: ChunaRow[];
  unknown: ChunaRow[];
}

async function enrichWithPayrollCode(rows: TeacherMonthlyFeeSummaryRow[]): Promise<ChunaRow[]> {
  const ids = rows.map((r) => r.substituteTeacherId).filter((id): id is string => !!id);
  const persons = ids.length > 0 ? await prisma.person.findMany({ where: { id: { in: ids } } }) : [];
  const codeById = new Map(persons.map((p) => [p.id, p.payrollCode]));
  return rows
    .map((r) => ({ ...r, payrollCode: r.substituteTeacherId ? codeById.get(r.substituteTeacherId) ?? null : null }))
    // 目前真實給出納檔案看不出明確排序規則（不是純姓名排序也不是純節數排序），
    // 這裡採用「節數由多到少」當預設，貼近真實檔案大部分列的呈現方式。
    .sort((a, b) => b.totalCount - a.totalCount);
}

export async function getChunaSummary(monthlyImportIds: string[]): Promise<ChunaSummary> {
  // 自費代課的虛擬容器批次不算「給出納」的一部分——給出納是公費核銷用的名冊，
  // 自費代課是學校/單位自行吸收的支出，兩者不應該混在同一張表或同一個總額裡
  // （跟 dashboardService 排除自費代課批次算公費總額是同一個原則）。
  const excelImports = await prisma.monthlyImport.findMany({
    where: { id: { in: monthlyImportIds }, fileName: { not: SELF_FUNDED_BATCH_FILE_NAME } },
  });
  const bdIds = excelImports.filter((i) => i.sourceStaffType === "BD").map((i) => i.id);
  const nonBdIds = excelImports.filter((i) => i.sourceStaffType === "NON_BD").map((i) => i.id);
  const unknownIds = excelImports.filter((i) => i.sourceStaffType === "UNKNOWN").map((i) => i.id);

  const [bd, nonBd, unknown] = await Promise.all([
    bdIds.length > 0 ? summarizeTeacherMonthlyFees(bdIds) : Promise.resolve([]),
    nonBdIds.length > 0 ? summarizeTeacherMonthlyFees(nonBdIds) : Promise.resolve([]),
    unknownIds.length > 0 ? summarizeTeacherMonthlyFees(unknownIds) : Promise.resolve([]),
  ]);

  return {
    bd: await enrichWithPayrollCode(bd),
    nonBd: await enrichWithPayrollCode(nonBd),
    unknown: await enrichWithPayrollCode(unknown),
  };
}

function writeSection(ws: ExcelJS.Worksheet, title: string, rows: ChunaRow[]) {
  ws.addRow([title]);
  ws.addRow(["代課教師代碼", "代課教師", "節數", "金額", "備註"]);
  let totalCount = 0;
  let totalAmount = 0;
  for (const r of rows) {
    ws.addRow([r.payrollCode ?? "", r.substituteTeacherName, r.totalCount, Number(r.totalAmount), ""]);
    totalCount += r.totalCount;
    totalAmount += Number(r.totalAmount);
  }
  ws.addRow(["", "總計", totalCount, totalAmount, ""]);
  ws.addRow([]);
}

export async function generateChunaExcelBuffer(monthlyImportIds: string[], year: number, month: number): Promise<Buffer> {
  const summary = await getChunaSummary(monthlyImportIds);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(`${year}年${month}月給出納`);

  if (summary.nonBd.length > 0) writeSection(ws, `${year}年${month}月公費代課（編制外）`, summary.nonBd);
  if (summary.bd.length > 0) writeSection(ws, `${year}年${month}月公費代課（編制內）`, summary.bd);
  if (summary.unknown.length > 0) writeSection(ws, `${year}年${month}月公費代課（未標示來源）`, summary.unknown);

  ws.columns.forEach((col) => {
    col.width = 18;
  });

  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}
