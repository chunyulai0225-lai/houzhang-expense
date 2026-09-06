// Implementation Batch：對帳
//
// 讀取使用者上傳的原始「給出納」Excel，逐教師跟系統算出來的結果比對。
// 不會只比總金額；每個差異都嘗試從資料本身找「可能原因」，但不會宣稱「已確認原因」——
// 能不能推導出原因，跟原因是不是真的正確，是兩件事，這裡只做前者。
import ExcelJS from "exceljs";
import { prisma } from "../prismaClient";
import { getChunaSummary, ChunaRow } from "./chunaService";

function cellToString(cell: ExcelJS.Cell): string {
  const v = cell.value as any;
  if (v == null) return "";
  if (typeof v === "object") {
    if (v.richText) return v.richText.map((r: any) => r.text).join("");
    if (v.text) return v.text;
    if (v.result != null) return String(v.result);
  }
  return String(v);
}

interface UploadedChunaRow {
  name: string;
  periodCount: number;
  amount: number;
}

// 掃描所有工作表，找出「代課教師」或「教師」這種姓名欄，搭配「節數」「金額／總額」欄，
// 自動找出對應的欄位位置，不假設固定在第幾欄——真實檔案的欄位順序不保證每次都一樣。
export function parseUploadedChunaWorkbook(buffer: Buffer): Promise<UploadedChunaRow[]> {
  const wb = new ExcelJS.Workbook();
  return wb.xlsx.load(buffer as any).then(() => {
    const rows: UploadedChunaRow[] = [];
    for (const ws of wb.worksheets) {
      let nameCol = -1;
      let periodCol = -1;
      let amountCol = -1;
      let headerRow = -1;

      const maxScanRow = Math.min(ws.rowCount, 15);
      for (let r = 1; r <= maxScanRow; r++) {
        const row = ws.getRow(r);
        let foundName = -1;
        let foundPeriod = -1;
        let foundAmount = -1;
        for (let c = 1; c <= Math.min(ws.columnCount, 30); c++) {
          const text = cellToString(row.getCell(c)).replace(/\s|\n/g, "");
          if (!text) continue;
          if (text.includes("教師") && !text.includes("代碼") && foundName === -1) foundName = c;
          if (text.includes("節數") && foundPeriod === -1) foundPeriod = c;
          if ((text.includes("總額") || text.includes("金額")) && foundAmount === -1) foundAmount = c;
        }
        if (foundName !== -1 && foundPeriod !== -1 && foundAmount !== -1) {
          nameCol = foundName;
          periodCol = foundPeriod;
          amountCol = foundAmount;
          headerRow = r;
          break;
        }
      }

      if (headerRow === -1) continue; // 這個工作表看起來不是給出納格式，略過

      for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const name = cellToString(row.getCell(nameCol)).trim();
        const periodText = cellToString(row.getCell(periodCol)).trim();
        const amountText = cellToString(row.getCell(amountCol)).trim();
        if (!name || name.includes("總計") || name.includes("小計")) continue;
        const periodCount = Number(periodText);
        const amount = Number(amountText);
        if (Number.isNaN(periodCount) || Number.isNaN(amount)) continue;
        rows.push({ name, periodCount, amount });
      }
    }
    return rows;
  });
}

export type ReconciliationStatus = "MATCH" | "SYSTEM_LESS" | "SYSTEM_MORE" | "ONLY_SYSTEM" | "ONLY_ORIGINAL" | "UNCERTAIN";

export interface ReconciliationRow {
  name: string;
  systemPeriodCount: number | null;
  originalPeriodCount: number | null;
  systemAmount: number | null;
  originalAmount: number | null;
  amountDiff: number | null;
  status: ReconciliationStatus;
  possibleReason: string | null; // 只是「可能原因」，不是「已確認原因」
}

export interface ReconciliationResult {
  rows: ReconciliationRow[];
  totals: {
    systemAmount: number;
    originalAmount: number;
    diff: number;
  };
}

async function guessPossibleReason(name: string, monthlyImportIds: string[], diffDirection: "less" | "more" | "onlyOriginal"): Promise<string | null> {
  if (diffDirection === "less" || diffDirection === "onlyOriginal") {
    // 系統比原始少：檢查這個名字是不是曾經以「代課教師」身分出現在匯入錯誤的原始資料裡
    // （代表原始檔案裡有這個人的紀錄，但因為格式問題沒有成功變成 SubstituteRecord）。
    const rawWithError = await prisma.substituteRecordRaw.findFirst({
      where: { monthlyImportId: { in: monthlyImportIds }, substituteTeacherText: name, processedRecord: null },
    });
    if (rawWithError) {
      const count = await prisma.substituteRecordRaw.count({
        where: { monthlyImportId: { in: monthlyImportIds }, substituteTeacherText: name, processedRecord: null },
      });
      return `可能原因：有 ${count} 筆該教師的原始資料因為日期／節次格式問題未能成功匯入（可在匯入錯誤列表查證）`;
    }
    return "尚無法從資料自動判斷原因，請人工確認（例如：原始給出納可能包含固定課表／整月代理等系統目前資料來源沒有涵蓋的項目）";
  }
  if (diffDirection === "more") {
    return "可能原因：系統包含原始給出納未列出的紀錄，或原始給出納對這筆金額另有特殊處理（例如超鐘點另計）——請人工確認";
  }
  return null;
}

export async function reconcile(monthlyImportIds: string[], uploadedRows: UploadedChunaRow[]): Promise<ReconciliationResult> {
  const summary = await getChunaSummary(monthlyImportIds);
  const allSystemRows: ChunaRow[] = [...summary.bd, ...summary.nonBd, ...summary.unknown];
  // 同一個姓名可能同時出現在 BD／非BD／未標示來源這幾個分組（例如同一位教師這個月
  // 同時有編制內跟編制外的代課紀錄）——這裡要把節數／金額加總起來，絕對不能讓後面
  // 的分組直接蓋掉前面的，否則系統總額會悄悄少算，對帳就失去意義了。
  const systemByName = new Map<string, ChunaRow>();
  for (const r of allSystemRows) {
    const key = r.substituteTeacherName.replace("（未配對）", "").trim();
    const existing = systemByName.get(key);
    if (existing) {
      systemByName.set(key, {
        ...existing,
        generalCount: existing.generalCount + r.generalCount,
        generalAmount: (Number(existing.generalAmount) + Number(r.generalAmount)).toString(),
        overtimeCount: existing.overtimeCount + r.overtimeCount,
        overtimeAmount: (Number(existing.overtimeAmount) + Number(r.overtimeAmount)).toString(),
        projectCount: existing.projectCount + r.projectCount,
        projectAmount: (Number(existing.projectAmount) + Number(r.projectAmount)).toString(),
        totalCount: existing.totalCount + r.totalCount,
        totalAmount: (Number(existing.totalAmount) + Number(r.totalAmount)).toString(),
      });
    } else {
      systemByName.set(key, r);
    }
  }
  const uploadedByName = new Map(uploadedRows.map((r) => [r.name.trim(), r]));

  const allNames = new Set([...systemByName.keys(), ...uploadedByName.keys()]);
  const rows: ReconciliationRow[] = [];
  let totalSystem = 0;
  let totalOriginal = 0;

  for (const name of allNames) {
    const sys = systemByName.get(name);
    const orig = uploadedByName.get(name);
    const systemAmount = sys ? Number(sys.totalAmount) : null;
    const originalAmount = orig ? orig.amount : null;
    totalSystem += systemAmount ?? 0;
    totalOriginal += originalAmount ?? 0;

    let status: ReconciliationStatus;
    let possibleReason: string | null = null;
    const amountDiff = systemAmount !== null && originalAmount !== null ? systemAmount - originalAmount : null;

    if (sys && !orig) {
      status = "ONLY_SYSTEM";
    } else if (!sys && orig) {
      status = "ONLY_ORIGINAL";
      possibleReason = await guessPossibleReason(name, monthlyImportIds, "onlyOriginal");
    } else if (amountDiff === 0) {
      status = "MATCH";
    } else if (amountDiff !== null && amountDiff < 0) {
      status = "SYSTEM_LESS";
      possibleReason = await guessPossibleReason(name, monthlyImportIds, "less");
    } else if (amountDiff !== null && amountDiff > 0) {
      status = "SYSTEM_MORE";
      possibleReason = await guessPossibleReason(name, monthlyImportIds, "more");
    } else {
      status = "UNCERTAIN";
    }

    rows.push({
      name,
      systemPeriodCount: sys?.totalCount ?? null,
      originalPeriodCount: orig?.periodCount ?? null,
      systemAmount,
      originalAmount,
      amountDiff,
      status,
      possibleReason,
    });
  }

  rows.sort((a, b) => Math.abs(b.amountDiff ?? 0) - Math.abs(a.amountDiff ?? 0));

  return {
    rows,
    totals: { systemAmount: totalSystem, originalAmount: totalOriginal, diff: totalSystem - totalOriginal },
  };
}
