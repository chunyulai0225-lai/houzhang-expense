// Phase 7：公費代課 Excel 匯入
//
// 流程：Excel → Raw Data（永不覆寫）→ 標準化代課資料（SubstituteRecord）。
// 刻意不做的事：一般公費／超鐘點／專案自動分類、任何費率計算（405元等）、
// 代導師費／日薪／半日薪、專案扣除計算——這些全部留給 Phase 8。
// 這裡的 SubstituteRecord 只是「已標準化但尚未分類」的資料
// （fundingSource=UNDETERMINED、staffType=UNKNOWN、classificationMethod=GENERAL_DEFAULT 皆為預設值）。

import ExcelJS from "exceljs";
import type { MonthlyImport, Person, Weekday } from "@prisma/client";
import { prisma } from "../prismaClient";

// ============================================================
// 一、Excel 解析：欄位別名對照（尚未取得真實 Excel，先依規格書列出的欄位名稱比對，
//     真實檔案格式若有出入，只需要調整這份對照表，不需要更動匯入流程本身）
// ============================================================

const COLUMN_ALIASES: Record<string, string[]> = {
  rowNumber: ["序", "序號"],
  originalTeacherCode: ["原教師代碼", "原教師代号"],
  originalTeacherName: ["原教師"],
  dateText: ["日期"],
  leaveType: ["假別"],
  hoursOrDaysText: ["時數天數", "時數/天數", "時數／天數"],
  periodText: ["節次"],
  className: ["班級"],
  subject: ["科目"],
  substituteTeacherCode: ["代課教師代碼", "代課教師代号"],
  substituteTeacherName: ["代課教師"],
  teacherCert: ["教師證"],
  payGrade: ["薪等"],
  homeroomFeeText: ["代導師"],
};

// 缺少這些欄位標題時，整份 Excel 直接視為結構異常，不會建立任何批次
const REQUIRED_FIELDS = ["dateText", "originalTeacherName", "substituteTeacherName", "periodText"] as const;

export interface ParsedExcelRow {
  rowNumber: number;
  raw: Record<string, string | null>;
  originalTeacherCode?: string;
  originalTeacherName?: string;
  dateText?: string;
  leaveType?: string;
  hoursOrDaysText?: string;
  periodText?: string;
  className?: string;
  subject?: string;
  substituteTeacherCode?: string;
  substituteTeacherName?: string;
  teacherCert?: string;
  payGrade?: string;
  homeroomFeeText?: string;
}

export interface ParsedWorkbook {
  headers: string[];
  rows: ParsedExcelRow[];
}

function cellToString(value: ExcelJS.CellValue): string | null {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "object") {
    const v = value as { text?: unknown; result?: unknown };
    if ("text" in v && v.text != null) return String(v.text).trim() || null;
    if ("result" in v && v.result != null) return String(v.result).trim() || null;
    return null;
  }
  const s = String(value).trim();
  return s === "" ? null : s;
}

function isRowEffectivelyEmpty(row: ParsedExcelRow): boolean {
  return !(
    row.dateText ||
    row.originalTeacherName ||
    row.substituteTeacherName ||
    row.periodText ||
    row.className ||
    row.subject
  );
}

export async function parseWorkbook(buffer: Buffer): Promise<ParsedWorkbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as unknown as ExcelJS.Buffer);
  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new Error("Excel 檔案沒有任何工作表");
  }

  const headerRow = sheet.getRow(1);
  const headers: string[] = [];
  const columnMap: Record<string, number> = {};
  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const text = cellToString(cell.value);
    if (!text) return;
    headers.push(text);
    for (const [field, aliases] of Object.entries(COLUMN_ALIASES)) {
      if (aliases.includes(text)) {
        columnMap[field] = colNumber;
      }
    }
  });

  const missingRequired = REQUIRED_FIELDS.filter((f) => !(f in columnMap));
  if (missingRequired.length > 0) {
    const labels = missingRequired.map((f) => COLUMN_ALIASES[f][0]).join("、");
    throw new Error(`Excel 缺少必要欄位：${labels}（找到的欄位：${headers.join("、") || "無"}）`);
  }

  const rows: ParsedExcelRow[] = [];
  const lastRowNumber = sheet.lastRow?.number ?? 1;
  const get = (row: ExcelJS.Row, field: string): string | undefined => {
    const col = columnMap[field];
    if (!col) return undefined;
    return cellToString(row.getCell(col).value) ?? undefined;
  };

  for (let r = 2; r <= lastRowNumber; r++) {
    const row = sheet.getRow(r);
    const raw: Record<string, string | null> = {};
    headerRow.eachCell({ includeEmpty: false }, (headerCell, colNumber) => {
      const headerText = cellToString(headerCell.value);
      if (!headerText) return;
      raw[headerText] = cellToString(row.getCell(colNumber).value);
    });

    const parsed: ParsedExcelRow = {
      rowNumber: r - 1,
      raw,
      originalTeacherCode: get(row, "originalTeacherCode"),
      originalTeacherName: get(row, "originalTeacherName"),
      dateText: get(row, "dateText"),
      leaveType: get(row, "leaveType"),
      hoursOrDaysText: get(row, "hoursOrDaysText"),
      periodText: get(row, "periodText"),
      className: get(row, "className"),
      subject: get(row, "subject"),
      substituteTeacherCode: get(row, "substituteTeacherCode"),
      substituteTeacherName: get(row, "substituteTeacherName"),
      teacherCert: get(row, "teacherCert"),
      payGrade: get(row, "payGrade"),
      homeroomFeeText: get(row, "homeroomFeeText"),
    };

    if (isRowEffectivelyEmpty(parsed)) continue; // 跳過完全空白的間隔列
    rows.push(parsed);
  }

  return { headers, rows };
}

// ============================================================
// 二、日期／節次標準化（只做格式標準化，不做分類判斷）
// ============================================================

const WEEKDAY_BY_JS_INDEX: Weekday[] = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"];
const WEEKDAY_CHAR_MAP: Record<string, Weekday> = {
  一: "MON",
  二: "TUE",
  三: "WED",
  四: "THU",
  五: "FRI",
  六: "SAT",
  日: "SUN",
  天: "SUN",
};

export interface ParsedDateResult {
  date: Date;
  weekday: Weekday;
  weekdayMismatch: boolean;
  monthMismatch: boolean;
}

// 年份／月份由匯入時管理者選擇的學期＋年月提供（例如「2026年6月」），
// 不從 Excel 本身推斷，因為原始文字通常只有「MM-DD」，沒有完整年份。
export function parseDateText(
  text: string,
  year: number,
  expectedMonth: number
): ParsedDateResult | { error: string } {
  const match = text.trim().match(/^(\d{1,2})-(\d{1,2})(?:\(([一二三四五六日天])\))?/);
  if (!match) {
    return { error: `無法解析日期格式："${text}"，預期格式如 "06-18(四)"` };
  }
  const month = Number(match[1]);
  const day = Number(match[2]);
  if (month < 1 || month > 12) {
    return { error: `日期月份不合理："${text}"` };
  }
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    return { error: `日期不存在："${text}"` };
  }
  const computedWeekday = WEEKDAY_BY_JS_INDEX[date.getUTCDay()];
  const textWeekdayChar = match[3];
  const weekdayMismatch = Boolean(textWeekdayChar) && WEEKDAY_CHAR_MAP[textWeekdayChar!] !== computedWeekday;
  const monthMismatch = month !== expectedMonth;
  return { date, weekday: computedWeekday, weekdayMismatch, monthMismatch };
}

const CHINESE_DIGIT_MAP: Record<string, number> = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7 };

// validCodes 由呼叫端傳入目前系統設定的 PeriodSlot.code 清單，
// 確保產生的節次代碼一定是系統實際存在、可設定的值，不是憑空寫死。
export function parsePeriodText(text: string, validCodes: Set<string>): { periodCode: string } | { error: string } {
  const trimmed = text.trim();

  const check = (code: string): { periodCode: string } | { error: string } =>
    validCodes.has(code) ? { periodCode: code } : { error: `節次代碼「${code}」尚未設定於系統（原始文字："${text}"）` };

  if (trimmed === "早自修") return check("EARLY_STUDY");
  if (trimmed === "午休") return check("LUNCH");

  const match = trimmed.match(/^第([0-9一二三四五六七]+)節$/);
  if (match) {
    const raw = match[1];
    const num = /^[0-9]+$/.test(raw) ? Number(raw) : CHINESE_DIGIT_MAP[raw];
    if (num) return check(`P${num}`);
  }

  // 多節／範圍等複合格式（例如「第1,2節」）目前不猜測拆分方式，一律視為無法解析
  return { error: `無法解析節次："${text}"` };
}

// ============================================================
// 三、匯入批次（含重新匯入策略）
// ============================================================

async function computeNextVersionNo(semesterId: string, year: number, month: number): Promise<number> {
  const last = await prisma.monthlyImport.findFirst({
    where: { semesterId, year, month },
    orderBy: { versionNo: "desc" },
  });
  return (last?.versionNo ?? 0) + 1;
}

export interface ImportRowIssue {
  rowNumber: number;
  fieldName?: string;
  message: string;
}

export interface ImportSubstituteExcelResult {
  monthlyImport: MonthlyImport;
  detectedHeaders: string[];
  totalCount: number;
  successCount: number;
  errorCount: number;
  errors: ImportRowIssue[];
  supersededImportIds: string[];
}

export async function importSubstituteExcel(params: {
  semesterId: string;
  year: number;
  month: number;
  fileName: string;
  fileBuffer: Buffer;
  importedBy?: string;
}): Promise<ImportSubstituteExcelResult> {
  await prisma.semester.findUniqueOrThrow({ where: { id: params.semesterId } });

  const { headers, rows } = await parseWorkbook(params.fileBuffer);

  const periodSlots = await prisma.periodSlot.findMany({ select: { code: true } });
  const validPeriodCodes = new Set(periodSlots.map((p) => p.code));

  // 重新匯入策略：同學期＋同年月的舊「有效」批次全部轉為 SUPERSEDED，不刪除，
  // 新批次的 versionNo 遞增，成為新的 ACTIVE 批次。
  const previousActive = await prisma.monthlyImport.findMany({
    where: { semesterId: params.semesterId, year: params.year, month: params.month, status: "ACTIVE" },
  });
  const nextVersionNo = await computeNextVersionNo(params.semesterId, params.year, params.month);

  const monthlyImport = await prisma.monthlyImport.create({
    data: {
      semesterId: params.semesterId,
      year: params.year,
      month: params.month,
      fileName: params.fileName,
      importedBy: params.importedBy,
      totalCount: rows.length,
      versionNo: nextVersionNo,
      status: "ACTIVE",
    },
  });

  if (previousActive.length > 0) {
    await prisma.monthlyImport.updateMany({
      where: { id: { in: previousActive.map((b) => b.id) } },
      data: { status: "SUPERSEDED" },
    });
    for (const b of previousActive) {
      await prisma.changeLog.create({
        data: {
          tableName: "monthly_imports",
          recordId: b.id,
          fieldName: "status",
          oldValue: "ACTIVE",
          newValue: "SUPERSEDED",
          changedBy: params.importedBy,
          reason: `重新匯入，被新批次取代（${monthlyImport.id}）`,
        },
      });
    }
  }

  const issues: ImportRowIssue[] = [];
  let successCount = 0;

  for (const row of rows) {
    // 不論後面解析成不成功，原始資料一律先完整保存
    const raw = await prisma.substituteRecordRaw.create({
      data: {
        monthlyImportId: monthlyImport.id,
        rowNumber: row.rowNumber,
        originalTeacherCodeText: row.originalTeacherCode,
        originalTeacherText: row.originalTeacherName,
        dateText: row.dateText,
        leaveTypeText: row.leaveType,
        hoursOrDaysText: row.hoursOrDaysText,
        periodText: row.periodText,
        classText: row.className,
        subjectText: row.subject,
        substituteTeacherCodeText: row.substituteTeacherCode,
        substituteTeacherText: row.substituteTeacherName,
        teacherCertText: row.teacherCert,
        payGradeText: row.payGrade,
        homeroomFeeText: row.homeroomFeeText,
        rawJson: JSON.stringify(row.raw),
      },
    });

    const rowIssues: ImportRowIssue[] = [];
    if (!row.dateText) rowIssues.push({ rowNumber: row.rowNumber, fieldName: "日期", message: "缺少日期" });
    if (!row.originalTeacherName) rowIssues.push({ rowNumber: row.rowNumber, fieldName: "原教師", message: "缺少原教師" });
    if (!row.substituteTeacherName)
      rowIssues.push({ rowNumber: row.rowNumber, fieldName: "代課教師", message: "缺少代課教師" });
    if (!row.periodText) rowIssues.push({ rowNumber: row.rowNumber, fieldName: "節次", message: "缺少節次" });

    let parsedDate: ParsedDateResult | null = null;
    if (row.dateText) {
      const result = parseDateText(row.dateText, params.year, params.month);
      if ("error" in result) {
        rowIssues.push({ rowNumber: row.rowNumber, fieldName: "日期", message: result.error });
      } else {
        parsedDate = result;
      }
    }

    let periodCode: string | null = null;
    if (row.periodText) {
      const result = parsePeriodText(row.periodText, validPeriodCodes);
      if ("error" in result) {
        rowIssues.push({ rowNumber: row.rowNumber, fieldName: "節次", message: result.error });
      } else {
        periodCode = result.periodCode;
      }
    }

    if (rowIssues.length > 0 || !parsedDate || !periodCode) {
      issues.push(...rowIssues);
      continue; // 不建立 SubstituteRecord，只留下 raw 供追查
    }

    const noteParts: string[] = [];
    if (parsedDate.weekdayMismatch) {
      noteParts.push(`日期文字標示的星期與實際計算不符（原始："${row.dateText}"）`);
    }
    if (parsedDate.monthMismatch) {
      noteParts.push(`此列日期月份與所選匯入月份（${params.month}月）不同，請確認`);
    }

    await prisma.substituteRecord.create({
      data: {
        rawRecordId: raw.id,
        monthlyImportId: monthlyImport.id,
        date: parsedDate.date,
        weekday: parsedDate.weekday,
        periodCode,
        className: row.className,
        subject: row.subject,
        leaveType: row.leaveType,
        rawHoursOrDays: row.hoursOrDaysText,
        note: noteParts.length > 0 ? noteParts.join("；") : undefined,
      },
    });
    successCount += 1;
  }

  if (issues.length > 0) {
    await prisma.monthlyImportError.createMany({
      data: issues.map((i) => ({
        monthlyImportId: monthlyImport.id,
        rowNumber: i.rowNumber,
        fieldName: i.fieldName,
        message: i.message,
      })),
    });
  }

  const updated = await prisma.monthlyImport.update({
    where: { id: monthlyImport.id },
    data: { successCount, errorCount: issues.length },
  });

  return {
    monthlyImport: updated,
    detectedHeaders: headers,
    totalCount: rows.length,
    successCount,
    errorCount: issues.length,
    errors: issues,
    supersededImportIds: previousActive.map((b) => b.id),
  };
}

// ============================================================
// 四、查詢
// ============================================================

export async function listMonthlyImports(semesterId: string, filter: { year?: number; month?: number } = {}) {
  return prisma.monthlyImport.findMany({
    where: {
      semesterId,
      ...(filter.year ? { year: filter.year } : {}),
      ...(filter.month ? { month: filter.month } : {}),
    },
    orderBy: [{ year: "desc" }, { month: "desc" }, { versionNo: "desc" }],
  });
}

export async function getMonthlyImportDetail(monthlyImportId: string) {
  return prisma.monthlyImport.findUniqueOrThrow({
    where: { id: monthlyImportId },
    include: { errors: { orderBy: { rowNumber: "asc" } } },
  });
}

export async function listSubstituteRecords(monthlyImportId: string) {
  return prisma.substituteRecord.findMany({
    where: { monthlyImportId },
    include: { rawRecord: true, originalTeacher: true, substituteTeacher: true },
    orderBy: { date: "asc" },
  });
}

// ============================================================
// 五、教師對應（不自動合併、不偷偷建新人；只提供「建議」，須管理者明確確認）
// ============================================================

export interface UnmatchedTeacherRef {
  recordId: string;
  field: "original" | "substitute";
  rawName: string;
  candidates: Person[];
}

export async function listUnmatchedTeacherReferences(monthlyImportId: string): Promise<UnmatchedTeacherRef[]> {
  const records = await prisma.substituteRecord.findMany({
    where: { monthlyImportId, OR: [{ originalTeacherId: null }, { substituteTeacherId: null }] },
    include: { rawRecord: true },
  });

  const results: UnmatchedTeacherRef[] = [];
  for (const r of records) {
    if (!r.originalTeacherId && r.rawRecord.originalTeacherText) {
      const candidates = await prisma.person.findMany({ where: { name: r.rawRecord.originalTeacherText } });
      results.push({ recordId: r.id, field: "original", rawName: r.rawRecord.originalTeacherText, candidates });
    }
    if (!r.substituteTeacherId && r.rawRecord.substituteTeacherText) {
      const candidates = await prisma.person.findMany({ where: { name: r.rawRecord.substituteTeacherText } });
      results.push({ recordId: r.id, field: "substitute", rawName: r.rawRecord.substituteTeacherText, candidates });
    }
  }
  return results;
}

export async function resolveTeacherReference(
  recordId: string,
  field: "original" | "substitute",
  personId: string,
  changedBy?: string,
  reason = "人工配對教師"
) {
  const record = await prisma.substituteRecord.findUniqueOrThrow({ where: { id: recordId } });
  const fieldName = field === "original" ? "originalTeacherId" : "substituteTeacherId";
  const oldValue = field === "original" ? record.originalTeacherId : record.substituteTeacherId;

  const updated = await prisma.substituteRecord.update({
    where: { id: recordId },
    data: field === "original" ? { originalTeacherId: personId } : { substituteTeacherId: personId },
  });

  await prisma.changeLog.create({
    data: { tableName: "substitute_records", recordId, fieldName, oldValue, newValue: personId, changedBy, reason },
  });

  return updated;
}

// 明確的一次性「套用」動作（不是匯入當下自動發生）：只處理姓名剛好唯一相符的情況，
// 讓大量常見案例不用逐筆手動點選，但仍然是管理者主動觸發、可追查的一個動作。
export async function autoApplyUnambiguousTeacherMatches(
  monthlyImportId: string,
  changedBy?: string
): Promise<{ appliedCount: number; remainingCount: number }> {
  const unmatched = await listUnmatchedTeacherReferences(monthlyImportId);
  let applied = 0;
  for (const u of unmatched) {
    if (u.candidates.length === 1) {
      await resolveTeacherReference(
        u.recordId,
        u.field,
        u.candidates[0].id,
        changedBy,
        "系統依姓名找到唯一相符人員，經管理者一鍵套用"
      );
      applied += 1;
    }
  }
  const remaining = await listUnmatchedTeacherReferences(monthlyImportId);
  return { appliedCount: applied, remainingCount: remaining.length };
}
