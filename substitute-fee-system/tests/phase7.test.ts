// Phase 7 驗收測試：公費代課 Excel 匯入
import ExcelJS from "exceljs";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { prisma } from "../src/prismaClient";
import { createPerson } from "../src/services/personService";
import {
  autoApplyUnambiguousTeacherMatches,
  importSubstituteExcel,
  listMonthlyImports,
  listUnmatchedTeacherReferences,
  listWorkbookSheets,
  parseDateText,
  parsePeriodText,
  resolveTeacherReference,
} from "../src/services/excelImportService";

const STANDARD_HEADERS = [
  "序",
  "原教師代碼",
  "原教師",
  "日期",
  "假別",
  "時數天數",
  "節次",
  "班級",
  "科目",
  "代課教師代碼",
  "代課教師",
  "教師證",
  "薪等",
  "代導師",
];

async function buildWorkbookBuffer(headers: string[], rows: (string | number | null)[][]): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Sheet1");
  sheet.addRow(headers);
  for (const row of rows) sheet.addRow(row);
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

function standardRow(overrides: Partial<Record<(typeof STANDARD_HEADERS)[number], string | number | null>> = {}) {
  const defaults: Record<string, string | number | null> = {
    序: 1,
    原教師代碼: "T001",
    原教師: "陳心啓",
    日期: "09-01(二)",
    假別: "公假",
    時數天數: "2",
    節次: "第2節",
    班級: "五年甲班",
    科目: "自然",
    代課教師代碼: "T002",
    代課教師: "徐碧苓",
    教師證: "有",
    薪等: "D",
    代導師: "",
  };
  const merged = { ...defaults, ...overrides };
  return STANDARD_HEADERS.map((h) => merged[h] ?? null);
}

const cleanupPersonIds: string[] = [];
const cleanupSemesterIds: string[] = [];

async function makeTestSemester(schoolYear: number) {
  const sem = await prisma.semester.create({
    data: { schoolYear, term: 1, startDate: new Date("2026-08-01"), endDate: new Date("2027-01-31") },
  });
  cleanupSemesterIds.push(sem.id);
  return sem;
}

beforeAll(async () => {
  await prisma.$connect();
});

afterAll(async () => {
  const imports = await prisma.monthlyImport.findMany({ where: { semesterId: { in: cleanupSemesterIds } } });
  const importIds = imports.map((i) => i.id);
  await prisma.changeLog.deleteMany({
    where: { OR: [{ recordId: { in: importIds } }, { recordId: { in: cleanupPersonIds } }] },
  });
  await prisma.monthlyImportError.deleteMany({ where: { monthlyImportId: { in: importIds } } });
  await prisma.substituteRecord.deleteMany({ where: { monthlyImportId: { in: importIds } } });
  await prisma.substituteRecordRaw.deleteMany({ where: { monthlyImportId: { in: importIds } } });
  await prisma.monthlyImport.deleteMany({ where: { id: { in: importIds } } });
  await prisma.person.deleteMany({ where: { id: { in: cleanupPersonIds } } });
  await prisma.semester.deleteMany({ where: { id: { in: cleanupSemesterIds } } });
  await prisma.$disconnect();
});

describe("1/2. 正常匯入與欄位解析", () => {
  it("可以匯入一份格式正確的 Excel，並回報偵測到的欄位", async () => {
    const sem = await makeTestSemester(601);
    const buffer = await buildWorkbookBuffer(STANDARD_HEADERS, [standardRow()]);

    const result = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "測試代課明細.xlsx",
      fileBuffer: buffer,
      importedBy: "測試管理者",
    });

    expect(result.totalCount).toBe(1);
    expect(result.successCount).toBe(1);
    expect(result.errorCount).toBe(0);
    expect(result.detectedHeaders).toEqual(STANDARD_HEADERS);
    expect(result.monthlyImport.status).toBe("ACTIVE");
  });
});

describe("3. 日期解析", () => {
  it("可以把 MM-DD(星期) 解析成正確日期，並偵測星期不符", () => {
    const result = parseDateText("06-18(四)", 2026, 6);
    expect("error" in result).toBe(false);
    if ("error" in result) return;
    expect(result.date.toISOString().slice(0, 10)).toBe("2026-06-18");
    const actualWeekdayIndex = new Date("2026-06-18T00:00:00.000Z").getUTCDay();
    const actualWeekday = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"][actualWeekdayIndex];
    expect(result.weekday).toBe(actualWeekday);
    expect(result.monthMismatch).toBe(false);
  });

  it("星期文字與實際計算不符時會標示 weekdayMismatch", () => {
    // 2026-06-18 實際星期由程式計算，這裡故意填一個保證不符的星期字
    const actualWeekdayIndex = new Date("2026-06-18T00:00:00.000Z").getUTCDay();
    const chars = ["日", "一", "二", "三", "四", "五", "六"];
    const wrongChar = chars[(actualWeekdayIndex + 1) % 7];
    const result = parseDateText(`06-18(${wrongChar})`, 2026, 6);
    expect("error" in result).toBe(false);
    if ("error" in result) return;
    expect(result.weekdayMismatch).toBe(true);
  });

  it("月份與所選匯入月份不同時會標示 monthMismatch", () => {
    const result = parseDateText("07-02(四)", 2026, 6);
    if ("error" in result) throw new Error("unexpected error");
    expect(result.monthMismatch).toBe(true);
  });

  it("無法解析的日期格式會回傳錯誤", () => {
    const result = parseDateText("這不是日期", 2026, 6);
    expect("error" in result).toBe(true);
  });
});

describe("4. 節次解析", () => {
  const validCodes = new Set(["EARLY_STUDY", "P1", "P2", "P3", "P4", "LUNCH", "P5", "P6", "P7"]);

  it("阿拉伯數字、中文數字、早自修/午休都能正確解析", () => {
    expect(parsePeriodText("第1節", validCodes)).toEqual({ periodCode: "P1" });
    expect(parsePeriodText("第五節", validCodes)).toEqual({ periodCode: "P5" });
    expect(parsePeriodText("早自修", validCodes)).toEqual({ periodCode: "EARLY_STUDY" });
    expect(parsePeriodText("午休", validCodes)).toEqual({ periodCode: "LUNCH" });
  });

  it("多節等複合格式無法解析時回傳錯誤，不猜測拆分方式", () => {
    const result = parsePeriodText("第1,2節", validCodes);
    expect("error" in result).toBe(true);
  });

  it("解析出的代碼若不在系統設定的節次清單中，也視為錯誤", () => {
    const result = parsePeriodText("第9節", new Set(["P1", "P2"]));
    expect("error" in result).toBe(true);
  });
});

describe("5. 原始資料完整保存", () => {
  it("匯入後原始欄位（含教師代碼、代導師等）與原始 JSON 都完整保留", async () => {
    const sem = await makeTestSemester(602);
    const buffer = await buildWorkbookBuffer(STANDARD_HEADERS, [
      standardRow({ 代導師: "133", 時數天數: "1天" }),
    ]);
    const result = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "test.xlsx",
      fileBuffer: buffer,
    });

    const raw = await prisma.substituteRecordRaw.findFirstOrThrow({
      where: { monthlyImportId: result.monthlyImport.id },
    });
    expect(raw.originalTeacherCodeText).toBe("T001");
    expect(raw.originalTeacherText).toBe("陳心啓");
    expect(raw.substituteTeacherCodeText).toBe("T002");
    expect(raw.substituteTeacherText).toBe("徐碧苓");
    expect(raw.homeroomFeeText).toBe("133");
    expect(raw.hoursOrDaysText).toBe("1天");
    const rawJson = JSON.parse(raw.rawJson!);
    expect(rawJson["原教師"]).toBe("陳心啓");
    expect(rawJson["代導師"]).toBe("133");
  });
});

describe("6/7. 教師比對", () => {
  it("姓名唯一相符時，套用建議後可成功配對", async () => {
    const sem = await makeTestSemester(603);
    const createdOriginal = await createPerson({ name: "測試原教師甲" });
    const createdSub = await createPerson({ name: "測試代課教師甲" });
    if (createdOriginal.status !== "CREATED" || createdSub.status !== "CREATED") throw new Error("unreachable");
    cleanupPersonIds.push(createdOriginal.person.id, createdSub.person.id);

    const buffer = await buildWorkbookBuffer(STANDARD_HEADERS, [
      standardRow({ 原教師: "測試原教師甲", 代課教師: "測試代課教師甲" }),
    ]);
    const result = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "test.xlsx",
      fileBuffer: buffer,
    });

    // 匯入當下不會自動配對
    const beforeApply = await prisma.substituteRecord.findFirstOrThrow({
      where: { monthlyImportId: result.monthlyImport.id },
    });
    expect(beforeApply.originalTeacherId).toBeNull();
    expect(beforeApply.substituteTeacherId).toBeNull();

    const applyResult = await autoApplyUnambiguousTeacherMatches(result.monthlyImport.id, "測試管理者");
    expect(applyResult.appliedCount).toBe(2);
    expect(applyResult.remainingCount).toBe(0);

    const afterApply = await prisma.substituteRecord.findFirstOrThrow({
      where: { monthlyImportId: result.monthlyImport.id },
    });
    expect(afterApply.originalTeacherId).toBe(createdOriginal.person.id);
    expect(afterApply.substituteTeacherId).toBe(createdSub.person.id);
  });

  it("找不到對應人員時標示為待處理，且不會偷偷建立新人", async () => {
    const sem = await makeTestSemester(604);
    const beforeCount = await prisma.person.count({ where: { name: "測試查無此人乙" } });

    const buffer = await buildWorkbookBuffer(STANDARD_HEADERS, [
      standardRow({ 原教師: "測試查無此人乙" }),
    ]);
    const result = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "test.xlsx",
      fileBuffer: buffer,
    });

    const unmatched = await listUnmatchedTeacherReferences(result.monthlyImport.id);
    const originalIssue = unmatched.find((u) => u.field === "original");
    expect(originalIssue?.rawName).toBe("測試查無此人乙");
    expect(originalIssue?.candidates).toEqual([]);

    const afterCount = await prisma.person.count({ where: { name: "測試查無此人乙" } });
    expect(afterCount).toBe(beforeCount); // 沒有偷偷建立新人

    // 人工個別配對仍然可行
    const created = await createPerson({ name: "測試查無此人乙" });
    if (created.status !== "CREATED") throw new Error("unreachable");
    cleanupPersonIds.push(created.person.id);
    await resolveTeacherReference(originalIssue!.recordId, "original", created.person.id, "測試管理者");
    const updated = await prisma.substituteRecord.findUniqueOrThrow({ where: { id: originalIssue!.recordId } });
    expect(updated.originalTeacherId).toBe(created.person.id);
  });
});

describe("8/9. 重新匯入策略", () => {
  it("同學期同月份重新匯入，舊批次轉為 SUPERSEDED、新批次 ACTIVE，資料都保留", async () => {
    const sem = await makeTestSemester(605);
    const buffer1 = await buildWorkbookBuffer(STANDARD_HEADERS, [standardRow({ 班級: "第一版" })]);
    const first = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "v1.xlsx",
      fileBuffer: buffer1,
    });

    const buffer2 = await buildWorkbookBuffer(STANDARD_HEADERS, [
      standardRow({ 班級: "第二版A" }),
      standardRow({ 班級: "第二版B" }),
    ]);
    const second = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "v2.xlsx",
      fileBuffer: buffer2,
    });

    expect(second.supersededImportIds).toEqual([first.monthlyImport.id]);
    expect(second.monthlyImport.versionNo).toBe(first.monthlyImport.versionNo + 1);

    const firstAfter = await prisma.monthlyImport.findUniqueOrThrow({ where: { id: first.monthlyImport.id } });
    expect(firstAfter.status).toBe("SUPERSEDED");
    const secondAfter = await prisma.monthlyImport.findUniqueOrThrow({ where: { id: second.monthlyImport.id } });
    expect(secondAfter.status).toBe("ACTIVE");

    // 舊批次資料沒有被刪除
    const oldRecords = await prisma.substituteRecord.findMany({ where: { monthlyImportId: first.monthlyImport.id } });
    expect(oldRecords.length).toBe(1);

    const batches = await listMonthlyImports(sem.id, { year: 2026, month: 9 });
    expect(batches.length).toBe(2);
  });

  it("跨學期同年月匯入互不影響", async () => {
    const semA = await makeTestSemester(606);
    const semB = await makeTestSemester(607);
    const buffer = await buildWorkbookBuffer(STANDARD_HEADERS, [standardRow()]);

    const resultA = await importSubstituteExcel({
      semesterId: semA.id,
      year: 2026,
      month: 9,
      fileName: "a.xlsx",
      fileBuffer: buffer,
    });
    const resultB = await importSubstituteExcel({
      semesterId: semB.id,
      year: 2026,
      month: 9,
      fileName: "b.xlsx",
      fileBuffer: buffer,
    });

    expect(resultA.monthlyImport.status).toBe("ACTIVE");
    expect(resultB.monthlyImport.status).toBe("ACTIVE"); // 不會因為 A 學期匯入而被取代
    expect(resultA.supersededImportIds).toEqual([]);
    expect(resultB.supersededImportIds).toEqual([]);
  });
});

describe("10. 缺欄位", () => {
  it("Excel 缺少必要欄位（例如日期）時，整份拒絕匯入且不建立任何批次", async () => {
    const sem = await makeTestSemester(608);
    const headersWithoutDate = STANDARD_HEADERS.filter((h) => h !== "日期");
    const buffer = await buildWorkbookBuffer(
      headersWithoutDate,
      [headersWithoutDate.map(() => "x")]
    );

    const beforeCount = await prisma.monthlyImport.count({ where: { semesterId: sem.id } });
    await expect(
      importSubstituteExcel({ semesterId: sem.id, year: 2026, month: 9, fileName: "bad.xlsx", fileBuffer: buffer })
    ).rejects.toThrow();
    const afterCount = await prisma.monthlyImport.count({ where: { semesterId: sem.id } });
    expect(afterCount).toBe(beforeCount);
  });
});

describe("11. 異常資料列（部分失敗，不整批放棄）", () => {
  it("缺原教師的資料列不會建立 SubstituteRecord，但原始資料仍保存，其他列正常匯入", async () => {
    const sem = await makeTestSemester(609);
    const buffer = await buildWorkbookBuffer(STANDARD_HEADERS, [
      standardRow({ 原教師: null }),
      standardRow({ 班級: "正常列" }),
    ]);

    const result = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "mixed.xlsx",
      fileBuffer: buffer,
    });

    expect(result.totalCount).toBe(2);
    expect(result.successCount).toBe(1);
    expect(result.errorCount).toBe(1);
    expect(result.errors[0]).toMatchObject({ rowNumber: 1, fieldName: "原教師" });

    const rawRows = await prisma.substituteRecordRaw.findMany({ where: { monthlyImportId: result.monthlyImport.id } });
    expect(rawRows.length).toBe(2); // 兩列的原始資料都保存

    const errorRecords = await prisma.monthlyImportError.findMany({ where: { monthlyImportId: result.monthlyImport.id } });
    expect(errorRecords.length).toBe(1);
    expect(errorRecords[0].message).toBe("缺少原教師");
  });
});

describe("12. Raw 與標準化資料可追溯回同一 Import Batch", () => {
  it("SubstituteRecord 與其 raw 都指向同一個 monthlyImportId，且 1:1 對應", async () => {
    const sem = await makeTestSemester(610);
    const buffer = await buildWorkbookBuffer(STANDARD_HEADERS, [standardRow()]);
    const result = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "trace.xlsx",
      fileBuffer: buffer,
    });

    const record = await prisma.substituteRecord.findFirstOrThrow({
      where: { monthlyImportId: result.monthlyImport.id },
      include: { rawRecord: true },
    });
    expect(record.rawRecord.monthlyImportId).toBe(result.monthlyImport.id);
    expect(record.monthlyImportId).toBe(result.monthlyImport.id);
    expect(record.rawRecordId).toBe(record.rawRecord.id);
  });
});

// 以下涵蓋用真實「114學年2026.06月代課(公費).xlsx」驗證時發現、原本測試沒覆蓋到的情況。
describe("13. 真實檔案驗證後發現並修正的能力", () => {
  it("listWorkbookSheets 會列出檔案裡所有工作表", async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet("教師代碼對照");
    const detailSheet = workbook.addWorksheet("6月公費 非BD");
    detailSheet.addRow(STANDARD_HEADERS);
    detailSheet.addRow(standardRow());
    const buffer = Buffer.from(await workbook.xlsx.writeBuffer());

    const sheets = await listWorkbookSheets(buffer);
    expect(sheets.map((s) => s.name)).toEqual(["教師代碼對照", "6月公費 非BD"]);
  });

  it("可以指定要解析哪一個工作表，不會永遠假設是第一個（真實檔案第一個 Sheet 是教師代碼對照表）", async () => {
    const sem = await makeTestSemester(611);
    const workbook = new ExcelJS.Workbook();
    const lookupSheet = workbook.addWorksheet("教師代碼對照");
    lookupSheet.addRow(["姓名", "員工代號", "部門名稱"]);
    lookupSheet.addRow(["某某人", "A00001", "A"]);
    const detailSheet = workbook.addWorksheet("代課明細");
    detailSheet.addRow(STANDARD_HEADERS);
    detailSheet.addRow(standardRow());
    const buffer = Buffer.from(await workbook.xlsx.writeBuffer());

    const result = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "multi.xlsx",
      fileBuffer: buffer,
      sheetName: "代課明細",
    });
    expect(result.totalCount).toBe(1);
    expect(result.successCount).toBe(1);
  });

  it("表頭以 richText 儲存（例如 Alt+Enter 換行）時仍能正確辨識欄位，不會整欄消失", async () => {
    const sem = await makeTestSemester(612);
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet1");
    const headerRow = sheet.addRow([...STANDARD_HEADERS]);
    // 把「原教師代碼」改成 richText 且中間帶換行，模擬真實檔案的表頭儲存格
    headerRow.getCell(2).value = { richText: [{ text: "原教師\n" }, { text: "代碼" }] } as unknown as string;
    sheet.addRow(standardRow());
    const buffer = Buffer.from(await workbook.xlsx.writeBuffer());

    const result = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "richtext.xlsx",
      fileBuffer: buffer,
    });
    expect(result.successCount).toBe(1);
    const raw = await prisma.substituteRecordRaw.findFirstOrThrow({ where: { monthlyImportId: result.monthlyImport.id } });
    expect(raw.originalTeacherCodeText).toBe("T001");
  });

  it("表頭前面有標題列（合併儲存格式的月份標題）時，仍能自動找到真正的表頭列", async () => {
    const sem = await makeTestSemester(613);
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet1");
    sheet.addRow(["115學年第1學期 2026.09月公費代課 ( 編制外 )"]); // 標題列
    sheet.addRow([...STANDARD_HEADERS]); // 真正表頭
    sheet.addRow(standardRow());
    const buffer = Buffer.from(await workbook.xlsx.writeBuffer());

    const result = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "titlerow.xlsx",
      fileBuffer: buffer,
    });
    expect(result.totalCount).toBe(1);
    expect(result.successCount).toBe(1);
  });

  it("表頭包含逐月變動的金額／天數文字時，仍可用局部比對找到欄位並保留原始文字", async () => {
    const sem = await makeTestSemester(614);
    const extendedHeaders = [
      ...STANDARD_HEADERS.map((h) => (h === "代導師" ? "6月(30天)代導師費$133" : h)),
      "日薪$1399\n半日薪$700",
      "代課\n鐘點$405",
      "節數",
    ];
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet1");
    sheet.addRow(extendedHeaders);
    sheet.addRow([...standardRow(), "700", "405", "3"]);
    const buffer = Buffer.from(await workbook.xlsx.writeBuffer());

    const result = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "dynamic-headers.xlsx",
      fileBuffer: buffer,
    });
    expect(result.successCount).toBe(1);
    const raw = await prisma.substituteRecordRaw.findFirstOrThrow({ where: { monthlyImportId: result.monthlyImport.id } });
    expect(raw.dailyOrHalfDayWageText).toBe("700");
    expect(raw.substitutePeriodFeeText).toBe("405");
    expect(raw.periodCountText).toBe("3");
  });

  it("日期使用斜線分隔、沒有補零也能解析（真實檔案是 M/D 而非規格書範例的 MM-DD）", () => {
    const result = parseDateText("6/12(五)", 2026, 6);
    if ("error" in result) throw new Error("unexpected error");
    expect(result.date.toISOString().slice(0, 10)).toBe("2026-06-12");
  });

  it("整月區間或跨日區間不會被誤判成單一日期，一律視為無法解析", () => {
    expect("error" in parseDateText("6/1~6/30", 2026, 6)).toBe(true);
    expect("error" in parseDateText("06-26(五) 07:50 ~ 06-29(一) 15:50", 2026, 6)).toBe(true);
  });

  it("BD／非BD 來源會標示在 SubstituteRecord.staffType，且不影響任何分類或金額欄位", async () => {
    const sem = await makeTestSemester(615);
    const buffer = await buildWorkbookBuffer(STANDARD_HEADERS, [standardRow()]);
    const result = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "bd.xlsx",
      fileBuffer: buffer,
      sourceStaffType: "BD",
    });
    const record = await prisma.substituteRecord.findFirstOrThrow({ where: { monthlyImportId: result.monthlyImport.id } });
    expect(record.staffType).toBe("BD");
    expect(record.fundingSource).toBe("UNDETERMINED");
    expect(record.amount).toBeNull();
  });

  it("同月份 BD 與非BD 分開匯入互不取代；只有相同來源類型重新匯入才會取代舊批次", async () => {
    const sem = await makeTestSemester(616);
    const buffer = await buildWorkbookBuffer(STANDARD_HEADERS, [standardRow()]);

    const bdResult = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "bd.xlsx",
      fileBuffer: buffer,
      sourceStaffType: "BD",
    });
    const nonBdResult = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "nonbd.xlsx",
      fileBuffer: buffer,
      sourceStaffType: "NON_BD",
    });
    expect(nonBdResult.supersededImportIds).toEqual([]);

    const bdBatch = await prisma.monthlyImport.findUniqueOrThrow({ where: { id: bdResult.monthlyImport.id } });
    expect(bdBatch.status).toBe("ACTIVE");

    const bdResult2 = await importSubstituteExcel({
      semesterId: sem.id,
      year: 2026,
      month: 9,
      fileName: "bd-v2.xlsx",
      fileBuffer: buffer,
      sourceStaffType: "BD",
    });
    expect(bdResult2.supersededImportIds).toEqual([bdResult.monthlyImport.id]);

    const nonBdBatchAfter = await prisma.monthlyImport.findUniqueOrThrow({ where: { id: nonBdResult.monthlyImport.id } });
    expect(nonBdBatchAfter.status).toBe("ACTIVE");
  });
});
