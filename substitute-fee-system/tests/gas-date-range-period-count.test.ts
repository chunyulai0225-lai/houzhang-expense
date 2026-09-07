// 實際 2026/09 公費代課 Excel 資料顯示：前一輪把「日期是區間」本身當成要擋下、
// 不建立 SubstituteRecord 的原因是錯的。實際情況是：
//
//   劉馨憶：日期「09-07(一) 07:50 ~ 09-11(五) 15:50」只是原教師請假／出差的期間，
//   Excel 已經把應代課的每個節次拆成不同列（第二節/徐碧苓/5日、第三節/徐碧苓/5日…），
//   真正的代課計算單位是「節次 × 時數天數」——5日 × 第二節 = 5 個一般代課節次。
//
// 這裡驗證 gas/Import.gs 的修正：
//   1. 日期區間本身不再導致該列無法建立 SubstituteRecord、也不再標記成 ImportError。
//   2. 不會把日期區間展開成每天一筆的 SubstituteRecord（永遠只有一筆）。
//   3. 有明確第一節～第七節、且「時數天數」能安全解析成純日數（例如「5日」）時，
//      SubstituteRecord.periodCount 直接採用該數字，不是從日期區間自己算出天數。
//   4. 「時數天數」無法安全解析（例如「3日4時」這種混合單位）時，保留 RawRecord、
//      標記「時數天數／待確認」，不建立 SubstituteRecord——但這是時數天數的問題，
//      不是日期區間的問題，訊息不能誤植成日期相關的錯誤。
//   5. 單一日期的既有資料、以及導師時間／午休的既有規則，完全不受這次修改影響。
import { describe, expect, it } from "vitest";
import { createGasSandbox, seedRealSemester115_1 } from "./helpers/gasHarness";

function importOneRow(sandbox: any, semesterId: string, row: Record<string, any>) {
  return sandbox.api_importSubstituteRows({
    semesterId, year: 2026, month: 9, fileName: "測試.xlsx", sheetName: "工作表1",
    sourceStaffType: "NON_BD", importedBy: "測試", rows: [{ rowNumber: 1, raw: {}, ...row }],
    detectedHeaders: ["日期", "節次", "時數天數"],
  });
}

describe("parseHoursOrDaysToPeriodCount()：只認得純日數格式，不猜混合單位", () => {
  it("「5日」「8日」「12日」都能安全解析成對應的節次計費數量", () => {
    const sandbox = createGasSandbox();
    expect(sandbox.parseHoursOrDaysToPeriodCount("5日")).toEqual({ periodCount: 5 });
    expect(sandbox.parseHoursOrDaysToPeriodCount("8日")).toEqual({ periodCount: 8 });
    expect(sandbox.parseHoursOrDaysToPeriodCount("12日")).toEqual({ periodCount: 12 });
  });

  it("混合單位「3日4時」、純時數「6時／2時／1時」、空白，都無法安全解析（不猜測換算規則）", () => {
    const sandbox = createGasSandbox();
    expect(sandbox.parseHoursOrDaysToPeriodCount("3日4時").error).toBeTruthy();
    expect(sandbox.parseHoursOrDaysToPeriodCount("6時").error).toBeTruthy();
    expect(sandbox.parseHoursOrDaysToPeriodCount("2時").error).toBeTruthy();
    expect(sandbox.parseHoursOrDaysToPeriodCount("1時").error).toBeTruthy();
    expect(sandbox.parseHoursOrDaysToPeriodCount("").error).toBeTruthy();
    expect(sandbox.parseHoursOrDaysToPeriodCount(undefined).error).toBeTruthy();
  });
});

describe("真實資料：劉馨憶 09-07~09-11、5日、第二節 → 不因日期區間擋下，計費數量為 5", () => {
  it("能正常建立 SubstituteRecord，date 為區間起始日，periodCode=P2，periodCount=5", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);

    const result = importOneRow(sandbox, semester.id, {
      originalTeacherName: "王老師", substituteTeacherName: "劉馨憶",
      dateText: "09-07(一) 07:50 ~ 09-11(五) 15:50", periodText: "第2節",
      hoursOrDaysText: "5日", className: "1年1班", subject: "",
    });

    // 不應因日期區間而 ImportError。
    expect(result.errorCount).toBe(0);
    expect(result.successCount).toBe(1);
    expect(result.errors).toHaveLength(0);

    const raw = sandbox.readRows("RawRecords").find((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(raw.dateText).toBe("09-07(一) 07:50 ~ 09-11(五) 15:50");
    expect(raw.hoursOrDaysText).toBe("5日");

    const records = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    // 只有一筆，沒有被展開成每天一筆。
    expect(records).toHaveLength(1);
    expect(records[0].date).toBe("2026-09-07"); // 區間起始日，不是自行推算出的某一天
    expect(records[0].periodCode).toBe("P2");
    expect(Number(records[0].periodCount)).toBe(5); // 計費數量＝時數天數，不是日期區間算出來的天數
    expect(records[0].rawHoursOrDays).toBe("5日");
  });
});

describe("真實資料：詹庭瑜兩筆不同區間／天數，都不應因日期區間而 ImportError", () => {
  it("08-31~09-09、8日、第一節 → 正常建立，periodCount=8", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const result = importOneRow(sandbox, semester.id, {
      originalTeacherName: "王老師", substituteTeacherName: "詹庭瑜",
      dateText: "08-31(一) 07:50 ~ 09-09(三) 15:50", periodText: "第1節",
      hoursOrDaysText: "8日", className: "2年1班", subject: "",
    });
    expect(result.errorCount).toBe(0);
    expect(result.successCount).toBe(1);
    const records = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(records).toHaveLength(1);
    expect(records[0].periodCode).toBe("P1");
    expect(Number(records[0].periodCount)).toBe(8);
  });

  it("09-10~09-29、12日、第一節 → 正常建立，periodCount=12", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const result = importOneRow(sandbox, semester.id, {
      originalTeacherName: "王老師", substituteTeacherName: "詹庭瑜",
      dateText: "09-10(四) 07:50 ~ 09-29(二) 15:50", periodText: "第1節",
      hoursOrDaysText: "12日", className: "2年1班", subject: "",
    });
    expect(result.errorCount).toBe(0);
    expect(result.successCount).toBe(1);
    const records = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(records).toHaveLength(1);
    expect(records[0].periodCode).toBe("P1");
    expect(Number(records[0].periodCount)).toBe(12);
  });
});

describe("真實資料：林家德 09-11~09-14、3日4時、第六節 → 日期區間不得直接造成 ImportError，但時數天數無法安全解析時暫不建立紀錄", () => {
  it("RawRecord 完整保留、標記「時數天數／待確認」；不建立 SubstituteRecord；訊息不是日期區間的錯", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const result = importOneRow(sandbox, semester.id, {
      originalTeacherName: "王老師", substituteTeacherName: "林家德",
      dateText: "09-11(五) 07:50 ~ 09-14(一) 15:50", periodText: "第6節",
      hoursOrDaysText: "3日4時", className: "3年1班", subject: "",
    });

    expect(result.successCount).toBe(0);
    expect(result.errorCount).toBe(1);

    const raw = sandbox.readRows("RawRecords").find((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(raw).toBeTruthy();
    expect(raw.dateText).toBe("09-11(五) 07:50 ~ 09-14(一) 15:50");
    expect(raw.hoursOrDaysText).toBe("3日4時");

    const records = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(records).toHaveLength(0);

    const message = result.errors[0].message as string;
    expect(message).toContain("時數天數／待確認");
    expect(message).toContain("3日4時");
    // 不可以誤植成日期區間本身的錯——日期區間不是這裡被擋下的原因。
    expect(message).not.toContain("日期區間／待確認");
    expect(message).not.toContain("無法解析日期格式");
  });
});

describe("單一日期＋時數天數：既有格式完全不受影響（periodCount 維持原本空白，2026/06 迴歸不受影響）", () => {
  it("單一日期列即使剛好也有「時數天數」欄位，periodCount 仍維持空白，不會被這次新增的判斷擋下", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const result = importOneRow(sandbox, semester.id, {
      originalTeacherName: "王老師", substituteTeacherName: "陳老師",
      dateText: "09-02(三)", periodText: "第1節", hoursOrDaysText: "1日",
      className: "1年1班", subject: "國語",
    });
    expect(result.errorCount).toBe(0);
    expect(result.successCount).toBe(1);
    const records = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(records[0].periodCount).toBeNull(); // 沿用既有行為，不因為這次新增的邏輯而改變
  });
});

describe("導師時間／午休：即使日期是區間，非計費規則仍然成立，不受這次修改影響", () => {
  it("午休＋日期區間：不會被時數天數擋下（非授課節次不受此限），仍會建立紀錄但不計一般代課鐘點費", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const result = importOneRow(sandbox, semester.id, {
      originalTeacherName: "王老師", substituteTeacherName: "陳老師",
      dateText: "09-07(一) 07:50 ~ 09-11(五) 15:50", periodText: "午休",
      className: "1年1班", subject: "",
    });
    expect(result.errorCount).toBe(0);
    expect(result.successCount).toBe(1);
    const records = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(records).toHaveLength(1);
    expect(records[0].periodCode).toBe("LUNCH");
    expect(sandbox.isNonPayablePeriodCode("LUNCH")).toBe(true);
  });

  it("導師時間＋日期區間：維持既有「特殊節次／待確認」行為，不建立 SubstituteRecord", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const result = importOneRow(sandbox, semester.id, {
      originalTeacherName: "王老師", substituteTeacherName: "陳老師",
      dateText: "09-07(一) 07:50 ~ 09-11(五) 15:50", periodText: "導師時間",
      className: "1年1班", subject: "",
    });
    expect(result.successCount).toBe(0);
    const records = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(records).toHaveLength(0);
    expect(result.errors[0].message).toContain("特殊節次／待確認");
  });
});
