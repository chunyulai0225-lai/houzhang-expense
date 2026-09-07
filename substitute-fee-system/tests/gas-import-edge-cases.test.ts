// 真實 2026/09 公費代課資料出現兩種原本會被判定「無法解析」的格式：
//   1. 日期區間：「09-11(五) 12:00 ~ 09-14(一) 16:00」
//   2. 特殊節次：「導師時間」
// 這裡驗證 gas/Import.gs 的修正：兩種都不再被當成「無法解析」的錯誤。
//
// 日期區間後來又依實際資料再修正過一次（見 tests/gas-date-range-period-count.test.ts）：
// 一開始把「日期是區間」本身當成要擋下、不建立 SubstituteRecord 的原因，但實際資料
// 顯示區間只是原教師請假／出差的期間，真正的代課計算單位是「節次 × 時數天數」——
// 所以現在日期區間本身「不會」阻止建立 SubstituteRecord（parseDateText 會用區間
// 起始日當作日期定位點），這裡的測試只保留「日期解析本身認得出區間格式」這件事；
// 「區間 + 節次 + 時數天數」完整搭配起來會不會真的建立紀錄，交給
// gas-date-range-period-count.test.ts 驗證。
// 「導師時間」則仍維持這裡驗證的行為：辨識但不硬猜成某一節，一律保留原始
// RawRecord、標記「待確認」、不建立 SubstituteRecord。同時要確認既有單一日期格式、
// 既有節次格式的解析結果完全沒變（2026/06 的迴歸基準不能被這次修改動到）。
import { describe, expect, it } from "vitest";
import { createGasSandbox, seedRealSemester115_1 } from "./helpers/gasHarness";

describe("日期解析：既有單一日期格式（不能被這次修改影響）", () => {
  it("標準「MM-DD(星期)」格式照常解析成功", () => {
    const sandbox = createGasSandbox();
    const result = sandbox.parseDateText("06-18(四)", 2026, 6);
    expect(result.error).toBeUndefined();
    expect(result.isDateRange).toBeUndefined();
    expect(result.date).toBe("2026-06-18");
    expect(result.weekday).toBe("THU");
  });

  it("同一天內的起訖時間（非跨日區間）照常解析成功，不會被誤判成日期區間", () => {
    const sandbox = createGasSandbox();
    const result = sandbox.parseDateText("6/18(四) 13:50 ~ 15:50", 2026, 6);
    expect(result.error).toBeUndefined();
    expect(result.isDateRange).toBeUndefined();
    expect(result.date).toBe("2026-06-18");
  });

  it("真的無法辨識的日期文字仍然回報錯誤（不會被誤判成日期區間）", () => {
    const sandbox = createGasSandbox();
    const result = sandbox.parseDateText("這不是日期", 2026, 6);
    expect(result.isDateRange).toBeUndefined();
    expect(result.error).toBeTruthy();
  });
});

describe("日期解析：新的日期區間格式（以區間起始日作為日期定位點，不是錯誤）", () => {
  it('"09-11(五) 12:00 ~ 09-14(一) 16:00" 辨識為日期區間，日期採用區間起始日 09-11', () => {
    const sandbox = createGasSandbox();
    const result = sandbox.parseDateText("09-11(五) 12:00 ~ 09-14(一) 16:00", 2026, 9);
    expect(result.error).toBeUndefined();
    expect(result.isDateRange).toBe(true);
    // 只取「起始日」當定位點，不猜區間裡哪一天才是真正代課日、也不展開成多筆紀錄。
    expect(result.date).toBe("2026-09-11");
    expect(result.weekday).toBe("FRI");
  });

  it('"09-21(一) 12:00 ~ 09-24(四) 16:00" 也辨識為日期區間，日期採用區間起始日 09-21', () => {
    const sandbox = createGasSandbox();
    const result = sandbox.parseDateText("09-21(一) 12:00 ~ 09-24(四) 16:00", 2026, 9);
    expect(result.isDateRange).toBe(true);
    expect(result.date).toBe("2026-09-21");
  });
});

describe("節次解析：既有格式（不能被這次修改影響）", () => {
  it("早自修／午休／第N節照常解析成功", () => {
    const sandbox = createGasSandbox();
    const codes: Record<string, boolean> = { EARLY_STUDY: true, LUNCH: true, P1: true, P5: true };
    expect(sandbox.parsePeriodText("早自修", codes)).toEqual({ periodCode: "EARLY_STUDY" });
    expect(sandbox.parsePeriodText("午休", codes)).toEqual({ periodCode: "LUNCH" });
    expect(sandbox.parsePeriodText("第1節", codes)).toEqual({ periodCode: "P1" });
    expect(sandbox.parsePeriodText("第五節", codes)).toEqual({ periodCode: "P5" });
  });

  it("真的無法辨識的節次文字仍然回報錯誤", () => {
    const sandbox = createGasSandbox();
    const result = sandbox.parsePeriodText("第1,2節", {});
    expect(result.isSpecialPeriod).toBeUndefined();
    expect(result.error).toBeTruthy();
  });
});

describe("節次解析：新的特殊節次「導師時間」", () => {
  it('"導師時間" 辨識為特殊節次，不是錯誤、也不會被硬轉成 P1~P7', () => {
    const sandbox = createGasSandbox();
    const result = sandbox.parsePeriodText("導師時間", { P1: true, P2: true });
    expect(result.error).toBeUndefined();
    expect(result.periodCode).toBeUndefined();
    expect(result.isSpecialPeriod).toBe(true);
    expect(result.specialPeriodText).toBe("導師時間");
  });
});

describe("完整匯入流程：特殊節次保留 RawRecord、不建立 SubstituteRecord、標記待確認；日期區間本身不會擋下", () => {
  it("一批資料裡混合正常列、日期區間但缺時數天數的列、特殊節次列：三筆 RawRecord 都保留，正常列跟「有沒有另外被時數天數擋下」互不影響", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);

    const rows = [
      {
        rowNumber: 1, raw: {}, originalTeacherName: "王老師", substituteTeacherName: "陳老師",
        dateText: "09-02(三)", periodText: "第1節", className: "1年1班", subject: "國語",
      },
      {
        // 日期是區間、節次是第2節，但沒有「時數天數」可以確認計費數量——這種情況
        // 不是被日期區間擋下，而是被「時數天數缺漏」擋下（見
        // tests/gas-date-range-period-count.test.ts 驗證「有時數天數」時能正常建立）。
        rowNumber: 2, raw: {}, originalTeacherName: "王老師", substituteTeacherName: "陳老師",
        dateText: "09-11(五) 12:00 ~ 09-14(一) 16:00", periodText: "第2節", className: "1年1班", subject: "數學",
      },
      {
        rowNumber: 3, raw: {}, originalTeacherName: "王老師", substituteTeacherName: "陳老師",
        dateText: "09-15(二)", periodText: "導師時間", className: "1年1班", subject: "",
      },
    ];

    const result = sandbox.api_importSubstituteRows({
      semesterId: semester.id, year: 2026, month: 9, fileName: "測試.xlsx", sheetName: "工作表1",
      sourceStaffType: "NON_BD", importedBy: "測試", rows, detectedHeaders: ["日期", "節次"],
    });

    expect(result.totalCount).toBe(3);
    expect(result.successCount).toBe(1);
    expect(result.errorCount).toBe(2);

    // RawRecords：三筆都要在，完全不會因為日期區間/特殊節次而消失。
    const rawRecords = sandbox.readRows("RawRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(rawRecords).toHaveLength(3);

    // SubstituteRecords：只有第 1 列（正常日期＋正常節次）會建立。
    const substituteRecords = sandbox.api_listSubstituteRecords({ id: result.monthlyImport.id });
    expect(substituteRecords).toHaveLength(1);
    expect(substituteRecords[0].date).toBe("2026-09-02");
    expect(substituteRecords[0].periodCode).toBe("P1");

    // ImportErrors：第2列要標示成「時數天數／待確認」（不是日期區間的錯，日期區間
    // 本身不再是被擋下的原因），第3列仍是「特殊節次／待確認」，不能還是原本嚇人的
    // 「無法解析」字樣。
    const messages = result.errors.map((e: any) => e.message);
    expect(messages.some((m: string) => m.includes("時數天數／待確認") && m.includes("09-11(五) 12:00 ~ 09-14(一) 16:00"))).toBe(true);
    expect(messages.some((m: string) => m.includes("日期區間／待確認"))).toBe(false);
    expect(messages.some((m: string) => m.includes("特殊節次／待確認") && m.includes("導師時間"))).toBe(true);
    expect(messages.some((m: string) => m.includes("無法解析日期格式"))).toBe(false);
    expect(messages.some((m: string) => m.includes('無法解析節次："導師時間"'))).toBe(false);
  });

  it("特殊節次「導師時間」的原始文字（即使日期也是區間），在 RawRecord 裡完整保留，之後仍查得到", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const rows = [
      {
        rowNumber: 1, raw: {}, originalTeacherName: "王老師", substituteTeacherName: "陳老師",
        dateText: "09-21(一) 12:00 ~ 09-24(四) 16:00", periodText: "導師時間", className: "2年3班", subject: "",
      },
    ];
    const result = sandbox.api_importSubstituteRows({
      semesterId: semester.id, year: 2026, month: 9, fileName: "測試2.xlsx", sheetName: "工作表1",
      sourceStaffType: "BD", importedBy: "測試", rows, detectedHeaders: ["日期", "節次"],
    });
    const raw = sandbox.readRows("RawRecords").find((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(raw.dateText).toBe("09-21(一) 12:00 ~ 09-24(四) 16:00");
    expect(raw.periodText).toBe("導師時間");
  });
});
