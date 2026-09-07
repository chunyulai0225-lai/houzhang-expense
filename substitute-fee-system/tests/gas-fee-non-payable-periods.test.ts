// 后庄國小正式作息表確認：115-1 只有「第一節～第七節」屬於會產生一般代課鐘點費
// （SUBSTITUTE_PERIOD）的授課節次；導師時間（08:15~08:45）、午休（12:10~13:20）
// 都不計一般代課鐘點費，即使這兩種節次出現在請假／代課資料中、且被自動分類成
// GENERAL 或 PROJECT，也一律不能算出一般代課鐘點費。
//
// 這裡驗證 gas/FeeCalculation.gs 新增的 isNonPayablePeriodCode() 判斷式與
// calculateSubstituteRecordFee() 的擋下邏輯：完全依賴 PeriodSlots 既有的
// isTeachingPeriod 欄位（早自修／午休原本就是 false，P1~P7 原本就是 true，
// 沒有新增欄位也沒有重新設計資料模型），只在 fundingSource 對應到
// SUBSTITUTE_PERIOD 費用類型、且節次被標記為非授課節次時才擋下——不影響
// OVERTIME_PERIOD 的計算路徑，也不影響第一～第七節既有的 Phase9-5 計算邏輯。
//
// 「導師時間」目前仍維持前一輪已確認的行為：辨識但不轉成 P1~P7、不建立
// SubstituteRecord、保留 RawRecord 標記待確認（見 gas-import-edge-cases.test.ts）；
// 這裡額外從「計算流程」角度再次確認：既然沒有 SubstituteRecord，自然也就
// 不可能算出任何一般代課鐘點費。
import { describe, expect, it } from "vitest";
import { createGasSandbox, seedRealSemester115_1 } from "./helpers/gasHarness";

describe("isNonPayablePeriodCode()：只依賴 PeriodSlots.isTeachingPeriod 判斷，不新增資料模型", () => {
  it("午休／早自修是非授課節次", () => {
    const sandbox = createGasSandbox();
    expect(sandbox.isNonPayablePeriodCode("LUNCH")).toBe(true);
    expect(sandbox.isNonPayablePeriodCode("EARLY_STUDY")).toBe(true);
  });

  it("第一節～第七節都是授課節次，完全不受影響", () => {
    const sandbox = createGasSandbox();
    for (let i = 1; i <= 7; i++) {
      expect(sandbox.isNonPayablePeriodCode("P" + i)).toBe(false);
    }
  });

  it("空節次代碼、或系統裡沒有設定的節次代碼，不影響既有判斷流程（回傳 false，不擋下）", () => {
    const sandbox = createGasSandbox();
    expect(sandbox.isNonPayablePeriodCode("")).toBe(false);
    expect(sandbox.isNonPayablePeriodCode(null)).toBe(false);
    expect(sandbox.isNonPayablePeriodCode("NOT_A_REAL_CODE")).toBe(false);
  });
});

describe("完整流程：匯入 → 分類 → 計費，午休即使被分類為一般公費也不計一般代課鐘點費，第一節不受影響", () => {
  it("同一批資料裡混合午休列與第1節列：午休列 fundingSource=GENERAL 但 amount 維持 null；第1節照常算出金額", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    sandbox.api_createFeeRule({
      semesterId: semester.id, feeType: "SUBSTITUTE_PERIOD", amount: 165,
      effectiveDate: "2026-08-31", changedBy: "測試",
    });

    var teacher = {
      id: sandbox.newId(), name: "王老師", payrollCode: "", enrollmentStatus: "ACTIVE",
      enrollDate: "", leaveDate: "", note: "", createdAt: sandbox.nowIso(), updatedAt: sandbox.nowIso(),
    };
    sandbox.appendRow("Persons", teacher);

    const rows = [
      {
        rowNumber: 1, raw: {}, originalTeacherName: "王老師", substituteTeacherName: "陳老師",
        dateText: "09-02(三)", periodText: "午休", className: "1年1班", subject: "",
      },
      {
        rowNumber: 2, raw: {}, originalTeacherName: "王老師", substituteTeacherName: "陳老師",
        dateText: "09-02(三)", periodText: "第1節", className: "1年1班", subject: "國語",
      },
    ];

    const result = sandbox.api_importSubstituteRows({
      semesterId: semester.id, year: 2026, month: 9, fileName: "測試.xlsx", sheetName: "工作表1",
      sourceStaffType: "NON_BD", importedBy: "測試", rows, detectedHeaders: ["日期", "節次"],
    });
    expect(result.successCount).toBe(2);

    // 教師配對是另一個獨立步驟（api_resolveTeacherReference 等），這裡直接指定配對結果，
    // 只測「配對後 → 分類 → 計費」這條流程對非授課節次的行為，不重新測配對邏輯本身。
    const importedRecords = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    importedRecords.forEach((r: any) => sandbox.updateRow("SubstituteRecords", r.id, { originalTeacherId: teacher.id, updatedAt: sandbox.nowIso() }));

    sandbox.api_classifyMonthlyImport({ id: result.monthlyImport.id, changedBy: "測試" });
    sandbox.api_calculateMonthlyImportFees({ id: result.monthlyImport.id, changedBy: "測試" });

    const lunchRecord = sandbox.readRows("SubstituteRecords").find((r: any) => r.monthlyImportId === result.monthlyImport.id && r.periodCode === "LUNCH");
    const p1Record = sandbox.readRows("SubstituteRecords").find((r: any) => r.monthlyImportId === result.monthlyImport.id && r.periodCode === "P1");

    // 先確認午休真的有被分類成一般公費（不是卡在分類階段，才有意義驗證「計費階段」把它擋下）。
    expect(lunchRecord.fundingSource).toBe("GENERAL");
    expect(lunchRecord.amount).toBeNull();
    expect(lunchRecord.unitPrice).toBeNull();

    // 第一節完全不受影響，正常照 FeeRule 算出金額。
    expect(p1Record.fundingSource).toBe("GENERAL");
    expect(p1Record.amount).toBe("165");
    expect(p1Record.unitPrice).toBe("165");

    // 直接呼叫單筆計費 API，確認回傳的 skippedReason 有清楚說明原因（非授課節次），
    // 不是被誤判成 fundingSource=UNDETERMINED 或其他原因跳過。
    const lunchResult = sandbox.api_calculateSubstituteRecordFee({ id: lunchRecord.id, changedBy: "測試" });
    expect(lunchResult.amount).toBeNull();
    expect(lunchResult.skippedReason).toContain("非授課節次");
    expect(lunchResult.skippedReason).not.toContain("UNDETERMINED");
  });

  it("彙總報表（summarizeTeacherMonthlyFees）不會把午休算進一般公費總額", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    sandbox.api_createFeeRule({
      semesterId: semester.id, feeType: "SUBSTITUTE_PERIOD", amount: 165,
      effectiveDate: "2026-08-31", changedBy: "測試",
    });
    var teacher = {
      id: sandbox.newId(), name: "王老師", payrollCode: "", enrollmentStatus: "ACTIVE",
      enrollDate: "", leaveDate: "", note: "", createdAt: sandbox.nowIso(), updatedAt: sandbox.nowIso(),
    };
    sandbox.appendRow("Persons", teacher);

    const rows = [
      {
        rowNumber: 1, raw: {}, originalTeacherName: "王老師", substituteTeacherName: "陳老師",
        dateText: "09-02(三)", periodText: "午休", className: "1年1班", subject: "",
      },
      {
        rowNumber: 2, raw: {}, originalTeacherName: "王老師", substituteTeacherName: "陳老師",
        dateText: "09-02(三)", periodText: "第2節", className: "1年1班", subject: "數學",
      },
    ];
    const result = sandbox.api_importSubstituteRows({
      semesterId: semester.id, year: 2026, month: 9, fileName: "測試.xlsx", sheetName: "工作表1",
      sourceStaffType: "NON_BD", importedBy: "測試", rows, detectedHeaders: ["日期", "節次"],
    });
    const importedRecords = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    importedRecords.forEach((r: any) => sandbox.updateRow("SubstituteRecords", r.id, { originalTeacherId: teacher.id, updatedAt: sandbox.nowIso() }));
    sandbox.api_classifyMonthlyImport({ id: result.monthlyImport.id, changedBy: "測試" });
    sandbox.api_calculateMonthlyImportFees({ id: result.monthlyImport.id, changedBy: "測試" });

    const summary = sandbox.api_summarizeTeacherMonthlyFees({ monthlyImportIds: [result.monthlyImport.id] });
    const chenRow = summary.find((s: any) => s.substituteTeacherName.indexOf("陳老師") !== -1);
    // 只有第2節那一筆有金額，午休那一筆 amount 是 null，彙總時本來就會被排除
    // （summarizeTeacherMonthlyFees 只加總 amount 不為空的紀錄），所以總數只有 1 筆、165 元。
    expect(chenRow.generalCount).toBe(1);
    expect(chenRow.generalAmount).toBe(165);
    expect(chenRow.totalCount).toBe(1);
    expect(chenRow.totalAmount).toBe(165);
  });
});

describe("導師時間：因為不會建立 SubstituteRecord，計費流程自然也不會產生任何一般代課鐘點費", () => {
  it("匯入含「導師時間」的批次後，計費 API 找不到任何以此為節次的紀錄可以計算", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    sandbox.api_createFeeRule({
      semesterId: semester.id, feeType: "SUBSTITUTE_PERIOD", amount: 165,
      effectiveDate: "2026-08-31", changedBy: "測試",
    });

    const rows = [
      {
        rowNumber: 1, raw: {}, originalTeacherName: "王老師", substituteTeacherName: "陳老師",
        dateText: "09-02(三)", periodText: "導師時間", className: "1年1班", subject: "",
      },
    ];
    const result = sandbox.api_importSubstituteRows({
      semesterId: semester.id, year: 2026, month: 9, fileName: "測試.xlsx", sheetName: "工作表1",
      sourceStaffType: "NON_BD", importedBy: "測試", rows, detectedHeaders: ["日期", "節次"],
    });

    expect(result.successCount).toBe(0);
    const records = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(records).toHaveLength(0);

    const feeResults = sandbox.api_calculateMonthlyImportFees({ id: result.monthlyImport.id, changedBy: "測試" });
    expect(feeResults).toHaveLength(0);

    const summary = sandbox.api_summarizeTeacherMonthlyFees({ monthlyImportIds: [result.monthlyImport.id] });
    expect(summary).toHaveLength(0);
  });
});
