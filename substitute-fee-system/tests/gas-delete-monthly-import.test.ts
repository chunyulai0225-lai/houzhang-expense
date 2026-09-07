// 開發／測試期間累積了很多測試匯入批次，需要能整批刪除，不提供單列刪除
// （RawRecord/SubstituteRecord 一律以 MonthlyImport 為刪除單位）。
//
// 這裡驗證 gas/Import.gs 新增的 api_deleteMonthlyImport()：
//   1. 刪除一個 MonthlyImport 時，一併刪除其 RawRecords／SubstituteRecords／
//      ImportErrors，以及指向這些即將被刪除資料的 IssueAcknowledgements
//      （避免留下指向不存在資料的孤兒確認紀錄）。
//   2. 完全不影響其他 MonthlyImport（不論 ACTIVE 或 SUPERSEDED）、也不影響
//      Persons／Semesters／PeriodSlots／WeeklyRules／DateRules／Projects／
//      FeeRules／MonthlyLocks 等其他資料。
//   3. 已鎖定月份的匯入批次禁止刪除（沿用既有 assertMonthNotLocked()，不是
//      另外發明一套判斷邏輯）。
//   4. 刪除之後「待處理」清單（api_listPendingIssues）不會留下這個批次的
//      孤兒問題（它本來就只查詢還存在、且 status=ACTIVE 的 MonthlyImport）。
import { describe, expect, it } from "vitest";
import { createGasSandbox, seedRealSemester115_1 } from "./helpers/gasHarness";

function importTestBatch(sandbox: any, semesterId: string, year = 2026, month = 9) {
  return sandbox.api_importSubstituteRows({
    semesterId, year, month, fileName: "測試批次.xlsx", sheetName: "工作表1",
    sourceStaffType: "NON_BD", importedBy: "測試",
    rows: [
      { rowNumber: 1, raw: {}, originalTeacherName: "王老師", substituteTeacherName: "陳老師", dateText: "09-02(三)", periodText: "第1節", className: "1年1班", subject: "國語" },
      { rowNumber: 2, raw: {}, originalTeacherName: "", substituteTeacherName: "陳老師", dateText: "09-03(四)", periodText: "第2節", className: "1年1班", subject: "" },
    ],
    detectedHeaders: ["日期", "節次"],
  });
}

describe("api_deleteMonthlyImport()：1~6. 建立測試批次（含 RawRecords/SubstituteRecords/ImportErrors）後刪除，全部清除乾淨", () => {
  it("刪除後 MonthlyImport／RawRecords／SubstituteRecords／ImportErrors／IssueAcknowledgements 全部消失", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const result = importTestBatch(sandbox, semester.id);
    const importId = result.monthlyImport.id;

    // 先確認測試資料真的建立成功，才有意義驗證刪除。
    expect(sandbox.readRows("RawRecords").filter((r: any) => r.monthlyImportId === importId)).toHaveLength(2);
    expect(sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === importId)).toHaveLength(1);
    expect(sandbox.readRows("ImportErrors").filter((e: any) => e.monthlyImportId === importId)).toHaveLength(1);

    // 對其中一筆 SubstituteRecord、一筆 ImportError 各建立一筆確認紀錄，
    // 驗證刪除時會一併清掉這些關聯的「待處理」確認資料，不留孤兒。
    const substituteRecord = sandbox.readRows("SubstituteRecords").find((r: any) => r.monthlyImportId === importId);
    const importError = sandbox.readRows("ImportErrors").find((e: any) => e.monthlyImportId === importId);
    sandbox.api_acknowledgeIssue({
      semesterId: semester.id, year: 2026, month: 9, targetTable: "SubstituteRecord", targetId: substituteRecord.id,
      reason: "測試確認", acknowledgedBy: "測試",
    });
    sandbox.api_acknowledgeIssue({
      semesterId: semester.id, year: 2026, month: 9, targetTable: "MonthlyImportError", targetId: importError.id,
      reason: "測試確認", acknowledgedBy: "測試",
    });
    expect(sandbox.readRows("IssueAcknowledgements").filter((a: any) =>
      (a.targetTable === "SubstituteRecord" && a.targetId === substituteRecord.id) ||
      (a.targetTable === "MonthlyImportError" && a.targetId === importError.id)
    )).toHaveLength(2);

    const deleteResult = sandbox.api_deleteMonthlyImport({ id: importId, changedBy: "測試" });
    expect(deleteResult.deletedRawRecords).toBe(2);
    expect(deleteResult.deletedSubstituteRecords).toBe(1);
    expect(deleteResult.deletedImportErrors).toBe(1);
    expect(deleteResult.deletedAcknowledgements).toBe(2);

    // 6. 全部清除乾淨。
    expect(sandbox.findById("MonthlyImports", importId)).toBeNull();
    expect(sandbox.readRows("RawRecords").filter((r: any) => r.monthlyImportId === importId)).toHaveLength(0);
    expect(sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === importId)).toHaveLength(0);
    expect(sandbox.readRows("ImportErrors").filter((e: any) => e.monthlyImportId === importId)).toHaveLength(0);
    expect(sandbox.readRows("IssueAcknowledgements").filter((a: any) =>
      (a.targetTable === "SubstituteRecord" && a.targetId === substituteRecord.id) ||
      (a.targetTable === "MonthlyImportError" && a.targetId === importError.id)
    )).toHaveLength(0);
  });
});

describe("7. 其他 MonthlyImport 不受影響（不論 ACTIVE 或 SUPERSEDED）", () => {
  it("刪除批次 A 之後，批次 B（不同月份）與批次 A 自己重新匯入後產生的 SUPERSEDED 批次都完整保留", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);

    const batchOther = importTestBatch(sandbox, semester.id, 2026, 10); // 不同月份，完全獨立的批次
    const batchFirstVersion = importTestBatch(sandbox, semester.id, 2026, 9); // 之後會被同月份重新匯入取代成 SUPERSEDED
    const batchSecondVersion = importTestBatch(sandbox, semester.id, 2026, 9); // 重新匯入，把上面那筆變成 SUPERSEDED

    const supersededBatch = sandbox.findById("MonthlyImports", batchFirstVersion.monthlyImport.id);
    expect(supersededBatch.status).toBe("SUPERSEDED");

    // 刪除「最新這筆 2026/09 批次」，不應該動到批次 batchOther、也不應該動到
    // 已經是 SUPERSEDED 的舊版本批次。
    sandbox.api_deleteMonthlyImport({ id: batchSecondVersion.monthlyImport.id, changedBy: "測試" });

    expect(sandbox.findById("MonthlyImports", batchOther.monthlyImport.id)).toBeTruthy();
    expect(sandbox.readRows("RawRecords").filter((r: any) => r.monthlyImportId === batchOther.monthlyImport.id)).toHaveLength(2);
    expect(sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === batchOther.monthlyImport.id)).toHaveLength(1);

    expect(sandbox.findById("MonthlyImports", batchFirstVersion.monthlyImport.id)).toBeTruthy();
    expect(sandbox.readRows("RawRecords").filter((r: any) => r.monthlyImportId === batchFirstVersion.monthlyImport.id)).toHaveLength(2);
    expect(sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === batchFirstVersion.monthlyImport.id)).toHaveLength(1);
  });

  it("也可以單獨刪除一筆已經是 SUPERSEDED 的舊測試批次，不影響取代它的 ACTIVE 批次", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const firstVersion = importTestBatch(sandbox, semester.id, 2026, 9);
    const secondVersion = importTestBatch(sandbox, semester.id, 2026, 9);

    sandbox.api_deleteMonthlyImport({ id: firstVersion.monthlyImport.id, changedBy: "測試" });

    expect(sandbox.findById("MonthlyImports", firstVersion.monthlyImport.id)).toBeNull();
    const stillActive = sandbox.findById("MonthlyImports", secondVersion.monthlyImport.id);
    expect(stillActive).toBeTruthy();
    expect(stillActive.status).toBe("ACTIVE");
    expect(sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === secondVersion.monthlyImport.id)).toHaveLength(1);
  });
});

describe("8. 其他 Persons / Semesters / Rules 不受影響", () => {
  it("刪除匯入批次不會動到 Persons、Semesters、PeriodSlots、WeeklyRules、DateRules、Projects、FeeRules、MonthlyLocks", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const person = { id: sandbox.newId(), name: "王老師", payrollCode: "", enrollmentStatus: "ACTIVE", enrollDate: "", leaveDate: "", note: "", createdAt: sandbox.nowIso(), updatedAt: sandbox.nowIso() };
    sandbox.appendRow("Persons", person);
    const project = sandbox.api_createProject({ semesterId: semester.id, name: "測試專案", changedBy: "測試" });
    sandbox.api_createWeeklyRule({
      semesterId: semester.id, personId: person.id, ruleType: "OVERTIME", weekday: "MON", periodCode: "P1",
      effectiveDate: "2026-08-31", changedBy: "測試",
    });
    sandbox.api_createFeeRule({ semesterId: semester.id, feeType: "SUBSTITUTE_PERIOD", amount: 165, effectiveDate: "2026-08-31", changedBy: "測試" });

    const beforeCounts = {
      persons: sandbox.readRows("Persons").length,
      semesters: sandbox.readRows("Semesters").length,
      periodSlots: sandbox.readRows("PeriodSlots").length,
      weeklyRules: sandbox.readRows("WeeklyRules").length,
      projects: sandbox.readRows("Projects").length,
      feeRules: sandbox.readRows("FeeRules").length,
      monthlyLocks: sandbox.readRows("MonthlyLocks").length,
    };

    const result = importTestBatch(sandbox, semester.id);
    sandbox.api_deleteMonthlyImport({ id: result.monthlyImport.id, changedBy: "測試" });

    expect(sandbox.readRows("Persons").length).toBe(beforeCounts.persons);
    expect(sandbox.readRows("Semesters").length).toBe(beforeCounts.semesters);
    expect(sandbox.readRows("PeriodSlots").length).toBe(beforeCounts.periodSlots);
    expect(sandbox.readRows("WeeklyRules").length).toBe(beforeCounts.weeklyRules);
    expect(sandbox.readRows("Projects").length).toBe(beforeCounts.projects);
    expect(sandbox.readRows("FeeRules").length).toBe(beforeCounts.feeRules);
    expect(sandbox.readRows("MonthlyLocks").length).toBe(beforeCounts.monthlyLocks);
    expect(sandbox.findById("Persons", person.id)).toBeTruthy();
    expect(sandbox.findById("Projects", project.id)).toBeTruthy();
  });
});

describe("9. 已鎖定月份不能刪除", () => {
  it("該匯入批次所屬年月已鎖定時，api_deleteMonthlyImport 拋出錯誤，資料完全保留", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    // 只匯入一筆完全正常、不會產生任何阻擋性問題的資料，才鎖得起來。
    const result = sandbox.api_importSubstituteRows({
      semesterId: semester.id, year: 2026, month: 9, fileName: "正式資料.xlsx", sheetName: "工作表1",
      sourceStaffType: "NON_BD", importedBy: "測試",
      rows: [{ rowNumber: 1, raw: {}, originalTeacherName: "王老師", substituteTeacherName: "陳老師", dateText: "09-02(三)", periodText: "第1節", className: "1年1班", subject: "國語" }],
      detectedHeaders: ["日期", "節次"],
    });
    sandbox.api_lockMonth({ semesterId: semester.id, year: 2026, month: 9, lockedBy: "測試" });

    expect(() => sandbox.api_deleteMonthlyImport({ id: result.monthlyImport.id, changedBy: "測試" })).toThrow(/鎖定/);

    // 資料完全沒被刪掉。
    expect(sandbox.findById("MonthlyImports", result.monthlyImport.id)).toBeTruthy();
    expect(sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id)).toHaveLength(1);
  });
});

describe("10. 待處理清單不會留下該批次的孤兒問題", () => {
  it("刪除有問題的批次後，api_listPendingIssues 完全查不到屬於這個批次的任何問題", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const result = importTestBatch(sandbox, semester.id);

    const beforeIssues = sandbox.api_listPendingIssues({ semesterId: semester.id, year: 2026, month: 9 });
    expect(beforeIssues.length).toBeGreaterThan(0);

    sandbox.api_deleteMonthlyImport({ id: result.monthlyImport.id, changedBy: "測試" });

    const afterIssues = sandbox.api_listPendingIssues({ semesterId: semester.id, year: 2026, month: 9 });
    expect(afterIssues).toHaveLength(0);
  });
});
