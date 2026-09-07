// 「待處理／問題清單」可追蹤性補強：使用者反應目前的錯誤畫面只顯示「列號｜欄位｜
// 問題」（例如「34｜時數天數｜時數天數／待確認……」），完全看不出是哪一位老師。
//
// 這裡驗證 gas/MonthlyClose.gs 的 api_listPendingIssues() 修正：每一筆問題現在都
// 附上 rowNumber／原教師／代課教師／日期原文／節次原文（或解析後節次）／時數天數
// 原文／問題欄位／問題說明／rawRecordId（可追溯回 RawRecord），四種 issueType
// （TEACHER_UNMATCHED／CONFLICT／AMOUNT_MISSING／IMPORT_ERROR）共用同一組欄位，
// 不再是 IMPORT_ERROR 特別稀疏。同時新增 api_summarizePendingIssuesByOriginalTeacher()
// 供「待處理」頁面做依原教師的彙總／篩選，以及 api_listPendingIssues() 本身支援
// originalTeacher／substituteTeacher／issueType／periodCode 篩選。
//
// 不新增資料表：IMPORT_ERROR 的原教師/代課教師/日期/節次/時數天數是用「同一個
// monthlyImportId + 同一個 rowNumber」去對照既有的 RawRecords 補上的，AMOUNT_MISSING
// 也沿用既有 SubstituteRecords.note 欄位判斷是不是「時數天數待確認」造成的。
//
// 同時確認上一輪再次修正的規則：日期區間、以及「3日4時」這種無法安全解析的時數
// 天數，都不會再造成 ImportError、也不會讓 RawRecord/SubstituteRecord 消失。
import { describe, expect, it } from "vitest";
import { createGasSandbox, seedRealSemester115_1 } from "./helpers/gasHarness";

function importBatch(sandbox: any, semesterId: string, rows: any[], year = 2026, month = 9) {
  return sandbox.api_importSubstituteRows({
    semesterId, year, month, fileName: "測試.xlsx", sheetName: "工作表1",
    sourceStaffType: "NON_BD", importedBy: "測試", rows, detectedHeaders: ["日期", "節次", "時數天數"],
  });
}

describe("1~5. 問題清單（IMPORT_ERROR）可以顯示原教師／代課教師／節次／日期原文／時數天數原文", () => {
  it("第34列風格範例：日期區間 + 第三節 + 時數天數缺漏（未提供），仍能顯示完整的老師與節次資訊", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const rows = [];
    // 先塞 32 筆正常列，湊出「第 34 列」的情境（不影響驗證邏輯，只是貼近使用者原本
    // 舉的例子：34｜時數天數｜時數天數／待確認）。
    for (let i = 1; i <= 32; i++) {
      rows.push({ rowNumber: i, raw: {}, originalTeacherName: "陪襯老師" + i, substituteTeacherName: "陪襯代課" + i, dateText: "09-02(三)", periodText: "第1節", className: "1年1班", subject: "" });
    }
    rows.push({ rowNumber: 33, raw: {}, originalTeacherName: "王小明", substituteTeacherName: "林大華", dateText: "09-21(一) 07:50 ~ 09-24(四) 15:50", periodText: "第3節", className: "3年2班", subject: "自然" });
    // 第34列缺少「原教師」欄位，讓 ImportErrors 真的產生一筆（缺少原教師是既有的
    // 必填檢查，不是這次修改的一部分），用來驗證 IMPORT_ERROR 這一類能不能顯示
    // 節次/日期/代課教師（即使原教師本身缺漏，其餘欄位仍要顯示）。
    rows.push({ rowNumber: 34, raw: {}, originalTeacherName: "", substituteTeacherName: "林大華", dateText: "09-21(一) 07:50 ~ 09-24(四) 15:50", periodText: "第3節", hoursOrDaysText: "", className: "3年2班", subject: "自然" });

    const result = importBatch(sandbox, semester.id, rows);
    expect(result.errorCount).toBe(1);

    const issues = sandbox.api_listPendingIssues({ semesterId: semester.id, year: 2026, month: 9 });
    const row34 = issues.find((i: any) => i.issueType === "IMPORT_ERROR" && i.rowNumber === 34);
    expect(row34).toBeTruthy();

    // 1. 原教師姓名欄位存在（這筆本身就缺原教師，值是 null，不是完全沒有這個欄位）
    expect(row34).toHaveProperty("originalTeacher");
    expect(row34.originalTeacher).toBeNull();
    // 2. 代課教師姓名
    expect(row34.substituteTeacher).toBe("林大華");
    // 3. 節次（解析後）
    expect(row34.periodCode).toBe("P3");
    expect(row34.periodText).toBe("第3節");
    // 4. 日期原始文字
    expect(row34.dateText).toBe("09-21(一) 07:50 ~ 09-24(四) 15:50");
    // 5. 時數天數原始文字欄位存在（這筆本身是空白，值是 null，不是完全沒有這個欄位）
    expect(row34).toHaveProperty("hoursOrDaysText");
    expect(row34.hoursOrDaysText).toBeNull();
    // 問題欄位／問題說明／可追溯回 RawRecord
    expect(row34.fieldName).toBe("原教師");
    expect(row34.description).toContain("原教師");
    expect(row34.rawRecordId).toBeTruthy();
    expect(row34.rowNumber).toBe(34);
  });

  it("真實範例：詹庭瑜 09-21~09-24、第3節、時數天數空白 → AMOUNT_MISSING 類型也能顯示完整節次/日期/時數天數，並清楚標示是時數天數的問題", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const teacher = { id: sandbox.newId(), name: "詹庭瑜", payrollCode: "", enrollmentStatus: "ACTIVE", enrollDate: "", leaveDate: "", note: "", createdAt: sandbox.nowIso(), updatedAt: sandbox.nowIso() };
    sandbox.appendRow("Persons", teacher);

    const result = importBatch(sandbox, semester.id, [{
      rowNumber: 1, raw: {}, originalTeacherName: "詹庭瑜", substituteTeacherName: "林○○",
      dateText: "09-21(一) 07:50 ~ 09-24(四) 15:50", periodText: "第3節", hoursOrDaysText: "3日4時",
      className: "3年2班", subject: "自然",
    }]);
    expect(result.errorCount).toBe(0);
    expect(result.successCount).toBe(1);

    const records = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    sandbox.updateRow("SubstituteRecords", records[0].id, { originalTeacherId: teacher.id, updatedAt: sandbox.nowIso() });
    sandbox.api_classifyMonthlyImport({ id: result.monthlyImport.id, changedBy: "測試" });
    sandbox.api_calculateMonthlyImportFees({ id: result.monthlyImport.id, changedBy: "測試" });

    const issues = sandbox.api_listPendingIssues({ semesterId: semester.id, year: 2026, month: 9 });
    const issue = issues.find((i: any) => i.issueType === "AMOUNT_MISSING");
    expect(issue).toBeTruthy();
    expect(issue.originalTeacher).toBe("詹庭瑜");
    expect(issue.substituteTeacher).toBe("林○○");
    expect(issue.periodCode).toBe("P3");
    expect(issue.dateText).toBe("09-21(一) 07:50 ~ 09-24(四) 15:50");
    expect(issue.hoursOrDaysText).toBe("3日4時");
    expect(issue.fieldName).toBe("時數天數");
    expect(issue.description).toContain("時數天數");
    expect(issue.rawRecordId).toBeTruthy();
  });
});

describe("6. 可以依原教師篩選", () => {
  it("api_listPendingIssues 傳入 originalTeacher 篩選條件，只回傳該老師的問題", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    importBatch(sandbox, semester.id, [
      { rowNumber: 1, raw: {}, originalTeacherName: "", substituteTeacherName: "陳老師", dateText: "09-02(三)", periodText: "第1節", className: "", subject: "" },
      { rowNumber: 2, raw: {}, originalTeacherName: "王小明", substituteTeacherName: "陳老師", dateText: "09-03(四)", periodText: "第2,3節", className: "", subject: "" },
    ]);
    const all = sandbox.api_listPendingIssues({ semesterId: semester.id, year: 2026, month: 9 });
    expect(all.length).toBeGreaterThanOrEqual(2);

    const filtered = sandbox.api_listPendingIssues({ semesterId: semester.id, year: 2026, month: 9, originalTeacher: "王小明" });
    expect(filtered.length).toBeGreaterThan(0);
    expect(filtered.every((i: any) => i.originalTeacher === "王小明")).toBe(true);
    expect(filtered.length).toBeLessThan(all.length);
  });
});

describe("7. 同一位老師多筆問題會正確彙總", () => {
  it("api_summarizePendingIssuesByOriginalTeacher 依原教師分組計數，跟畫面上「王○○：8筆」的呈現一致", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const rows = [];
    // 王小明缺原教師配對用的節次錯誤 x3（用不同缺漏欄位製造 3 筆 IMPORT_ERROR，
    // 但都屬於同一個「原教師：王小明」）。
    for (let i = 1; i <= 3; i++) {
      rows.push({ rowNumber: i, raw: {}, originalTeacherName: "王小明", substituteTeacherName: "陳老師", dateText: "09-0" + i + "(三)", periodText: "不合法節次" + i, className: "", subject: "" });
    }
    rows.push({ rowNumber: 4, raw: {}, originalTeacherName: "林小華", substituteTeacherName: "陳老師", dateText: "09-05(五)", periodText: "不合法節次4", className: "", subject: "" });
    importBatch(sandbox, semester.id, rows);

    const summary = sandbox.api_summarizePendingIssuesByOriginalTeacher({ semesterId: semester.id, year: 2026, month: 9 });
    const wang = summary.find((s: any) => s.originalTeacher === "王小明");
    const lin = summary.find((s: any) => s.originalTeacher === "林小華");
    expect(wang.count).toBe(3);
    expect(lin.count).toBe(1);
    // 依筆數由多到少排序，方便畫面直接依序渲染成清單。
    expect(summary[0].originalTeacher).toBe("王小明");
  });
});

describe("8. 不會因日期區間而產生問題", () => {
  it("日期區間 + 明確節次 + 可解析時數天數：完全不會出現在待處理清單裡", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const result = importBatch(sandbox, semester.id, [{
      rowNumber: 1, raw: {}, originalTeacherName: "劉馨憶", substituteTeacherName: "徐碧苓",
      dateText: "09-07(一) 07:50 ~ 09-11(五) 15:50", periodText: "第2節", hoursOrDaysText: "5日",
      className: "1年1班", subject: "",
    }]);
    expect(result.errorCount).toBe(0);
    const issues = sandbox.api_listPendingIssues({ semesterId: semester.id, year: 2026, month: 9 });
    expect(issues.filter((i: any) => i.issueType === "IMPORT_ERROR")).toHaveLength(0);
    expect(issues.some((i: any) => (i.description || "").includes("日期區間"))).toBe(false);
  });
});

describe("9. 不會因「3日4時」而讓原始資料消失", () => {
  it("RawRecord、SubstituteRecord 都完整保留，問題清單能追溯回 RawRecord 並看到原始時數天數文字", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const teacher = { id: sandbox.newId(), name: "林家德", payrollCode: "", enrollmentStatus: "ACTIVE", enrollDate: "", leaveDate: "", note: "", createdAt: sandbox.nowIso(), updatedAt: sandbox.nowIso() };
    sandbox.appendRow("Persons", teacher);

    const result = importBatch(sandbox, semester.id, [{
      rowNumber: 1, raw: {}, originalTeacherName: "林家德", substituteTeacherName: "陳老師",
      dateText: "09-11(五) 07:50 ~ 09-14(一) 15:50", periodText: "第6節", hoursOrDaysText: "3日4時",
      className: "3年1班", subject: "",
    }]);
    expect(result.errorCount).toBe(0);
    expect(result.successCount).toBe(1);

    const raw = sandbox.readRows("RawRecords").find((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(raw).toBeTruthy();
    expect(raw.hoursOrDaysText).toBe("3日4時");

    const records = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(records).toHaveLength(1);
    sandbox.updateRow("SubstituteRecords", records[0].id, { originalTeacherId: teacher.id, updatedAt: sandbox.nowIso() });
    sandbox.api_classifyMonthlyImport({ id: result.monthlyImport.id, changedBy: "測試" });
    sandbox.api_calculateMonthlyImportFees({ id: result.monthlyImport.id, changedBy: "測試" });

    const issues = sandbox.api_listPendingIssues({ semesterId: semester.id, year: 2026, month: 9 });
    const issue = issues.find((i: any) => i.rawRecordId === raw.id);
    expect(issue).toBeTruthy();
    expect(issue.originalTeacher).toBe("林家德");
    expect(issue.hoursOrDaysText).toBe("3日4時");
    expect(issue.issueType).toBe("AMOUNT_MISSING");
    expect(issue.fieldName).toBe("時數天數");
  });
});

describe("篩選：代課教師／問題類型／節次", () => {
  it("substituteTeacher／issueType／periodCode 三種篩選條件都能正確縮小結果", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    importBatch(sandbox, semester.id, [
      { rowNumber: 1, raw: {}, originalTeacherName: "王小明", substituteTeacherName: "陳老師", dateText: "09-02(三)", periodText: "不合法A", className: "", subject: "" },
      { rowNumber: 2, raw: {}, originalTeacherName: "王小明", substituteTeacherName: "林老師", dateText: "09-03(四)", periodText: "不合法B", className: "", subject: "" },
    ]);
    const bySub = sandbox.api_listPendingIssues({ semesterId: semester.id, year: 2026, month: 9, substituteTeacher: "陳老師" });
    expect(bySub).toHaveLength(1);
    expect(bySub[0].substituteTeacher).toBe("陳老師");

    const byType = sandbox.api_listPendingIssues({ semesterId: semester.id, year: 2026, month: 9, issueType: "IMPORT_ERROR" });
    expect(byType.length).toBeGreaterThan(0);
    expect(byType.every((i: any) => i.issueType === "IMPORT_ERROR")).toBe(true);

    // periodCode 篩選：另外匯入一筆能正常解析出 P1 的資料，確認篩選只留下 P1。
    importBatch(sandbox, semester.id, [
      { rowNumber: 3, raw: {}, originalTeacherName: "", substituteTeacherName: "陳老師", dateText: "09-05(五)", periodText: "第1節", className: "", subject: "" },
    ], 2026, 10);
    const byPeriod = sandbox.api_listPendingIssues({ semesterId: semester.id, year: 2026, month: 10, periodCode: "P1" });
    expect(byPeriod.length).toBeGreaterThan(0);
    expect(byPeriod.every((i: any) => i.periodCode === "P1")).toBe(true);
  });
});
