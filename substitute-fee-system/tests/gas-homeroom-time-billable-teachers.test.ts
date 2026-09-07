// 后庄國小補充確認的業務規則：一般老師的「導師時間」（08:15~08:45）不計一般代課
// 鐘點費，但「培力班」6 位老師的導師時間代課例外，需要正常計算一般代課鐘點費：
//   陳志平、張珊華、林慧玲、林芳燕、林芳竹、游雅筑
//
// 重要：「導師時間是否計費」要依「原教師」判斷，不是依代課教師判斷。
//
// 實作方式（gas/FeeCalculation.gs）：不新增資料表或欄位，PeriodSlots.isTeachingPeriod
// 對 HOMEROOM_TIME 維持全校預設 false（一般老師仍然不計費）；改為新增一個只依姓名
// 判斷的例外清單 isHomeroomTimeBillableForOriginalTeacher()，查的是既有 Persons 資料，
// 針對「原教師」而非代課教師。這裡驗證：
//   1. 一般老師 + 導師時間 → 不計一般鐘點
//   2~7. 培力班 6 人（分別）+ 導師時間 → 可以計一般鐘點
//   8. 培力班 6 人 + 午休 → 仍然不計（跟原教師是誰無關）
//   9~11. 見 tests/gas-date-range-period-count.test.ts 涵蓋日期區間/時數天數案例，
//        這裡只再次確認 P1~P7 的既有計費完全不受影響
//   12. P1~P7 原本計費測試仍全部通過（見下面「維持 Phase 9-5」區塊）
import { describe, expect, it } from "vitest";
import { createGasSandbox, seedRealSemester115_1 } from "./helpers/gasHarness";

const HOMEROOM_TIME_BILLABLE_NAMES = ["陳志平", "張珊華", "林慧玲", "林芳燕", "林芳竹", "游雅筑"];

function setupTeacherAndFeeRule(sandbox: any, semesterId: string, teacherName: string) {
  sandbox.api_createFeeRule({
    semesterId, feeType: "SUBSTITUTE_PERIOD", amount: 165, effectiveDate: "2026-08-31", changedBy: "測試",
  });
  const teacher = {
    id: sandbox.newId(), name: teacherName, payrollCode: "", enrollmentStatus: "ACTIVE",
    enrollDate: "", leaveDate: "", note: "", createdAt: sandbox.nowIso(), updatedAt: sandbox.nowIso(),
  };
  sandbox.appendRow("Persons", teacher);
  return teacher;
}

function importAndCalculate(sandbox: any, semesterId: string, teacherId: string, periodText: string) {
  const result = sandbox.api_importSubstituteRows({
    semesterId, year: 2026, month: 9, fileName: "測試.xlsx", sheetName: "工作表1",
    sourceStaffType: "NON_BD", importedBy: "測試",
    rows: [{
      rowNumber: 1, raw: {}, originalTeacherName: "原教師", substituteTeacherName: "代課教師",
      dateText: "09-02(三)", periodText, className: "1年1班", subject: "",
    }],
    detectedHeaders: ["日期", "節次"],
  });
  const record = sandbox.readRows("SubstituteRecords").find((r: any) => r.monthlyImportId === result.monthlyImport.id);
  sandbox.updateRow("SubstituteRecords", record.id, { originalTeacherId: teacherId, updatedAt: sandbox.nowIso() });
  sandbox.api_classifyMonthlyImport({ id: result.monthlyImport.id, changedBy: "測試" });
  const feeResult = sandbox.api_calculateSubstituteRecordFee({ id: record.id, changedBy: "測試" });
  return { result, record, feeResult };
}

describe("isHomeroomTimeBillableForOriginalTeacher()：只依「原教師」姓名判斷，不依代課教師", () => {
  it("培力班 6 位老師都回傳 true", () => {
    const sandbox = createGasSandbox();
    HOMEROOM_TIME_BILLABLE_NAMES.forEach((name) => {
      const person = { id: sandbox.newId(), name, payrollCode: "", enrollmentStatus: "ACTIVE", enrollDate: "", leaveDate: "", note: "", createdAt: sandbox.nowIso(), updatedAt: sandbox.nowIso() };
      sandbox.appendRow("Persons", person);
      expect(sandbox.isHomeroomTimeBillableForOriginalTeacher(person.id)).toBe(true);
    });
  });

  it("一般老師（不在名單內）回傳 false；空值／不存在的 id 也回傳 false，不擋下既有流程", () => {
    const sandbox = createGasSandbox();
    const person = { id: sandbox.newId(), name: "王老師", payrollCode: "", enrollmentStatus: "ACTIVE", enrollDate: "", leaveDate: "", note: "", createdAt: sandbox.nowIso(), updatedAt: sandbox.nowIso() };
    sandbox.appendRow("Persons", person);
    expect(sandbox.isHomeroomTimeBillableForOriginalTeacher(person.id)).toBe(false);
    expect(sandbox.isHomeroomTimeBillableForOriginalTeacher("")).toBe(false);
    expect(sandbox.isHomeroomTimeBillableForOriginalTeacher(null)).toBe(false);
    expect(sandbox.isHomeroomTimeBillableForOriginalTeacher("not-a-real-id")).toBe(false);
  });

  it("PeriodSlots.HOMEROOM_TIME 的 isTeachingPeriod 全校預設仍是 false，沒有被改成 true", () => {
    const sandbox = createGasSandbox();
    const slot = sandbox.findOne("PeriodSlots", (p: any) => p.code === "HOMEROOM_TIME");
    expect(slot).toBeTruthy();
    expect(sandbox.toBool(slot.isTeachingPeriod)).toBe(false);
    expect(sandbox.isNonPayablePeriodCode("HOMEROOM_TIME")).toBe(true);
  });
});

describe("1. 一般老師 + 導師時間 → 不計一般鐘點", () => {
  it("原教師是一般老師（不在培力班名單）：amount 維持 null", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const teacher = setupTeacherAndFeeRule(sandbox, semester.id, "王老師");
    const { feeResult } = importAndCalculate(sandbox, semester.id, teacher.id, "導師時間");
    expect(feeResult.amount).toBeNull();
    expect(feeResult.skippedReason).toContain("非授課節次");
  });
});

describe("2~7. 培力班 6 位老師 + 導師時間 → 可以計一般鐘點", () => {
  HOMEROOM_TIME_BILLABLE_NAMES.forEach((teacherName) => {
    it(`${teacherName} + 導師時間 → 正常算出一般代課鐘點費`, () => {
      const sandbox = createGasSandbox();
      const semester = seedRealSemester115_1(sandbox);
      const teacher = setupTeacherAndFeeRule(sandbox, semester.id, teacherName);
      const { record, feeResult } = importAndCalculate(sandbox, semester.id, teacher.id, "導師時間");

      expect(record.periodCode).toBe("HOMEROOM_TIME");
      expect(feeResult.amount).toBe("165"); // 沿用 FeeRule 費率，跟 P1~P7 走同一套 Phase9-5 計算邏輯
      expect(feeResult.unitPrice).toBe("165");
      expect(feeResult.skippedReason).toBeUndefined();
    });
  });
});

describe("8. 培力班 6 人 + 午休 → 不計（跟原教師是誰無關）", () => {
  HOMEROOM_TIME_BILLABLE_NAMES.forEach((teacherName) => {
    it(`${teacherName} + 午休 → 仍然不計一般代課鐘點費`, () => {
      const sandbox = createGasSandbox();
      const semester = seedRealSemester115_1(sandbox);
      const teacher = setupTeacherAndFeeRule(sandbox, semester.id, teacherName);
      const { feeResult } = importAndCalculate(sandbox, semester.id, teacher.id, "午休");

      expect(feeResult.amount).toBeNull();
      expect(feeResult.skippedReason).toContain("非授課節次");
    });
  });
});

describe("9. 劉馨憶 + 09-07~09-11 + P2 + 5日 → periodCount=5（不因原教師是誰而改變日期區間規則）", () => {
  it("正常建立紀錄，periodCount 依時數天數＝5，不從日期區間展開或重算", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const result = sandbox.api_importSubstituteRows({
      semesterId: semester.id, year: 2026, month: 9, fileName: "測試.xlsx", sheetName: "工作表1",
      sourceStaffType: "NON_BD", importedBy: "測試",
      rows: [{
        rowNumber: 1, raw: {}, originalTeacherName: "劉馨憶", substituteTeacherName: "徐碧苓",
        dateText: "09-07(一) 07:50 ~ 09-11(五) 15:50", periodText: "第2節", hoursOrDaysText: "5日",
        className: "1年1班", subject: "",
      }],
      detectedHeaders: ["日期", "節次", "時數天數"],
    });
    expect(result.errorCount).toBe(0);
    expect(result.successCount).toBe(1);
    const records = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(records).toHaveLength(1);
    expect(records[0].periodCode).toBe("P2");
    expect(Number(records[0].periodCount)).toBe(5);
    expect(records[0].date).toBe("2026-09-07"); // 區間起始日僅作定位點，不是重算出來的代課日
  });
});

describe("10. 日期區間本身不再產生 ImportError", () => {
  it("三組真實日期區間（09-07~09-11／09-11~09-14／09-21~09-24）搭配合法時數天數，都不會有任何 ImportError", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const rows = [
      { rowNumber: 1, raw: {}, originalTeacherName: "劉馨憶", substituteTeacherName: "徐碧苓", dateText: "09-07(一) 07:50 ~ 09-11(五) 15:50", periodText: "第2節", hoursOrDaysText: "5日", className: "1年1班", subject: "" },
      { rowNumber: 2, raw: {}, originalTeacherName: "詹庭瑜", substituteTeacherName: "陳老師", dateText: "09-11(五) 07:50 ~ 09-14(一) 15:50", periodText: "第1節", hoursOrDaysText: "3日", className: "1年1班", subject: "" },
      { rowNumber: 3, raw: {}, originalTeacherName: "詹庭瑜", substituteTeacherName: "陳老師", dateText: "09-21(一) 07:50 ~ 09-24(四) 15:50", periodText: "第1節", hoursOrDaysText: "4日", className: "1年1班", subject: "" },
    ];
    const result = sandbox.api_importSubstituteRows({
      semesterId: semester.id, year: 2026, month: 9, fileName: "測試.xlsx", sheetName: "工作表1",
      sourceStaffType: "NON_BD", importedBy: "測試", rows, detectedHeaders: ["日期", "節次", "時數天數"],
    });
    expect(result.errorCount).toBe(0);
    expect(result.successCount).toBe(3);
    const messages = result.errors.map((e: any) => e.message);
    expect(messages.some((m: string) => m.includes("日期區間"))).toBe(false);
  });
});

describe("11. 3日4時仍然標記為時數天數待確認", () => {
  it("林家德 + 09-11~09-14 + 3日4時 + 第六節 → 標記「時數天數／待確認」，不建立 SubstituteRecord", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const result = sandbox.api_importSubstituteRows({
      semesterId: semester.id, year: 2026, month: 9, fileName: "測試.xlsx", sheetName: "工作表1",
      sourceStaffType: "NON_BD", importedBy: "測試",
      rows: [{
        rowNumber: 1, raw: {}, originalTeacherName: "林家德", substituteTeacherName: "陳老師",
        dateText: "09-11(五) 07:50 ~ 09-14(一) 15:50", periodText: "第6節", hoursOrDaysText: "3日4時",
        className: "3年1班", subject: "",
      }],
      detectedHeaders: ["日期", "節次", "時數天數"],
    });
    expect(result.successCount).toBe(0);
    expect(result.errorCount).toBe(1);
    expect(result.errors[0].message).toContain("時數天數／待確認");
    expect(result.errors[0].message).toContain("3日4時");
    expect(result.errors[0].message).not.toContain("日期區間／待確認");
    const records = sandbox.readRows("SubstituteRecords").filter((r: any) => r.monthlyImportId === result.monthlyImport.id);
    expect(records).toHaveLength(0);
  });
});

describe("12. P1~P7 原本計費測試仍全部通過（Phase 9-5 完全未被這次修改影響）", () => {
  it("一般老師的 P1~P7 正常計費，跟培力班例外互不影響", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const teacher = setupTeacherAndFeeRule(sandbox, semester.id, "一般老師甲");
    for (let i = 1; i <= 7; i++) {
      const { feeResult } = importAndCalculate(sandbox, semester.id, teacher.id, "第" + i + "節");
      expect(feeResult.amount).toBe("165");
      expect(feeResult.skippedReason).toBeUndefined();
    }
  });

  it("培力班老師本人的 P1~P7 也正常計費（例外只針對導師時間，不影響一般節次）", () => {
    const sandbox = createGasSandbox();
    const semester = seedRealSemester115_1(sandbox);
    const teacher = setupTeacherAndFeeRule(sandbox, semester.id, "陳志平");
    const { feeResult } = importAndCalculate(sandbox, semester.id, teacher.id, "第3節");
    expect(feeResult.amount).toBe("165");
    expect(feeResult.skippedReason).toBeUndefined();
  });
});
