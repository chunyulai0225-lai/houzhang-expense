// 專案刪除功能測試——載入真正的 gas/*.gs 原始碼執行（見 tests/helpers/gasHarness.ts）。
// 重點驗證：沒被任何資料引用的專案可以真的刪除；已經被 WeeklyRules／DateRules／
// SubstituteRecords 任何一個引用過的專案一律拒絕刪除（不做 cascade delete，也不會
// 動到引用它的那些資料本身），並且仍然可以改用「停用」。
import { describe, expect, it, beforeEach } from "vitest";
import { createGasSandbox, seedRealSemester115_1 } from "./helpers/gasHarness";

describe("專案管理：刪除／停用（GAS 端，載入真正的 gas/*.gs 原始碼執行）", () => {
  let sandbox: any;
  let semester: any;

  beforeEach(() => {
    sandbox = createGasSandbox();
    semester = seedRealSemester115_1(sandbox);
  });

  it("api_listProjects 會回傳 isInUse，用來決定前端要不要顯示刪除按鈕", () => {
    const created = sandbox.api_createProject({ semesterId: semester.id, name: "測試專案", changedBy: "測試" });
    const list = sandbox.api_listProjects({ semesterId: semester.id });
    expect(list).toHaveLength(1);
    expect(list[0].id).toBe(created.id);
    expect(list[0].isInUse).toBe(false);
  });

  it("沒有被任何資料引用的專案可以刪除", () => {
    const created = sandbox.api_createProject({ semesterId: semester.id, name: "未使用專案", changedBy: "測試" });
    const result = sandbox.api_deleteProject({ id: created.id, changedBy: "測試" });
    expect(result.ok).toBe(true);
    expect(sandbox.api_listProjects({ semesterId: semester.id })).toHaveLength(0);
  });

  it("刪除會寫入 ChangeLog", () => {
    const created = sandbox.api_createProject({ semesterId: semester.id, name: "未使用專案2", changedBy: "測試" });
    const before = sandbox.readRows("ChangeLog").length;
    sandbox.api_deleteProject({ id: created.id, changedBy: "測試" });
    const after = sandbox.readRows("ChangeLog").length;
    expect(after).toBeGreaterThan(before);
    const log = sandbox.readRows("ChangeLog").find((l: any) => l.recordId === created.id && l.tableName === "projects");
    expect(log).toBeTruthy();
  });

  it("被 WeeklyRules 引用的專案不能刪除，錯誤訊息要跟使用者要求的一致", () => {
    const created = sandbox.api_createProject({ semesterId: semester.id, name: "有排課的專案", changedBy: "測試" });
    sandbox.appendRow("WeeklyRules", {
      id: sandbox.newId(), semesterId: semester.id, personId: "person-1", ruleType: "PROJECT",
      projectId: created.id, weekday: "MON", periodCode: "P1", subject: "", weeklyPeriods: 1,
      effectiveDate: "2026-08-31", endDate: "", note: "", createdAt: sandbox.nowIso(), updatedAt: sandbox.nowIso(),
    });

    expect(sandbox.api_listProjects({ semesterId: semester.id })[0].isInUse).toBe(true);
    expect(() => sandbox.api_deleteProject({ id: created.id, changedBy: "測試" })).toThrow(
      "此專案已被資料使用，無法刪除。若不再使用，請改為停用。"
    );
    // 拒絕刪除之後，專案跟引用它的那筆 WeeklyRules 都必須還在，不能被連帶動到。
    expect(sandbox.findById("Projects", created.id)).toBeTruthy();
    expect(sandbox.readRows("WeeklyRules")).toHaveLength(1);
  });

  it("被 DateRules 引用的專案不能刪除", () => {
    const created = sandbox.api_createProject({ semesterId: semester.id, name: "單日例外用的專案", changedBy: "測試" });
    sandbox.appendRow("DateRules", {
      id: sandbox.newId(), semesterId: semester.id, date: "2026-09-01", personId: "person-1",
      periodCode: "P1", originalClassificationNote: "", overrideClassification: "PROJECT", projectId: created.id,
      note: "", isCancelled: false, cancelledAt: "", cancelledBy: "", createdAt: sandbox.nowIso(), updatedAt: sandbox.nowIso(),
    });
    expect(() => sandbox.api_deleteProject({ id: created.id, changedBy: "測試" })).toThrow();
  });

  it("被 SubstituteRecords 引用的專案不能刪除", () => {
    const created = sandbox.api_createProject({ semesterId: semester.id, name: "已經有代課紀錄的專案", changedBy: "測試" });
    sandbox.appendRow("SubstituteRecords", {
      id: sandbox.newId(), rawRecordId: "", entryType: "EXCEL_IMPORT", monthlyImportId: "mi-1",
      originalTeacherId: "", substituteTeacherId: "", date: "2026-09-01", weekday: "TUE", periodCode: "P1",
      className: "", subject: "", leaveType: "", rawHoursOrDays: "", periodCount: "", staffType: "BD",
      fundingSource: "PROJECT", projectId: created.id, unitPrice: "", amount: "",
      classificationMethod: "WEEKLY_RULE", classificationRuleId: "", classifiedAt: "",
      autoFundingSource: "", autoClassificationMethod: "", autoClassificationRuleId: "", autoProjectId: "",
      conflictCandidatesJson: "", isManuallyModified: false, manualOverrideReason: "", manualOverrideAt: "",
      manualOverrideBy: "", note: "", createdBy: "", updatedBy: "", createdAt: sandbox.nowIso(), updatedAt: sandbox.nowIso(),
    });
    expect(() => sandbox.api_deleteProject({ id: created.id, changedBy: "測試" })).toThrow();
  });

  it("已被使用的專案仍然可以停用（isActive=false），不會被刪除擋住這條路", () => {
    const created = sandbox.api_createProject({ semesterId: semester.id, name: "有在用但要停用", changedBy: "測試" });
    sandbox.appendRow("WeeklyRules", {
      id: sandbox.newId(), semesterId: semester.id, personId: "person-1", ruleType: "PROJECT",
      projectId: created.id, weekday: "MON", periodCode: "P1", subject: "", weeklyPeriods: 1,
      effectiveDate: "2026-08-31", endDate: "", note: "", createdAt: sandbox.nowIso(), updatedAt: sandbox.nowIso(),
    });

    const deactivated = sandbox.api_setProjectActive({ id: created.id, isActive: false, changedBy: "測試" });
    expect(deactivated.isActive).toBe(false);
    // 停用不動歷史資料：WeeklyRules 那筆引用還在，專案本身也還查得到。
    expect(sandbox.readRows("WeeklyRules")).toHaveLength(1);
    expect(sandbox.findById("Projects", created.id)).toBeTruthy();
  });

  it("刪除不存在的專案要報錯，不會靜默成功", () => {
    expect(() => sandbox.api_deleteProject({ id: "not-a-real-id", changedBy: "測試" })).toThrow("找不到專案");
  });
});
