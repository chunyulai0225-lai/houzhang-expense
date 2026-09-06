// 學期管理功能測試——測的是「真正會部署到 Google Apps Script 的原始碼本身」
// （gas/*.gs），不是另外重寫一份邏輯來測。共用的假 Sheets 執行環境見
// tests/helpers/gasHarness.ts（跟專案管理／學校上課日曆的測試共用同一份，
// 不重複寫三次）。
//
// 這裡刻意不用真正的 Google Sheets/Apps Script 帳號——這個測試套件（`npm test`）
// 要能在任何一台開發機器上重複執行、不依賴外部帳號或網路，這是既有 Phase 1-9
// 測試套件的慣例，學期管理延用同一個慣例。「打真正的 GAS Web App URL」是額外
// 的手動驗證，不是這裡的自動化測試範圍（那樣做會對正式環境的 Google Sheets
// 寫入測試資料，不適合放進可重複執行的自動化測試）。
import { describe, expect, it, beforeEach } from "vitest";
import { createGasSandbox, seedRealSemester115_1 } from "./helpers/gasHarness";


describe("學期管理（GAS 端，載入真正的 gas/*.gs 原始碼執行）", () => {
  let sandbox: any;
  let sem115_1: any;

  beforeEach(() => {
    sandbox = createGasSandbox();
    sem115_1 = seedRealSemester115_1(sandbox);
  });

  it("api_listSemesters 沿用既有實作，可以讀到既有的 115-1", () => {
    const list = sandbox.api_listSemesters();
    expect(list).toHaveLength(1);
    expect(list[0].schoolYear).toBe(115);
    expect(list[0].term).toBe(1);
    expect(list[0].isCurrent).toBe(true);
  });

  it("新增 116 學年度第1學期", () => {
    const created = sandbox.api_createSemester({
      schoolYear: 116,
      term: 1,
      startDate: "2027-02-11",
      endDate: "2027-06-30",
      changedBy: "測試",
    });
    expect(created.id).toBeTruthy();
    expect(created.createdAt).toBeTruthy();
    expect(created.updatedAt).toBeTruthy();
    expect(created.schoolYear).toBe(116);
    expect(created.term).toBe(1);
    expect(created.status).toBe("NOT_STARTED");
    expect(created.isCurrent).toBe(false);
    // 沒指定 overtimeMatchMode 時，預設用系統目前實際在用的 SUBJECT 模式
    expect(created.overtimeMatchMode).toBe("TEACHER_WEEKDAY_PERIOD_SUBJECT");

    const all = sandbox.api_listSemesters();
    expect(all).toHaveLength(2);
  });

  it("新增 116 學年度第2學期", () => {
    sandbox.api_createSemester({ schoolYear: 116, term: 1, startDate: "2027-02-11", endDate: "2027-06-30", changedBy: "測試" });
    const created2 = sandbox.api_createSemester({
      schoolYear: 116,
      term: 2,
      startDate: "2027-08-30",
      endDate: "2028-01-19",
      changedBy: "測試",
    });
    expect(created2.schoolYear).toBe(116);
    expect(created2.term).toBe(2);
    expect(sandbox.api_listSemesters()).toHaveLength(3);
  });

  it("不能重複建立同一個學年度＋學期（比照 Node schema 的 @@unique([schoolYear, term])）", () => {
    expect(() =>
      sandbox.api_createSemester({ schoolYear: 115, term: 1, startDate: "2026-08-31", endDate: "2027-01-20", changedBy: "測試" })
    ).toThrow();
  });

  it("學期只能是第1或第2學期，其他值要拒絕", () => {
    expect(() =>
      sandbox.api_createSemester({ schoolYear: 117, term: 3, startDate: "2028-08-01", endDate: "2029-01-01", changedBy: "測試" })
    ).toThrow();
  });

  it("結束日期不能早於開始日期", () => {
    expect(() =>
      sandbox.api_createSemester({ schoolYear: 117, term: 1, startDate: "2028-08-01", endDate: "2027-01-01", changedBy: "測試" })
    ).toThrow();
  });

  it("同一時間只能有一個 isCurrent=true：設 116-1 為 current 後，115-1 自動變 false", () => {
    const created = sandbox.api_createSemester({ schoolYear: 116, term: 1, startDate: "2027-02-11", endDate: "2027-06-30", changedBy: "測試" });
    const result = sandbox.api_setCurrentSemester({ id: created.id, changedBy: "測試" });

    expect(result.isCurrent).toBe(true);
    expect(result.status).toBe("ACTIVE");

    const all = sandbox.api_listSemesters();
    const currents = all.filter((s: any) => s.isCurrent);
    expect(currents).toHaveLength(1);
    expect(currents[0].id).toBe(created.id);

    const oldCurrent = all.find((s: any) => s.id === sem115_1.id);
    expect(oldCurrent.isCurrent).toBe(false);
  });

  it("不管建立幾個學期、切換幾次 current，永遠不會同時有兩個 isCurrent=true", () => {
    const a = sandbox.api_createSemester({ schoolYear: 116, term: 1, startDate: "2027-02-11", endDate: "2027-06-30", changedBy: "測試" });
    const b = sandbox.api_createSemester({ schoolYear: 116, term: 2, startDate: "2027-08-30", endDate: "2028-01-19", changedBy: "測試" });
    sandbox.api_setCurrentSemester({ id: a.id, changedBy: "測試" });
    sandbox.api_setCurrentSemester({ id: b.id, changedBy: "測試" });
    sandbox.api_setCurrentSemester({ id: sem115_1.id, changedBy: "測試" });

    const currents = sandbox.api_listSemesters().filter((s: any) => s.isCurrent);
    expect(currents).toHaveLength(1);
    expect(currents[0].id).toBe(sem115_1.id);
  });

  it("停用目前使用中的學期後，isCurrent 必須變 false、status 變 INACTIVE", () => {
    const result = sandbox.api_deactivateSemester({ id: sem115_1.id, changedBy: "測試", reason: "測試停用" });
    expect(result.isCurrent).toBe(false);
    expect(result.status).toBe("INACTIVE");

    const currents = sandbox.api_listSemesters().filter((s: any) => s.isCurrent);
    expect(currents).toHaveLength(0);
  });

  it("停用後資料不會被刪除，仍然查得到歷史資料", () => {
    sandbox.api_deactivateSemester({ id: sem115_1.id, changedBy: "測試", reason: "測試停用" });
    const all = sandbox.api_listSemesters();
    expect(all).toHaveLength(1);
    const found = all.find((s: any) => s.id === sem115_1.id);
    expect(found).toBeTruthy();
    expect(found.status).toBe("INACTIVE");
    expect(found.schoolYear).toBe(115);
    expect(found.startDate).toBe("2026-08-31");
  });

  it("重新啟用停用的學期：狀態解除、但不會自動變回目前使用中", () => {
    sandbox.api_deactivateSemester({ id: sem115_1.id, changedBy: "測試", reason: "測試停用" });
    const reactivated = sandbox.api_activateSemester({ id: sem115_1.id, changedBy: "測試" });
    expect(reactivated.status).toBe("NOT_STARTED");
    expect(reactivated.isCurrent).toBe(false);
  });

  it("重新啟用一個本來就沒有停用的學期要報錯", () => {
    expect(() => sandbox.api_activateSemester({ id: sem115_1.id, changedBy: "測試" })).toThrow();
  });

  it("不提供刪除功能：Router 的 action 對照表裡沒有任何刪除學期的 action", () => {
    expect(sandbox.ACTIONS.deleteSemester).toBeUndefined();
    expect(sandbox.ACTIONS.removeSemester).toBeUndefined();
    // dispatch 對照表裡實際存在的 5 個學期管理 action，且都要走 LockService（不是 READ_ONLY_ACTIONS）
    expect(typeof sandbox.ACTIONS.createSemester).toBe("function");
    expect(typeof sandbox.ACTIONS.updateSemester).toBe("function");
    expect(typeof sandbox.ACTIONS.setCurrentSemester).toBe("function");
    expect(typeof sandbox.ACTIONS.deactivateSemester).toBe("function");
    expect(typeof sandbox.ACTIONS.activateSemester).toBe("function");
    expect(sandbox.READ_ONLY_ACTIONS.createSemester).toBeUndefined();
    expect(sandbox.READ_ONLY_ACTIONS.setCurrentSemester).toBeUndefined();
    expect(sandbox.READ_ONLY_ACTIONS.deactivateSemester).toBeUndefined();
  });

  it("編輯學期：可以改開始/結束日期與備註，不會動到 id/createdAt，也不會動到目前使用中狀態", () => {
    const originalCreatedAt = sem115_1.createdAt;
    const updated = sandbox.api_updateSemester({
      id: sem115_1.id,
      endDate: "2027-01-31",
      note: "延後一週結束",
      changedBy: "測試",
    });
    expect(updated.id).toBe(sem115_1.id);
    expect(updated.createdAt).toBe(originalCreatedAt);
    expect(updated.endDate).toBe("2027-01-31");
    expect(updated.note).toBe("延後一週結束");
    // 115-1 編輯前就是 current，編輯之後應該仍然是 current，不會被編輯動作取消
    expect(updated.isCurrent).toBe(true);
  });

  it("每一次新增／設為目前使用／停用／重新啟用都要寫入 ChangeLog", () => {
    const created = sandbox.api_createSemester({ schoolYear: 116, term: 1, startDate: "2027-02-11", endDate: "2027-06-30", changedBy: "測試" });
    const afterCreate = sandbox.readRows("ChangeLog").length;
    expect(afterCreate).toBeGreaterThan(0);

    sandbox.api_setCurrentSemester({ id: created.id, changedBy: "測試" });
    const afterSetCurrent = sandbox.readRows("ChangeLog").length;
    expect(afterSetCurrent).toBeGreaterThan(afterCreate);

    sandbox.api_deactivateSemester({ id: created.id, changedBy: "測試", reason: "測試" });
    const afterDeactivate = sandbox.readRows("ChangeLog").length;
    expect(afterDeactivate).toBeGreaterThan(afterSetCurrent);

    sandbox.api_activateSemester({ id: created.id, changedBy: "測試" });
    const afterActivate = sandbox.readRows("ChangeLog").length;
    expect(afterActivate).toBeGreaterThan(afterDeactivate);

    const logs = sandbox.readRows("ChangeLog").filter((l: any) => l.recordId === created.id);
    logs.forEach((l: any) => {
      expect(l.tableName).toBe("semesters");
      expect(l.changedAt).toBeTruthy();
      expect(l.changedBy).toBe("測試");
    });
  });
});
