// 效能迴歸測試：驗證「同一個 action 裡不會重複開試算表／重複找同一個分頁／
// 重複讀同一張表的表頭」這件事本身，不是測商業邏輯的正確性（正確性已經在
// gas-semester-management/gas-project-management/gas-school-calendar 測過）。
//
// 量測的是 SpreadsheetApp.openById()／getSheetByName()／getRange().getValues()／
// setValues() 的實際呼叫次數，而不是毫秒數——因為這裡的「假 Sheets」是純記憶體
// 陣列操作，跟真正 Google Sheets 網路來回的延遲完全不成比例，毫秒數量測沒有意義；
// 但呼叫次數在真正的 Apps Script 環境裡都對應到有實際延遲的遠端呼叫，呼叫次數
// 減少就直接等於真實延遲減少，這是可以誠實驗證、也可以防止未來不小心退步的指標。
import { describe, expect, it } from "vitest";
import { createGasSandbox, seedRealSemester115_1, resetPerfCounters } from "./helpers/gasHarness";

describe("效能：同一次請求內快取試算表/分頁/表頭，不重複呼叫 Sheets API", () => {
  it("getSpreadsheet() 在同一次執行裡只會真的 openById 一次，之後再呼叫幾次都不會增加", () => {
    // createGasSandbox() 內部的 setupSheets() 已經呼叫過一次 getSpreadsheet()，
    // 把 _cachedSpreadsheet 填好了，這裡的計數器已經歸零（見 gasHarness.ts）；
    // 之後不管再呼叫幾次 getSpreadsheet()，都應該沿用同一份快取，計數器要維持 0。
    const sandbox = createGasSandbox();
    sandbox.getSpreadsheet();
    sandbox.getSpreadsheet();
    sandbox.getSpreadsheet();
    expect(sandbox.__perfCounters.openByIdCalls).toBe(0);
  });

  it("getSheet() 在同一次執行裡同一個分頁只會呼叫一次 getSheetByName()", () => {
    const sandbox = createGasSandbox();
    resetPerfCounters(sandbox);
    sandbox.getSheet("Semesters");
    sandbox.getSheet("Semesters");
    sandbox.getSheet("Semesters");
    expect(sandbox.__perfCounters.getSheetByNameCalls).toBe(1);
  });

  it("新增學期：整個 action 完全不會呼叫 SpreadsheetApp.openById()（沿用同一次已經連好的試算表）", () => {
    const sandbox = createGasSandbox();
    seedRealSemester115_1(sandbox);
    resetPerfCounters(sandbox);
    sandbox.api_createSemester({ schoolYear: 116, term: 1, startDate: "2027-02-11", endDate: "2027-06-30", changedBy: "測試" });
    expect(sandbox.__perfCounters.openByIdCalls).toBe(0);
  });

  it("設為目前使用：完全不會呼叫 openById()，getSheetByName() 也不會超過分頁種類數", () => {
    const sandbox = createGasSandbox();
    const sem115 = seedRealSemester115_1(sandbox);
    const created = sandbox.api_createSemester({ schoolYear: 116, term: 1, startDate: "2027-02-11", endDate: "2027-06-30", changedBy: "測試" });
    resetPerfCounters(sandbox);
    sandbox.api_setCurrentSemester({ id: created.id, changedBy: "測試" });
    expect(sandbox.__perfCounters.openByIdCalls).toBe(0);
    // 這個 action 只會碰到 Semesters 跟 ChangeLog 兩張分頁，getSheetByName 不應該
    // 超過「碰到的分頁種類數」，不能每次讀寫都重新查一次同一張表在哪裡。
    expect(sandbox.__perfCounters.getSheetByNameCalls).toBeLessThanOrEqual(2);
  });

  it("刪除專案：完全不會呼叫 openById()", () => {
    const sandbox = createGasSandbox();
    const sem = seedRealSemester115_1(sandbox);
    const project = sandbox.api_createProject({ semesterId: sem.id, name: "測試專案", changedBy: "測試" });
    resetPerfCounters(sandbox);
    sandbox.api_deleteProject({ id: project.id, changedBy: "測試" });
    expect(sandbox.__perfCounters.openByIdCalls).toBe(0);
  });

  it("載入專案清單：不管清單裡有幾個專案，openById() 都是 0；只有第一次會真的去查 WeeklyRules/DateRules/SubstituteRecords，之後（不管專案數量再怎麼增加）完全沿用快取、不再重複查", () => {
    const sandbox = createGasSandbox();
    const sem = seedRealSemester115_1(sandbox);
    for (let i = 0; i < 10; i++) sandbox.api_createProject({ semesterId: sem.id, name: "專案" + i, changedBy: "測試" });

    resetPerfCounters(sandbox);
    sandbox.api_listProjects({ semesterId: sem.id });
    const firstCallSheetLookups = sandbox.__perfCounters.getSheetByNameCalls;
    expect(sandbox.__perfCounters.openByIdCalls).toBe(0);
    // 第一次呼叫至少要真的查過 WeeklyRules／DateRules／SubstituteRecords 這三張表
    // （修好之前，這三張表原本會被查 10 次——每個專案各查一次；修好之後只查一次）。
    expect(firstCallSheetLookups).toBeGreaterThan(0);

    // 修好之前是「每個專案都各自掃 3 張表」，呼叫次數會跟專案數量成正比；
    // 修好之後這三張表在同一次執行裡只會真的查一次，之後不管新增多少專案、
    // 再呼叫幾次 listProjects，都完全沿用快取，getSheetByName 不會再增加。
    for (let i = 10; i < 30; i++) sandbox.api_createProject({ semesterId: sem.id, name: "專案" + i, changedBy: "測試" });
    resetPerfCounters(sandbox);
    sandbox.api_listProjects({ semesterId: sem.id });
    expect(sandbox.__perfCounters.getSheetByNameCalls).toBe(0);
  });

  it("產生學校上課日曆：完全不會呼叫 openById()", () => {
    const sandbox = createGasSandbox();
    const sem = seedRealSemester115_1(sandbox);
    resetPerfCounters(sandbox);
    sandbox.api_generateSemesterCalendar({ semesterId: sem.id, changedBy: "測試" });
    expect(sandbox.__perfCounters.openByIdCalls).toBe(0);
  });
});
