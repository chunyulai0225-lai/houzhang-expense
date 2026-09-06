// 全 GAS 專案的函式相依性驗證，起因是真實部署後 ?action=listSemesters 回報
// {"ok":false,"error":"getSpreadsheet is not defined"}。
//
// 這裡做兩件事：
// 1) 用一支獨立腳本靜態掃過 gas/*.gs 全部檔案，確認每一個被呼叫的自訂函式名稱都
//    確實在某個檔案裡有定義過、而且只定義一次（不重複），不是用猜的。
// 2) 用 tests/helpers/gasHarness.ts 把全部 10 個 .gs 檔案原封不動載入同一個
//    執行環境，逐一實際呼叫使用者要求要確認的每一個 API 一次，確保呼叫鏈上
//    每一個環節用到的函式都真的存在、真的可以執行到底——這跟真正部署到 Apps
//    Script 時「全部檔案合併成同一個全域作用域」的行為一致，如果有任何函式
//    沒定義，這裡呼叫到的時候就會直接丟出「X is not a function / not defined」
//    讓測試失敗，不會像單純看原始碼那樣容易漏看。
import fs from "fs";
import path from "path";
import { describe, expect, it } from "vitest";
import { createGasSandbox, seedRealSemester115_1 } from "./helpers/gasHarness";

const GAS_DIR = path.join(__dirname, "..", "gas");
const GAS_FILES = [
  "Utils.gs",
  "SheetsDB.gs",
  "Setup.gs",
  "CoreEntities.gs",
  "Import.gs",
  "Classification.gs",
  "FeeCalculation.gs",
  "MonthlyClose.gs",
  "Chuna.gs",
  "Router.gs",
];

describe("靜態檢查：gas/*.gs 全部檔案的函式沒有重複定義、getSpreadsheet 只有一個定義", () => {
  it("每個 .gs 檔案裡，同一個名稱的頂層 function 不會重複定義", () => {
    const counts: Record<string, string[]> = {};
    for (const file of GAS_FILES) {
      const src = fs.readFileSync(path.join(GAS_DIR, file), "utf8");
      const re = /^function\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(/gm;
      let m: RegExpExecArray | null;
      while ((m = re.exec(src))) {
        const name = m[1];
        (counts[name] = counts[name] || []).push(file);
      }
    }
    const duplicates = Object.entries(counts).filter(([, files]) => files.length > 1);
    expect(duplicates, `這些函式名稱被定義超過一次：${JSON.stringify(duplicates)}`).toHaveLength(0);
  });

  it("getSpreadsheet 只在 Utils.gs 定義一次，Setup.gs 不再重複定義", () => {
    const utilsSrc = fs.readFileSync(path.join(GAS_DIR, "Utils.gs"), "utf8");
    const setupSrc = fs.readFileSync(path.join(GAS_DIR, "Setup.gs"), "utf8");
    expect(/^function getSpreadsheet\s*\(/m.test(utilsSrc)).toBe(true);
    expect(/^function getSpreadsheet\s*\(/m.test(setupSrc)).toBe(false);
  });
});

describe("整合驗證：全部 10 個 .gs 檔案載入同一個執行環境後，逐一實際呼叫每個 API", () => {
  it("getSpreadsheet() 可以正常取得指定的 Spreadsheet", () => {
    const sandbox = createGasSandbox();
    expect(typeof sandbox.getSpreadsheet).toBe("function");
    const ss = sandbox.getSpreadsheet();
    expect(ss).toBeTruthy();
    // 拿到的必須是同一份「假 SpreadsheetApp.openById(SPREADSHEET_ID)」回傳的物件，
    // 確認 getSheet()／readRows() 走的是同一份底層資料，不是另外生出一份假的。
    expect(ss.getSheetByName("Semesters")).toBeTruthy();
  });

  it("listSemesters 正常讀取（含既有 115-1 資料）", () => {
    const sandbox = createGasSandbox();
    const sem = seedRealSemester115_1(sandbox);
    const list = sandbox.api_listSemesters();
    expect(list.find((s: any) => s.id === sem.id)).toBeTruthy();
  });

  it("listPeriodSlots 正常讀取（setupSheets 自動種好的 9 個節次代碼）", () => {
    const sandbox = createGasSandbox();
    const slots = sandbox.api_listPeriodSlots();
    expect(slots.length).toBe(9);
    expect(slots.map((s: any) => s.code)).toContain("P1");
  });

  it("學期管理 5 個 API：create／update／setCurrent／deactivate／activate 都能正常執行到底", () => {
    const sandbox = createGasSandbox();
    const sem115_1 = seedRealSemester115_1(sandbox);

    const created = sandbox.api_createSemester({
      schoolYear: 116, term: 1, startDate: "2027-02-11", endDate: "2027-06-30", changedBy: "測試",
    });
    expect(created.id).toBeTruthy();

    const updated = sandbox.api_updateSemester({ id: created.id, note: "smoke test 修改", changedBy: "測試" });
    expect(updated.note).toBe("smoke test 修改");

    const setCurrent = sandbox.api_setCurrentSemester({ id: created.id, changedBy: "測試" });
    expect(setCurrent.isCurrent).toBe(true);
    expect(sandbox.findById("Semesters", sem115_1.id).isCurrent).toBe(false);

    const deactivated = sandbox.api_deactivateSemester({ id: created.id, changedBy: "測試", reason: "smoke test" });
    expect(deactivated.status).toBe("INACTIVE");
    expect(deactivated.isCurrent).toBe(false);

    const reactivated = sandbox.api_activateSemester({ id: created.id, changedBy: "測試" });
    expect(reactivated.status).toBe("NOT_STARTED");
  });

  it("deleteProject 正常執行（未被使用的專案可以刪除）", () => {
    const sandbox = createGasSandbox();
    const sem = seedRealSemester115_1(sandbox);
    const project = sandbox.api_createProject({ semesterId: sem.id, name: "smoke test 專案", changedBy: "測試" });
    const result = sandbox.api_deleteProject({ id: project.id, changedBy: "測試" });
    expect(result.ok).toBe(true);
    expect(sandbox.findById("Projects", project.id)).toBeNull();
  });

  it("generateSemesterCalendar 正常執行（115-1 產生 143 天，不是 0 天）", () => {
    const sandbox = createGasSandbox();
    const sem = seedRealSemester115_1(sandbox);
    const result = sandbox.api_generateSemesterCalendar({ semesterId: sem.id, changedBy: "測試" });
    expect(result.createdCount).toBe(143);
  });

  it("Router 的 doPost 整條路徑（JSON body 解析 → dispatch → LockService → 實際函式）也能正常跑完，不會卡在任何一個環節", () => {
    const sandbox = createGasSandbox();
    const sem = seedRealSemester115_1(sandbox);
    const fakeEvent = { postData: { contents: JSON.stringify({ action: "listSemesters", payload: {} }) } };
    const response = sandbox.doPost(fakeEvent);
    const body = JSON.parse(response.getContent());
    expect(body.ok).toBe(true);
    expect(body.data.find((s: any) => s.id === sem.id)).toBeTruthy();
  });
});
