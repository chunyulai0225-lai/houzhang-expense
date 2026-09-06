// 共用的 GAS 執行環境模擬器，供 tests/gas-*.test.ts 共用。
//
// 做法：把 gas/*.gs 檔案原封不動用 Node 的 vm 模組載入，搭配一個記憶體版「假
// Sheets」（同樣的 2D 陣列儲存格模型，跟 SpreadsheetApp/Utilities/LockService 等
// GAS API 的介面一致），這樣斷言的就是這幾個 .gs 檔案實際會執行出來的行為，
// 不是另外重寫一份邏輯來測。
import fs from "fs";
import path from "path";
import vm from "vm";
import crypto from "crypto";

const GAS_DIR = path.join(__dirname, "..", "..", "gas");
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

class FakeRange {
  sheet: FakeSheet;
  row: number;
  col: number;
  numRows: number;
  numCols: number;
  constructor(sheet: FakeSheet, row: number, col: number, numRows: number, numCols: number) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  getValues() {
    const out: any[][] = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowArr: any[] = [];
      for (let c = 0; c < this.numCols; c++) {
        const actualRow = this.row - 1 + r;
        const actualCol = this.col - 1 + c;
        const rowData = this.sheet.data[actualRow];
        rowArr.push(rowData && rowData[actualCol] !== undefined ? rowData[actualCol] : "");
      }
      out.push(rowArr);
    }
    return out;
  }
  setValues(values: any[][]) {
    for (let r = 0; r < values.length; r++) {
      const actualRow = this.row - 1 + r;
      while (this.sheet.data.length <= actualRow) this.sheet.data.push([]);
      for (let c = 0; c < values[r].length; c++) {
        const actualCol = this.col - 1 + c;
        this.sheet.data[actualRow][actualCol] = values[r][c];
      }
    }
  }
  setNumberFormat() {
    return this;
  }
}

class FakeSheet {
  name: string;
  data: any[][];
  frozenRows: number;
  constructor(name: string) {
    this.name = name;
    this.data = [];
    this.frozenRows = 0;
  }
  getRange(row: number, col: number, numRows?: number, numCols?: number) {
    return new FakeRange(this, row, col, numRows || 1, numCols || 1);
  }
  getLastRow() {
    for (let r = this.data.length - 1; r >= 0; r--) {
      if (this.data[r] && this.data[r].some((v) => v !== undefined && v !== "")) return r + 1;
    }
    return 0;
  }
  getLastColumn() {
    let max = 0;
    this.data.forEach((row) => {
      if (row) max = Math.max(max, row.length);
    });
    return max;
  }
  getMaxRows() {
    return Math.max(this.data.length, 2);
  }
  getFrozenRows() {
    return this.frozenRows;
  }
  setFrozenRows(n: number) {
    this.frozenRows = n;
  }
  appendRow(arr: any[]) {
    this.data.push(arr.slice());
  }
  deleteRow(rowNum: number) {
    this.data.splice(rowNum - 1, 1);
  }
}

class FakeSpreadsheet {
  sheets: Record<string, FakeSheet> = {};
  getSheetByName(name: string) {
    return this.sheets[name] || null;
  }
  insertSheet(name: string) {
    const s = new FakeSheet(name);
    this.sheets[name] = s;
    return s;
  }
  getSheets() {
    return Object.values(this.sheets);
  }
  getId() {
    return "fake-spreadsheet-id";
  }
}

// 測試用的腳本時區固定為 Asia/Taipei，跟這個系統實際的使用情境一致。
const SCRIPT_TIME_ZONE = "Asia/Taipei";

function formatDateInTimeZone(date: Date, timeZone: string, pattern: string): string {
  // 這裡只需要支援 gas/*.gs 實際用到的 "yyyy-MM-dd" 這個 pattern。
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  const map: Record<string, string> = {};
  parts.forEach((p) => (map[p.type] = p.value));
  if (pattern === "yyyy-MM-dd") return `${map.year}-${map.month}-${map.day}`;
  throw new Error("測試用的 formatDate 只支援 yyyy-MM-dd pattern");
}

export function createGasSandbox() {
  const fakeSpreadsheet = new FakeSpreadsheet();
  const sandbox: any = {
    console,
    SpreadsheetApp: {
      openById: () => fakeSpreadsheet,
      create: (name: string) => {
        const ss = new FakeSpreadsheet();
        ss.insertSheet("Sheet1");
        (ss as any).name = name;
        return ss;
      },
      flush: () => {},
    },
    Utilities: {
      getUuid: () => crypto.randomUUID(),
      base64Encode: (bytes: any) => Buffer.from(bytes).toString("base64"),
      formatDate: (date: Date, timeZone: string, pattern: string) => formatDateInTimeZone(date, timeZone, pattern),
    },
    Session: {
      getScriptTimeZone: () => SCRIPT_TIME_ZONE,
    },
    ContentService: {
      createTextOutput: (text: string) => ({ setMimeType: () => ({ getContent: () => text }) }),
      MimeType: { JSON: "json" },
    },
    MimeType: { MICROSOFT_EXCEL: "xlsx" },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    Logger: { log: () => {} },
    DriveApp: { getFileById: () => ({ getAs: () => ({ getBytes: () => Buffer.alloc(0) }), setTrashed: () => {} }) },
  };
  sandbox.global = sandbox;
  vm.createContext(sandbox);
  const src = GAS_FILES.map((f) => fs.readFileSync(path.join(GAS_DIR, f), "utf8")).join("\n;\n");
  vm.runInContext(src, sandbox, { filename: "gas-bundle.js" });
  sandbox.setupSheets();
  return sandbox;
}

// 模擬真實生產環境目前已經有的資料：115學年度第1學期，ACTIVE，isCurrent=true，
// 2026-08-31 ~ 2027-01-20（跟使用者說的真實 Google Sheets 現況一致，不重複建立這筆）。
export function seedRealSemester115_1(sandbox: any) {
  const row = {
    id: sandbox.newId(),
    schoolYear: 115,
    term: 1,
    status: "ACTIVE",
    isCurrent: true,
    startDate: "2026-08-31",
    endDate: "2027-01-20",
    overtimeMatchMode: "TEACHER_WEEKDAY_PERIOD_SUBJECT",
    note: "",
    createdAt: sandbox.nowIso(),
    updatedAt: sandbox.nowIso(),
  };
  sandbox.appendRow("Semesters", row);
  return row;
}

// 模擬使用者直接在 Google Sheets 手動輸入日期、被 Sheets 自動辨識成日期格式的情境
// （目前真實的 115-1 很可能就是這樣建立的）：讀出來是原生 Date 物件，不是字串。
// 用來重現並驗證「產生學校上課日曆顯示 0 天」這個 bug 的根本原因與修法。
export function seedSemesterWithNativeDateCells(sandbox: any, overrides: Partial<{ schoolYear: number; term: number; startDate: string; endDate: string }> = {}) {
  const schoolYear = overrides.schoolYear ?? 115;
  const term = overrides.term ?? 1;
  const startDateStr = overrides.startDate ?? "2026-08-31";
  const endDateStr = overrides.endDate ?? "2027-01-20";
  const row = {
    id: sandbox.newId(),
    schoolYear,
    term,
    status: "ACTIVE",
    isCurrent: true,
    // 刻意存成原生 Date 物件（用 UTC 正午避免時區換日造成的誤差），模擬 Sheets 儲存格
    // 本身就是日期型別、setNumberFormat("@") 也無法回頭轉換既有內容的真實狀況。
    startDate: new Date(startDateStr + "T12:00:00Z"),
    endDate: new Date(endDateStr + "T12:00:00Z"),
    overtimeMatchMode: "TEACHER_WEEKDAY_PERIOD_SUBJECT",
    note: "",
    createdAt: sandbox.nowIso(),
    updatedAt: sandbox.nowIso(),
  };
  sandbox.appendRow("Semesters", row);
  return row;
}
