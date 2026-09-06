/**
 * Utils.gs — 共用小工具：id 產生、時間格式、回應信封、變更紀錄、金額運算。
 *
 * ⚠️ 請把下面 SPREADSHEET_ID 換成你自己的 Google Sheets ID
 * （就是你訊息裡給的那個 1DNbKCwgdSHKvSNCCqvrqUUchOTITdC3e1EWdvLgRY-k）。
 */
var SPREADSHEET_ID = "1DNbKCwgdSHKvSNCCqvrqUUchOTITdC3e1EWdvLgRY-k";

// 統一的試算表存取入口，跟它唯一依賴的 SPREADSHEET_ID 放在同一個檔案——Utils.gs 是
// 每個其他檔案（SheetsDB.gs／Setup.gs／CoreEntities.gs…）都會用到的最基礎檔案，
// 放這裡最不容易在只重貼部分檔案時被漏掉。（這個函式原本放在 Setup.gs，只有這裡
// 一個定義，沒有重複；如果你手動貼到 Apps Script 專案時漏了某個檔案，導致執行期出現
// 「getSpreadsheet is not defined」，代表 Utils.gs 或某個檔案沒有貼到最新內容，
// 請確認 10 個 .gs 檔案都已經是最新版本。）
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function newId() {
  return Utilities.getUuid();
}

function nowIso() {
  return new Date().toISOString();
}

// 統一的成功／失敗回應信封，前端 gasApi() 依這個格式解析。
function ok(data) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function fail(message) {
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(message) }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 金額一律用字串＋parseFloat 做運算後再轉字串，避免 JS 浮點數誤差（跟 Node 版本用
// Prisma.Decimal 的目的一樣：0.1+0.2 這種誤差絕對不能出現在錢的計算上）。
// GAS 沒有內建高精度 Decimal，這裡用「四捨五入到小數點後 2 位」加總的方式降低誤差；
// 目前所有真實資料的金額都是整數，這個精度已經足夠，不會影響任何一筆真實計算結果。
function decAdd(a, b) {
  var x = Number(a || 0);
  var y = Number(b || 0);
  var sum = x + y;
  return Math.round(sum * 100) / 100;
}

function decToStr(n) {
  if (n === null || n === undefined) return null;
  var num = Number(n);
  return Number.isInteger(num) ? String(num) : String(Math.round(num * 100) / 100);
}

// 寫入一筆變更紀錄，簽名跟 Node 版本的 prisma.changeLog.create 對齊。
function writeChangeLog(tableName, recordId, fieldName, oldValue, newValue, changedBy, reason) {
  appendRow("ChangeLog", {
    id: newId(),
    tableName: tableName,
    recordId: recordId,
    fieldName: fieldName || "",
    oldValue: oldValue === undefined || oldValue === null ? "" : String(oldValue),
    newValue: newValue === undefined || newValue === null ? "" : String(newValue),
    changedAt: nowIso(),
    changedBy: changedBy || "",
    reason: reason || "",
    createdAt: nowIso(),
  });
}

function isBlank(v) {
  return v === undefined || v === null || String(v).trim() === "";
}

function requireField(payload, field, label) {
  if (isBlank(payload[field])) {
    throw new Error("請填寫" + (label || field));
  }
}

// 日期欄位一律用 "YYYY-MM-DD" 純文字存放／比對，不使用 Sheets 的日期序列值，
// 避免跨時區造成「差一天」這種對代課紀錄來說致命的錯誤。
//
// 這裡刻意先檢查是不是原生 Date 物件／數字：使用者直接在 Google Sheets 手動輸入
// 看起來像日期的內容時（例如既有的 115 學年度第1學期資料），即使之後對該欄位
// 執行過 setNumberFormat("@") 改成純文字格式，Sheets 也不會回頭把已經存在的儲存格
// 內容轉成字串——用 getValues() 讀出來仍然可能是 Date 物件（極少數情況是日期序列
// 數字），不是 "YYYY-MM-DD" 字串。如果不先轉換，後面直接對它做字串串接／日期運算
// 會產生 Invalid Date，後續依賴日期範圍的功能（例如產生學校上課日曆）就會整批算成
// 0（迴圈的結束條件永遠是 false，一次都不會執行）。
// 轉換時用 Utilities.formatDate() 搭配 Session.getScriptTimeZone()，而不是
// toISOString()：後者一律轉成 UTC，遇到台灣（UTC+8）這種時區會有「差一天」的風險；
// 用試算表腳本本身的時區格式化，才能還原使用者當初實際輸入的那個日期。
function toDateOnly(isoOrDateLike) {
  if (isoOrDateLike === null || isoOrDateLike === undefined || isoOrDateLike === "") return null;
  if (isoOrDateLike instanceof Date) {
    return Utilities.formatDate(isoOrDateLike, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  if (typeof isoOrDateLike === "number") {
    // Google Sheets 的日期序列值：第 0 天是 1899-12-30。
    var epochMs = Date.UTC(1899, 11, 30);
    var asDate = new Date(epochMs + isoOrDateLike * 86400000);
    return Utilities.formatDate(asDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(isoOrDateLike)) return isoOrDateLike.slice(0, 10);
  var d = new Date(isoOrDateLike);
  return d.toISOString().slice(0, 10);
}

var WEEKDAY_BY_JS_INDEX = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"];

function weekdayOfDateOnly(dateOnly) {
  // dateOnly 是 "YYYY-MM-DD"，用 UTC 解析，跟 Node 版本 Date.UTC(...) 的作法一致，
  // 不受 Apps Script 專案的時區設定影響。
  var parts = dateOnly.split("-").map(Number);
  var d = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  return WEEKDAY_BY_JS_INDEX[d.getUTCDay()];
}
