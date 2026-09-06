/**
 * Utils.gs — 共用小工具：id 產生、時間格式、回應信封、變更紀錄、金額運算。
 *
 * ⚠️ 請把下面 SPREADSHEET_ID 換成你自己的 Google Sheets ID
 * （就是你訊息裡給的那個 1DNbKCwgdSHKvSNCCqvrqUUchOTITdC3e1EWdvLgRY-k）。
 */
var SPREADSHEET_ID = "1DNbKCwgdSHKvSNCCqvrqUUchOTITdC3e1EWdvLgRY-k";

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
function toDateOnly(isoOrDateLike) {
  if (!isoOrDateLike) return null;
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
