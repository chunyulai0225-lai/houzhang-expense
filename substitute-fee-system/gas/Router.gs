/**
 * Router.gs — doGet(e)/doPost(e) 進入點與 action 分派表，取代 Express 的路由層。
 *
 * 前端一律用 POST，body 是純文字 JSON：{ action: "...", payload: {...} }，
 * 刻意不設自訂 Content-Type（維持瀏覽器預設 text/plain;charset=UTF-8），這樣才會
 * 被瀏覽器視為 CORS「simple request」，不會觸發 preflight（OPTIONS）——Apps Script
 * Web App 沒有 doOptions()，沒辦法處理 preflight，這是唯一可行的繞法。
 *
 * doGet 額外支援 ?action=xxx&payload=<JSON 字串> 這種查詢字串形式，只是為了方便
 * 用瀏覽器網址列或 curl 做快速測試，前端正式功能一律走 doPost。
 *
 * ACTIONS 表把每個 action 名稱對應到對應檔案裡的 api_* 函式，一個 action 只對應
 * 一個函式，不重複散落判斷式。READ_ONLY_ACTIONS 是明確不需要互斥鎖的查詢類
 * action（純讀取，不寫入任何分頁）；其餘一律視為會寫入資料的 action，執行前用
 * LockService 取得腳本鎖，避免多人同時操作同一批資料造成競爭寫入。
 */

var ACTIONS = {
  // ---- CoreEntities ----
  listSemesters: api_listSemesters,
  createSemester: api_createSemester,
  updateSemester: api_updateSemester,
  setCurrentSemester: api_setCurrentSemester,
  deactivateSemester: api_deactivateSemester,
  activateSemester: api_activateSemester,
  listPersons: api_listPersons,
  listPeriodSlots: api_listPeriodSlots,
  listProjects: api_listProjects,
  createProject: api_createProject,
  updateProject: api_updateProject,
  setProjectActive: api_setProjectActive,
  listWeeklyRules: api_listWeeklyRules,
  weeklyRuleConflicts: api_weeklyRuleConflicts,
  createWeeklyRule: api_createWeeklyRule,
  updateWeeklyRule: api_updateWeeklyRule,
  deactivateWeeklyRule: api_deactivateWeeklyRule,
  listDateRules: api_listDateRules,
  createDateRule: api_createDateRule,
  cancelDateRule: api_cancelDateRule,
  getFeeRuleHistory: api_getFeeRuleHistory,
  createFeeRule: api_createFeeRule,
  deactivateFeeRule: api_deactivateFeeRule,
  generateSemesterCalendar: api_generateSemesterCalendar,
  listCalendarDays: api_listCalendarDays,
  calendarSummary: api_calendarSummary,
  addCalendarDay: api_addCalendarDay,
  updateCalendarDay: api_updateCalendarDay,

  // ---- Import ----
  importSubstituteRows: api_importSubstituteRows,
  listMonthlyImports: api_listMonthlyImports,
  getMonthlyImportDetail: api_getMonthlyImportDetail,
  listSubstituteRecords: api_listSubstituteRecords,
  listUnmatchedTeacherReferences: api_listUnmatchedTeacherReferences,
  resolveTeacherReference: api_resolveTeacherReference,
  autoApplyUnambiguousTeacherMatches: api_autoApplyUnambiguousTeacherMatches,

  // ---- Classification ----
  classifyMonthlyImport: api_classifyMonthlyImport,
  overrideClassification: api_overrideClassification,
  revertToAutoClassification: api_revertToAutoClassification,
  listClassificationPreview: api_listClassificationPreview,

  // ---- FeeCalculation ----
  calculateMonthlyImportFees: api_calculateMonthlyImportFees,
  calculateSubstituteRecordFee: api_calculateSubstituteRecordFee,
  summarizeTeacherMonthlyFees: api_summarizeTeacherMonthlyFees,

  // ---- MonthlyClose ----
  getMonthlyLockStatus: api_getMonthlyLockStatus,
  lockMonth: api_lockMonth,
  unlockMonth: api_unlockMonth,
  acknowledgeIssue: api_acknowledgeIssue,
  revokeAcknowledgement: api_revokeAcknowledgement,
  listAcknowledgements: api_listAcknowledgements,
  createSelfFunded: api_createSelfFunded,
  updateSelfFunded: api_updateSelfFunded,
  deleteSelfFunded: api_deleteSelfFunded,
  listSelfFunded: api_listSelfFunded,
  listPendingIssues: api_listPendingIssues,
  getMonthlyDashboard: api_getMonthlyDashboard,

  // ---- Chuna / Reconciliation ----
  getChunaSummary: api_getChunaSummary,
  generateChunaExcel: api_generateChunaExcel,
  reconcile: api_reconcile,
};

// 純讀取、不寫入任何分頁的 action，不需要互斥鎖，讓多人同時查詢不會互相卡住。
var READ_ONLY_ACTIONS = {
  listSemesters: true, listPersons: true, listPeriodSlots: true, listProjects: true,
  listWeeklyRules: true, weeklyRuleConflicts: true, listDateRules: true, getFeeRuleHistory: true,
  listCalendarDays: true, calendarSummary: true, listMonthlyImports: true, getMonthlyImportDetail: true,
  listSubstituteRecords: true, listUnmatchedTeacherReferences: true, listClassificationPreview: true,
  summarizeTeacherMonthlyFees: true, getMonthlyLockStatus: true, listAcknowledgements: true,
  listSelfFunded: true, listPendingIssues: true, getMonthlyDashboard: true, getChunaSummary: true,
  generateChunaExcel: true,
};

function dispatch(action, payload) {
  var fn = ACTIONS[action];
  if (!fn) throw new Error("未知的 action：" + action);
  payload = payload || {};

  if (READ_ONLY_ACTIONS[action]) {
    return fn(payload);
  }

  var lock = LockService.getScriptLock();
  var acquired = lock.tryLock(30000);
  if (!acquired) throw new Error("系統忙碌中（其他人正在寫入資料），請稍後再試一次");
  try {
    return fn(payload);
  } finally {
    lock.releaseLock();
  }
}

function handleRequest(action, payload) {
  try {
    var data = dispatch(action, payload);
    return ok(data);
  } catch (err) {
    return fail(err && err.message ? err.message : String(err));
  }
}

function doPost(e) {
  var action, payload;
  try {
    var body = JSON.parse(e.postData.contents);
    action = body.action;
    payload = body.payload;
  } catch (err) {
    return fail("無法解析請求內容：" + err.message);
  }
  return handleRequest(action, payload);
}

// 僅供瀏覽器網址列／curl 快速測試用，前端正式功能一律走 doPost。
function doGet(e) {
  var action = e.parameter.action;
  var payload = {};
  if (e.parameter.payload) {
    try {
      payload = JSON.parse(e.parameter.payload);
    } catch (err) {
      return fail("payload 不是合法的 JSON：" + err.message);
    }
  }
  if (!action) {
    return ok({ message: "后庄國小代課費管理系統 GAS API 運作中。用 ?action=listSemesters 測試查詢。" });
  }
  return handleRequest(action, payload);
}
