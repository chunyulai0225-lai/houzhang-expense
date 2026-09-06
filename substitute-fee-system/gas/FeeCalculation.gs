/**
 * FeeCalculation.gs — 逐行對照 feeCalculationService.ts 的 Phase9-5 費用計算引擎。
 *
 * 核心規則（不可更動）：
 * - 只計算 fundingSource ∈ {GENERAL, OVERTIME, PROJECT}；UNDETERMINED（涵蓋 CONFLICT／
 *   TEACHER_UNMATCHED／尚未分類）一律不計算，unitPrice/amount 維持 null，不猜。
 * - unitPrice：一律查當時生效的 FeeRule（參考費率，不受 Raw 影響）。
 * - amount：優先讀 RawRecord.substitutePeriodFeeText——Raw 有合法非負數字（包含 0）
 *   就直接採用，不因為 fundingSource=OVERTIME 就自動改成費率；只有 Raw 缺漏/空白/
 *   無法解析時才退回用 FeeRule 費率。連 FeeRule 都沒有就不硬湊數字，維持 null。
 * - PROJECT 目前刻意沿用 SUBSTITUTE_PERIOD 費率表，不建立 PROJECT_PERIOD。
 */

var FUNDING_SOURCE_FEE_TYPE = {
  GENERAL: "SUBSTITUTE_PERIOD",
  OVERTIME: "OVERTIME_PERIOD",
  PROJECT: "SUBSTITUTE_PERIOD",
};

// 解析 Raw 的代課鐘點費原文。回傳 null 代表「缺漏／空白／無法解析成非負數字」，
// 這種情況要退回 FeeRule，不是當成 0——0 必須是原文本身就明確寫 0 才算數。
function parseRawFee(feeText) {
  if (feeText === null || feeText === undefined) return null;
  var trimmed = String(feeText).trim();
  if (trimmed === "") return null;
  var num = Number(trimmed);
  if (!isFinite(num) || num < 0) return null;
  return num;
}

function clearAmountIfNeeded(recordId, record) {
  if (record.unitPrice !== "" || record.amount !== "") {
    updateRow("SubstituteRecords", recordId, { unitPrice: "", amount: "", updatedAt: nowIso() });
  }
}

// 對單一 SubstituteRecord 計算金額。
function calculateSubstituteRecordFee(recordId, changedBy) {
  var record = findById("SubstituteRecords", recordId);
  if (!record) throw new Error("找不到代課紀錄");
  var monthlyImport = findById("MonthlyImports", record.monthlyImportId);
  assertMonthNotLocked(Number(monthlyImport.year), Number(monthlyImport.month));

  var feeType = FUNDING_SOURCE_FEE_TYPE[record.fundingSource];
  if (!feeType) {
    // fundingSource = UNDETERMINED：不計算，不猜。
    clearAmountIfNeeded(recordId, record);
    return {
      recordId: recordId, unitPrice: null, amount: null, feeRuleId: null,
      skippedReason: "fundingSource=" + record.fundingSource + " 不在 Phase9 第一階段計算範圍",
    };
  }

  var rule = getEffectiveFeeRule(monthlyImport.semesterId, feeType, record.date);
  var referenceRate = rule ? Number(rule.amount) : null;
  var rawRecord = record.rawRecordId ? findById("RawRecords", record.rawRecordId) : null;
  var rawFee = parseRawFee(rawRecord ? rawRecord.substitutePeriodFeeText : null);

  var amount, reason;
  if (rawFee !== null) {
    // Raw 有有效數字（含 0）：這就是實際應付金額，不因 fundingSource 而改成費率。
    amount = rawFee;
    reason = "Phase9 費用計算：採用原始代課鐘點費 " + decToStr(rawFee) +
      (rule ? "（參考費率 " + feeType + "=" + decToStr(referenceRate) + "，FeeRule " + rule.id + "）" : "（此時段無生效中的 FeeRule 可供參考）");
  } else if (rule) {
    // Raw 缺漏／空白／無法解析：退回 FeeRule 費率計算，清楚記錄這是退回計算。
    amount = referenceRate;
    reason = "原始代課鐘點費缺漏，已退回 FeeRule 費率計算（" + feeType + "=" + decToStr(referenceRate) + "，FeeRule " + rule.id + "）";
  } else {
    // Raw 缺漏，且沒有生效中的 FeeRule 可退回：不要為了硬湊數字而亂猜。
    clearAmountIfNeeded(recordId, record);
    return {
      recordId: recordId, unitPrice: null, amount: null, feeRuleId: null,
      skippedReason: "原始代課鐘點費缺漏，且找不到 " + feeType + " 在 " + record.date + " 當時生效的 FeeRule",
    };
  }

  var updated = updateRow("SubstituteRecords", recordId, {
    unitPrice: referenceRate === null ? "" : decToStr(referenceRate),
    amount: decToStr(amount),
    updatedAt: nowIso(),
  });

  writeChangeLog("substitute_records", recordId, "amount", null, decToStr(amount), changedBy, reason);

  return {
    recordId: recordId,
    unitPrice: updated.unitPrice === "" ? null : updated.unitPrice,
    amount: updated.amount === "" ? null : updated.amount,
    feeRuleId: rule ? rule.id : null,
  };
}

// 對一個匯入批次的所有 SubstituteRecord 計算金額，供 API/UI 一次觸發整批。
function api_calculateMonthlyImportFees(payload) {
  requireField(payload, "id", "id");
  var records = readRows("SubstituteRecords").filter(function (r) { return r.monthlyImportId === payload.id; });
  var results = [];
  records.forEach(function (r) {
    results.push(calculateSubstituteRecordFee(r.id, payload.changedBy));
  });
  return results;
}

function api_calculateSubstituteRecordFee(payload) {
  requireField(payload, "id", "id");
  return calculateSubstituteRecordFee(payload.id, payload.changedBy);
}

// 「代課教師 × 月份 × fundingSource」彙總。只加總已經算好金額（amount 不為空）的紀錄，
// 不在這裡重新計算金額本身——彙總前請先呼叫 api_calculateMonthlyImportFees。
function summarizeTeacherMonthlyFees(monthlyImportIds) {
  var idSet = {};
  monthlyImportIds.forEach(function (id) { idSet[id] = true; });
  var records = readRows("SubstituteRecords").filter(function (r) {
    return idSet[r.monthlyImportId] && r.amount !== "" && r.amount !== null && r.amount !== undefined;
  });

  var byKey = {};
  var order = [];

  records.forEach(function (r) {
    var rawRecord = r.rawRecordId ? findById("RawRecords", r.rawRecordId) : null;
    var key = r.substituteTeacherId || ("未配對:" + (rawRecord ? rawRecord.substituteTeacherText : r.id));
    var person = getPersonRef(r.substituteTeacherId);
    var displayName = person ? person.name : ((rawRecord && rawRecord.substituteTeacherText ? rawRecord.substituteTeacherText : "未知") + "（未配對）");

    var row = byKey[key];
    if (!row) {
      row = {
        substituteTeacherId: r.substituteTeacherId || null, substituteTeacherName: displayName,
        generalCount: 0, generalAmount: "0", overtimeCount: 0, overtimeAmount: "0",
        projectCount: 0, projectAmount: "0", totalCount: 0, totalAmount: "0",
      };
      byKey[key] = row;
      order.push(key);
    }

    var amount = Number(r.amount) || 0;
    if (r.fundingSource === "GENERAL") {
      row.generalCount += 1;
      row.generalAmount = decAdd(row.generalAmount, amount);
    } else if (r.fundingSource === "OVERTIME") {
      row.overtimeCount += 1;
      row.overtimeAmount = decAdd(row.overtimeAmount, amount);
    } else if (r.fundingSource === "PROJECT") {
      row.projectCount += 1;
      row.projectAmount = decAdd(row.projectAmount, amount);
    }
    row.totalCount += 1;
    row.totalAmount = decAdd(row.totalAmount, amount);
  });

  var result = order.map(function (key) { return byKey[key]; });
  result.sort(function (a, b) { return String(a.substituteTeacherName).localeCompare(String(b.substituteTeacherName), "zh-Hant"); });
  return result;
}

function api_summarizeTeacherMonthlyFees(payload) {
  requireField(payload, "monthlyImportIds", "monthlyImportIds");
  return summarizeTeacherMonthlyFees(payload.monthlyImportIds);
}
