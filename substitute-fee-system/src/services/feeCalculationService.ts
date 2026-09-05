// Phase 9 第一階段（Phase9-5 修正版）：節次型公費代課費用計算引擎
//
// 只計算 fundingSource ∈ {GENERAL, OVERTIME, PROJECT} 的 SubstituteRecord。
// CONFLICT／TEACHER_UNMATCHED／尚未分類的資料，fundingSource 都是 UNDETERMINED，
// 一律不計算，unitPrice／amount 維持 null，不猜。
//
// 關鍵事實（已用真實資料核對過，不是假設）：Phase 7 的節次解析器不允許複合節次
// （例如「第1,2節」）進入合法的 SubstituteRecord——每一筆成功建立的紀錄本來就只
// 對應一個節次。
//
// Phase9-4/9-5 已確認的核心規則：「分類」與「實際付款金額」是兩件事。
// - unitPrice：這個 feeType 當時的正式參考費率，一律查 FeeRule（不受 Raw 影響）。
// - amount：這筆代課教師實際應領的代課鐘點費，優先讀 SubstituteRecordRaw.substitutePeriodFeeText
//   （真實資料已證明這欄位就是「實際應付金額」本身，不是費率——例如合理員額超鐘點覆蓋的
//   紀錄，Raw 裡明確是 0，因為那405元算在原教師的超鐘點費，不是代課教師的代課費）。
//   Raw 有值（包含 0）就直接採用，不因為 fundingSource=OVERTIME 就自動改成費率；
//   只有在 Raw 缺漏／空白／無法解析時，才退回用 FeeRule 費率計算，並清楚記錄「退回計算」。
//   如果連 FeeRule 都沒有，不要為了硬湊出數字而亂猜，unitPrice/amount 都維持 null。
//
// 刻意不做的事（Phase 9 第一階段之後才處理，這裡完全不碰）：
// - 代導師費／日薪／半日薪（計算單位可能是天／半天，公式尚未確認）
// - 自費代課（另一條手動輸入資料流程，尚未設計）
// - 專案學期累積上限的警示／控制（Project 目前沒有金額欄位）
// - BD／非BD 的費率差異（目前未確認有差異，先不做）
// - PROJECT 專屬費率（PROJECT_PERIOD）：目前仍沿用 SUBSTITUTE_PERIOD
import { FundingSource, FeeType, Prisma } from "@prisma/client";
import { prisma } from "../prismaClient";
import { getEffectiveFeeRule } from "./feeRuleService";

// fundingSource → feeType 的對應集中管理在這裡，不要散落在各處。
// PROJECT 目前刻意沿用 SUBSTITUTE_PERIOD（跟一般公費同一張費率表），
// 未來業務上如果真的出現專案專屬費率，再新增 PROJECT_PERIOD 這個 FeeType。
const FUNDING_SOURCE_FEE_TYPE: Partial<Record<FundingSource, FeeType>> = {
  GENERAL: FeeType.SUBSTITUTE_PERIOD,
  OVERTIME: FeeType.OVERTIME_PERIOD,
  PROJECT: FeeType.SUBSTITUTE_PERIOD,
};

export interface CalculateFeeResult {
  recordId: string;
  unitPrice: string | null;
  amount: string | null;
  feeRuleId: string | null;
  skippedReason?: string;
}

async function clearAmountIfNeeded(recordId: string, unitPrice: Prisma.Decimal | null, amount: Prisma.Decimal | null) {
  if (unitPrice !== null || amount !== null) {
    await prisma.substituteRecord.update({ where: { id: recordId }, data: { unitPrice: null, amount: null } });
  }
}

// 解析 Raw 的代課鐘點費原文。回傳 null 代表「缺漏／空白／無法解析成非負數字」，
// 這種情況要退回 FeeRule，不是當成 0——0 必須是原文本身就明確寫 0 才算數。
function parseRawFee(feeText: string | null | undefined): Prisma.Decimal | null {
  if (feeText == null) return null;
  const trimmed = feeText.trim();
  if (trimmed === "") return null;
  const num = Number(trimmed);
  if (!Number.isFinite(num) || num < 0) return null;
  return new Prisma.Decimal(num);
}

// 對單一 SubstituteRecord 計算金額。
export async function calculateSubstituteRecordFee(recordId: string, changedBy?: string): Promise<CalculateFeeResult> {
  const record = await prisma.substituteRecord.findUniqueOrThrow({
    where: { id: recordId },
    include: {
      monthlyImport: { select: { semesterId: true } },
      rawRecord: { select: { substitutePeriodFeeText: true } },
    },
  });

  const feeType = FUNDING_SOURCE_FEE_TYPE[record.fundingSource];
  if (!feeType) {
    // fundingSource = UNDETERMINED（涵蓋 CONFLICT／TEACHER_UNMATCHED／尚未分類）：不計算，不猜。
    await clearAmountIfNeeded(recordId, record.unitPrice, record.amount);
    return {
      recordId,
      unitPrice: null,
      amount: null,
      feeRuleId: null,
      skippedReason: `fundingSource=${record.fundingSource} 不在 Phase9 第一階段計算範圍`,
    };
  }

  const rule = await getEffectiveFeeRule({
    semesterId: record.monthlyImport.semesterId,
    feeType,
    onDate: record.date,
  });
  const referenceRate = rule?.amount ?? null;
  const rawFee = parseRawFee(record.rawRecord.substitutePeriodFeeText);

  let amount: Prisma.Decimal | null;
  let reason: string;
  if (rawFee !== null) {
    // Raw 有有效數字（含 0）：這就是實際應付金額，不因 fundingSource 而改成費率。
    amount = rawFee;
    reason = `Phase9 費用計算：採用原始代課鐘點費 ${rawFee.toString()}${rule ? `（參考費率 ${feeType}=${referenceRate!.toString()}，FeeRule ${rule.id}）` : "（此時段無生效中的 FeeRule 可供參考）"}`;
  } else if (rule) {
    // Raw 缺漏／空白／無法解析：退回 FeeRule 費率計算，清楚記錄這是退回計算。
    amount = rule.amount;
    reason = `原始代課鐘點費缺漏，已退回 FeeRule 費率計算（${feeType}=${rule.amount.toString()}，FeeRule ${rule.id}）`;
  } else {
    // Raw 缺漏，且沒有生效中的 FeeRule 可退回：不要為了硬湊數字而亂猜。
    await clearAmountIfNeeded(recordId, record.unitPrice, record.amount);
    return {
      recordId,
      unitPrice: null,
      amount: null,
      feeRuleId: null,
      skippedReason: `原始代課鐘點費缺漏，且找不到 ${feeType} 在 ${record.date.toISOString().slice(0, 10)} 當時生效的 FeeRule`,
    };
  }

  const updated = await prisma.substituteRecord.update({
    where: { id: recordId },
    data: { unitPrice: referenceRate, amount },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "substitute_records",
      recordId,
      fieldName: "amount",
      newValue: amount.toString(),
      changedBy,
      reason,
    },
  });

  return {
    recordId,
    unitPrice: updated.unitPrice?.toString() ?? null,
    amount: updated.amount?.toString() ?? null,
    feeRuleId: rule?.id ?? null,
  };
}

// 對一個匯入批次的所有 SubstituteRecord 計算金額，供 API/UI 一次觸發整批。
export async function calculateMonthlyImportFees(monthlyImportId: string, changedBy?: string): Promise<CalculateFeeResult[]> {
  const records = await prisma.substituteRecord.findMany({ where: { monthlyImportId }, select: { id: true } });
  const results: CalculateFeeResult[] = [];
  for (const r of records) {
    results.push(await calculateSubstituteRecordFee(r.id, changedBy));
  }
  return results;
}

export interface TeacherMonthlyFeeSummaryRow {
  substituteTeacherId: string | null;
  substituteTeacherName: string; // 未配對時用原始文字姓名顯示，並標註未配對，避免直接被合併不見
  generalCount: number;
  generalAmount: string;
  overtimeCount: number;
  overtimeAmount: string;
  projectCount: number;
  projectAmount: string;
  totalCount: number;
  totalAmount: string;
}

const ZERO = new Prisma.Decimal(0);

// 「代課教師 × 月份 × fundingSource」彙總。只加總已經算好金額（amount 不為 null）的紀錄，
// 不在這裡重新計算金額本身——彙總前請先呼叫 calculateMonthlyImportFees。
export async function summarizeTeacherMonthlyFees(monthlyImportIds: string[]): Promise<TeacherMonthlyFeeSummaryRow[]> {
  const records = await prisma.substituteRecord.findMany({
    where: { monthlyImportId: { in: monthlyImportIds }, amount: { not: null } },
    include: { substituteTeacher: true, rawRecord: { select: { substituteTeacherText: true } } },
  });

  const byKey = new Map<string, TeacherMonthlyFeeSummaryRow>();

  for (const r of records) {
    const key = r.substituteTeacherId ?? `未配對:${r.rawRecord?.substituteTeacherText ?? r.id}`;
    const displayName = r.substituteTeacher
      ? r.substituteTeacher.name
      : `${r.rawRecord?.substituteTeacherText ?? "未知"}（未配對）`;

    let row = byKey.get(key);
    if (!row) {
      row = {
        substituteTeacherId: r.substituteTeacherId,
        substituteTeacherName: displayName,
        generalCount: 0,
        generalAmount: "0",
        overtimeCount: 0,
        overtimeAmount: "0",
        projectCount: 0,
        projectAmount: "0",
        totalCount: 0,
        totalAmount: "0",
      };
      byKey.set(key, row);
    }

    const amount = r.amount ?? ZERO;
    if (r.fundingSource === "GENERAL") {
      row.generalCount += 1;
      row.generalAmount = new Prisma.Decimal(row.generalAmount).plus(amount).toString();
    } else if (r.fundingSource === "OVERTIME") {
      row.overtimeCount += 1;
      row.overtimeAmount = new Prisma.Decimal(row.overtimeAmount).plus(amount).toString();
    } else if (r.fundingSource === "PROJECT") {
      row.projectCount += 1;
      row.projectAmount = new Prisma.Decimal(row.projectAmount).plus(amount).toString();
    }
    row.totalCount += 1;
    row.totalAmount = new Prisma.Decimal(row.totalAmount).plus(amount).toString();
  }

  return Array.from(byKey.values()).sort((a, b) => a.substituteTeacherName.localeCompare(b.substituteTeacherName, "zh-Hant"));
}
