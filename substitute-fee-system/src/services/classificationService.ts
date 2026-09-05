// Phase 8（第一階段）：公費代課自動分類引擎
//
// 只做「這一節課屬於哪種經費來源？」，完全不做「應該付多少錢？」——
// 不使用 405/133/1399/700 等任何 Excel 原始金額欄位，不寫任何費率、
// 不判斷代導師費/日薪/半日薪何時產生，不做 BD/非BD 付款規則，不算節數。
// 這些全部留給後續階段。
//
// 分類優先順序（已確認，不自行擴充）：
//   1. SpecialDateRule（單日例外）完全比對 學期+原教師+日期+節次，優先於週規則
//   2. SpecialWeeklyRule（每週固定規則）比對 學期+原教師+星期+節次，
//      且 effectiveDate <= 代課日期 <= endDate（或 endDate 為空）；
//      是否比對科目由 Semester.overtimeMatchMode 決定，不自行新增規則
//   3. 以上都沒有 → 一般公費（GENERAL）
// 任何一層出現多個候選規則都不自動選一個，標記 CONFLICT 並保留所有候選。
// 原教師尚未配對到 Person 時，根本查不到規則，標記 TEACHER_UNMATCHED。

import type {
  FundingSource,
  OverrideClassification,
  Semester,
  SpecialDateRule,
  SpecialWeeklyRule,
  SubstituteRecord,
} from "@prisma/client";
import { ClassificationMethod } from "@prisma/client";
import { prisma } from "../prismaClient";

interface ConflictCandidate {
  ruleId: string;
  wouldBeFundingSource: FundingSource;
  projectId: string | null;
}

interface ConflictInfo {
  source: "DATE_EXCEPTION" | "WEEKLY_RULE";
  candidates: ConflictCandidate[];
}

export interface ClassificationOutcome {
  fundingSource: FundingSource;
  classificationMethod: ClassificationMethod;
  ruleId: string | null;
  projectId: string | null;
  conflict: ConflictInfo | null;
}

function toFundingSource(value: OverrideClassification): FundingSource {
  // OverrideClassification { GENERAL, OVERTIME, PROJECT } 是 FundingSource 的子集，
  // 直接對應，不需要轉換表。
  return value as unknown as FundingSource;
}

// 核心分類邏輯：純函式（不寫入資料庫），方便單獨測試每一種情境。
export async function computeClassification(
  record: Pick<SubstituteRecord, "originalTeacherId" | "date" | "weekday" | "periodCode" | "subject">,
  semester: Pick<Semester, "id" | "overtimeMatchMode">
): Promise<ClassificationOutcome> {
  if (!record.originalTeacherId) {
    return {
      fundingSource: "UNDETERMINED",
      classificationMethod: ClassificationMethod.TEACHER_UNMATCHED,
      ruleId: null,
      projectId: null,
      conflict: null,
    };
  }

  if (!record.periodCode) {
    // Phase 7 已保證能建立 SubstituteRecord 的資料列一定有 periodCode，這裡僅作防禦。
    return {
      fundingSource: "UNDETERMINED",
      classificationMethod: ClassificationMethod.TEACHER_UNMATCHED,
      ruleId: null,
      projectId: null,
      conflict: null,
    };
  }

  // 第一優先：單日例外
  const dateRules: SpecialDateRule[] = await prisma.specialDateRule.findMany({
    where: {
      semesterId: semester.id,
      personId: record.originalTeacherId,
      date: record.date,
      periodCode: record.periodCode,
      isCancelled: false,
    },
  });

  if (dateRules.length === 1) {
    const rule = dateRules[0];
    const fundingSource = toFundingSource(rule.overrideClassification);
    return {
      fundingSource,
      classificationMethod: ClassificationMethod.DATE_EXCEPTION,
      ruleId: rule.id,
      projectId: fundingSource === "PROJECT" ? rule.projectId : null,
      conflict: null,
    };
  }
  if (dateRules.length > 1) {
    return {
      fundingSource: "UNDETERMINED",
      classificationMethod: ClassificationMethod.CONFLICT,
      ruleId: null,
      projectId: null,
      conflict: {
        source: "DATE_EXCEPTION",
        candidates: dateRules.map((r) => ({
          ruleId: r.id,
          wouldBeFundingSource: toFundingSource(r.overrideClassification),
          projectId: r.overrideClassification === "PROJECT" ? r.projectId : null,
        })),
      },
    };
  }

  // 第二優先：每週固定規則（有效期間必須涵蓋代課日期）
  const weeklyRulesRaw: SpecialWeeklyRule[] = await prisma.specialWeeklyRule.findMany({
    where: {
      semesterId: semester.id,
      personId: record.originalTeacherId,
      weekday: record.weekday,
      periodCode: record.periodCode,
      effectiveDate: { lte: record.date },
      OR: [{ endDate: null }, { endDate: { gte: record.date } }],
    },
  });

  const subjectModeRequired = semester.overtimeMatchMode === "TEACHER_WEEKDAY_PERIOD_SUBJECT";
  const validCandidates = weeklyRulesRaw.filter((rule) => {
    // PROJECT 規則沒有 projectId 是資料異常，視為無效規則（不形成有效 PROJECT 判斷）
    if (rule.ruleType === "PROJECT" && !rule.projectId) return false;
    if (subjectModeRequired) {
      // 科目比對模式下，規則本身沒有設定科目就無法驗證是否相符，視為不適用；
      // 不把「沒設定科目」當成萬用比對，避免自行擴充判斷規則。
      return rule.subject != null && record.subject != null && rule.subject === record.subject;
    }
    return true;
  });

  if (validCandidates.length === 1) {
    const rule = validCandidates[0];
    const fundingSource = rule.ruleType as unknown as FundingSource; // RuleType { OVERTIME, PROJECT } 對應 FundingSource 子集
    return {
      fundingSource,
      classificationMethod: ClassificationMethod.WEEKLY_RULE,
      ruleId: rule.id,
      projectId: rule.ruleType === "PROJECT" ? rule.projectId : null,
      conflict: null,
    };
  }
  if (validCandidates.length > 1) {
    return {
      fundingSource: "UNDETERMINED",
      classificationMethod: ClassificationMethod.CONFLICT,
      ruleId: null,
      projectId: null,
      conflict: {
        source: "WEEKLY_RULE",
        candidates: validCandidates.map((r) => ({
          ruleId: r.id,
          wouldBeFundingSource: r.ruleType as unknown as FundingSource,
          projectId: r.ruleType === "PROJECT" ? r.projectId : null,
        })),
      },
    };
  }

  // 第三優先：一般公費
  return {
    fundingSource: "GENERAL",
    classificationMethod: ClassificationMethod.GENERAL_DEFAULT,
    ruleId: null,
    projectId: null,
    conflict: null,
  };
}

// 對單一 SubstituteRecord 執行分類並寫入資料庫。
// auto* 欄位永遠更新；只有在該筆「尚未被人工覆寫」時，才會同時更新目前生效的
// fundingSource/classificationMethod/classificationRuleId/projectId。
export async function classifySubstituteRecord(
  recordId: string,
  changedBy?: string
): Promise<SubstituteRecord> {
  const record = await prisma.substituteRecord.findUniqueOrThrow({
    where: { id: recordId },
    include: { monthlyImport: true },
  });
  const semester = await prisma.semester.findUniqueOrThrow({ where: { id: record.monthlyImport.semesterId } });

  const outcome = await computeClassification(record, semester);

  const data: Record<string, unknown> = {
    classifiedAt: new Date(),
    autoFundingSource: outcome.fundingSource,
    autoClassificationMethod: outcome.classificationMethod,
    autoClassificationRuleId: outcome.ruleId,
    autoProjectId: outcome.projectId,
    conflictCandidatesJson: outcome.conflict ? JSON.stringify(outcome.conflict) : null,
  };

  if (!record.isManuallyModified) {
    data.fundingSource = outcome.fundingSource;
    data.classificationMethod = outcome.classificationMethod;
    data.classificationRuleId = outcome.ruleId;
    data.projectId = outcome.projectId;
  }

  const updated = await prisma.substituteRecord.update({ where: { id: recordId }, data });

  if (!record.isManuallyModified && record.fundingSource !== outcome.fundingSource) {
    await prisma.changeLog.create({
      data: {
        tableName: "substitute_records",
        recordId,
        fieldName: "fundingSource",
        oldValue: record.fundingSource,
        newValue: outcome.fundingSource,
        changedBy,
        reason: `自動分類（${outcome.classificationMethod}）`,
      },
    });
  }

  return updated;
}

export interface ClassifyBatchSummary {
  total: number;
  general: number;
  overtime: number;
  project: number;
  conflict: number;
  teacherUnmatched: number;
  manualPreserved: number;
}

// 對一個匯入批次的所有 SubstituteRecord 執行分類，可重複執行（例如規則調整後
// 「重新自動分類」）；已人工覆寫的資料列不會被洗掉，只會更新 auto* 參考欄位。
export async function classifyMonthlyImport(monthlyImportId: string, changedBy?: string): Promise<ClassifyBatchSummary> {
  const records = await prisma.substituteRecord.findMany({
    where: { monthlyImportId },
    select: { id: true },
  });

  const summary: ClassifyBatchSummary = {
    total: records.length,
    general: 0,
    overtime: 0,
    project: 0,
    conflict: 0,
    teacherUnmatched: 0,
    manualPreserved: 0,
  };

  for (const { id } of records) {
    const updated = await classifySubstituteRecord(id, changedBy);
    if (updated.isManuallyModified) {
      summary.manualPreserved += 1;
      continue;
    }
    switch (updated.fundingSource) {
      case "GENERAL":
        summary.general += 1;
        break;
      case "OVERTIME":
        summary.overtime += 1;
        break;
      case "PROJECT":
        summary.project += 1;
        break;
      case "UNDETERMINED":
        if (updated.classificationMethod === ClassificationMethod.CONFLICT) summary.conflict += 1;
        else if (updated.classificationMethod === ClassificationMethod.TEACHER_UNMATCHED) summary.teacherUnmatched += 1;
        break;
    }
  }

  return summary;
}

export interface OverrideClassificationInput {
  fundingSource: FundingSource;
  projectId?: string;
}

// 人工覆寫：管理者的最終決定。保留原始自動分類結果（auto* 欄位不變動），
// 並記錄覆寫原因與 ChangeLog，讓下一次自動分類不會把這筆洗掉。
export async function overrideClassification(
  recordId: string,
  input: OverrideClassificationInput,
  changedBy: string | undefined,
  reason: string
): Promise<SubstituteRecord> {
  if (input.fundingSource === "PROJECT" && !input.projectId) {
    throw new Error("覆寫為 PROJECT 時必須指定 projectId");
  }
  if (!reason || !reason.trim()) {
    throw new Error("人工覆寫必須填寫原因");
  }

  const existing = await prisma.substituteRecord.findUniqueOrThrow({ where: { id: recordId } });

  const updated = await prisma.substituteRecord.update({
    where: { id: recordId },
    data: {
      fundingSource: input.fundingSource,
      projectId: input.fundingSource === "PROJECT" ? input.projectId : null,
      classificationMethod: ClassificationMethod.MANUAL_OVERRIDE,
      classificationRuleId: null,
      isManuallyModified: true,
      manualOverrideReason: reason,
      manualOverrideAt: new Date(),
      manualOverrideBy: changedBy,
    },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "substitute_records",
      recordId,
      fieldName: "fundingSource",
      oldValue: existing.fundingSource,
      newValue: input.fundingSource,
      changedBy,
      reason,
    },
  });

  return updated;
}

// 復原成目前的自動分類結果（撤銷人工覆寫）。auto* 欄位是上一次自動分類算出來的結果，
// 復原後建議重新執行一次分類以確保是最新的。
export async function revertToAutoClassification(
  recordId: string,
  changedBy: string | undefined,
  reason: string
): Promise<SubstituteRecord> {
  const record = await prisma.substituteRecord.findUniqueOrThrow({ where: { id: recordId } });
  if (!record.isManuallyModified) {
    throw new Error("此筆尚未被人工覆寫，不需要復原");
  }

  const updated = await prisma.substituteRecord.update({
    where: { id: recordId },
    data: {
      fundingSource: record.autoFundingSource ?? "UNDETERMINED",
      classificationMethod: record.autoClassificationMethod ?? ClassificationMethod.GENERAL_DEFAULT,
      classificationRuleId: record.autoClassificationRuleId,
      projectId: record.autoProjectId,
      isManuallyModified: false,
      manualOverrideReason: null,
      manualOverrideAt: null,
      manualOverrideBy: null,
    },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "substitute_records",
      recordId,
      fieldName: "isManuallyModified",
      oldValue: "true",
      newValue: "false",
      changedBy,
      reason,
    },
  });

  return updated;
}

export interface ClassificationPreviewFilter {
  fundingSource?: FundingSource;
  classificationMethod?: ClassificationMethod;
  isManuallyModified?: boolean;
  staffType?: "BD" | "NON_BD" | "UNKNOWN";
}

export async function listClassificationPreview(monthlyImportId: string, filter: ClassificationPreviewFilter = {}) {
  return prisma.substituteRecord.findMany({
    where: {
      monthlyImportId,
      ...(filter.fundingSource ? { fundingSource: filter.fundingSource } : {}),
      ...(filter.classificationMethod ? { classificationMethod: filter.classificationMethod } : {}),
      ...(filter.isManuallyModified !== undefined ? { isManuallyModified: filter.isManuallyModified } : {}),
      ...(filter.staffType ? { staffType: filter.staffType } : {}),
    },
    include: { originalTeacher: true, substituteTeacher: true, project: true, rawRecord: true },
    orderBy: [{ date: "asc" }],
  });
}
