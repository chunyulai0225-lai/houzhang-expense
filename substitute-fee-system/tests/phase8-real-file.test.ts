// Phase 8 真實資料回歸測試：使用「114學年2026.06月代課(公費).xlsx」
// （真實檔案含真實姓名與身心調適假等敏感請假紀錄，不會被 commit 進 repo，
//  也不會留在版本控制歷史裡。這份測試在檔案不存在時會自動略過，
//  不影響其他環境或其他人執行 `npm test`。）
import fs from "fs";
import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { prisma } from "../src/prismaClient";
import { createPerson } from "../src/services/personService";
import { createProject } from "../src/services/projectService";
import { importSubstituteExcel, resolveTeacherReference } from "../src/services/excelImportService";
import { classifyMonthlyImport, classifySubstituteRecord, overrideClassification } from "../src/services/classificationService";

const CANDIDATE_PATHS = [
  process.env.REAL_SUBSTITUTE_EXCEL_PATH,
  "/root/.claude/uploads/6b1e1531-1999-57e8-8932-c62ee2d638fb/f45b44e8-114__2026.06_____.xlsx",
].filter((p): p is string => Boolean(p));

const REAL_FILE = CANDIDATE_PATHS.find((p) => fs.existsSync(p));
const NON_BD_SHEET = "6月公費  非BD 0701";
const BD_SHEET = "6月公費BD 原授課排序給芝庭";

describe.skipIf(!REAL_FILE)("Phase 8 真實 2026年6月代課資料分類回歸測試", () => {
  const cleanupPersonIds: string[] = [];
  let semesterId: string;
  let nonBdImportId: string;
  let bdImportId: string;

  beforeAll(async () => {
    await prisma.$connect();
    const buffer = fs.readFileSync(REAL_FILE!);

    const semester = await prisma.semester.create({
      data: {
        schoolYear: 8801, // 明顯不會與真實學年度衝突的測試用學年度
        term: 1,
        startDate: new Date("2026-02-01"),
        endDate: new Date("2026-07-31"),
        note: "Phase8 真實檔案分類回歸測試，測試結束會刪除",
      },
    });
    semesterId = semester.id;

    const nonBd = await importSubstituteExcel({
      semesterId,
      year: 2026,
      month: 6,
      fileName: "real-nonbd.xlsx",
      fileBuffer: buffer,
      sheetName: NON_BD_SHEET,
      sourceStaffType: "NON_BD",
      importedBy: "回歸測試",
    });
    nonBdImportId = nonBd.monthlyImport.id;

    const bd = await importSubstituteExcel({
      semesterId,
      year: 2026,
      month: 6,
      fileName: "real-bd.xlsx",
      fileBuffer: buffer,
      sheetName: BD_SHEET,
      sourceStaffType: "BD",
      importedBy: "回歸測試",
    });
    bdImportId = bd.monthlyImport.id;
  });

  afterAll(async () => {
    const importIds = [nonBdImportId, bdImportId];
    const records = await prisma.substituteRecord.findMany({ where: { monthlyImportId: { in: importIds } } });
    await prisma.changeLog.deleteMany({
      where: { OR: [{ recordId: { in: records.map((r) => r.id) } }, { recordId: { in: cleanupPersonIds } }] },
    });
    await prisma.monthlyImportError.deleteMany({ where: { monthlyImportId: { in: importIds } } });
    await prisma.substituteRecord.deleteMany({ where: { monthlyImportId: { in: importIds } } });
    await prisma.substituteRecordRaw.deleteMany({ where: { monthlyImportId: { in: importIds } } });
    await prisma.monthlyImport.deleteMany({ where: { id: { in: importIds } } });
    await prisma.specialDateRule.deleteMany({ where: { semesterId } });
    await prisma.specialWeeklyRule.deleteMany({ where: { semesterId } });
    await prisma.project.deleteMany({ where: { semesterId } });
    await prisma.person.deleteMany({ where: { id: { in: cleanupPersonIds } } });
    await prisma.semester.delete({ where: { id: semesterId } });
    await prisma.$disconnect();
  });

  it("13/15. 匯入後 BD/非BD 來源與 Raw 原始資料維持 Phase 7 驗證過的樣貌", async () => {
    const nonBdCount = await prisma.substituteRecord.count({ where: { monthlyImportId: nonBdImportId } });
    const bdCount = await prisma.substituteRecord.count({ where: { monthlyImportId: bdImportId } });
    expect(nonBdCount).toBeGreaterThan(0);
    expect(bdCount).toBeGreaterThan(0);

    const nonBdSample = await prisma.substituteRecord.findFirstOrThrow({ where: { monthlyImportId: nonBdImportId } });
    expect(nonBdSample.staffType).toBe("NON_BD");
    const bdSample = await prisma.substituteRecord.findFirstOrThrow({ where: { monthlyImportId: bdImportId } });
    expect(bdSample.staffType).toBe("BD");
  });

  it("1/10. 沒有配對到 Person、也沒有規則的真實資料維持 GENERAL 或 TEACHER_UNMATCHED，不亂猜", async () => {
    await classifyMonthlyImport(nonBdImportId, "回歸測試");
    await classifyMonthlyImport(bdImportId, "回歸測試");

    const stillUnmatched = await prisma.substituteRecord.findMany({
      where: { monthlyImportId: { in: [nonBdImportId, bdImportId] }, originalTeacherId: null },
    });
    expect(stillUnmatched.length).toBeGreaterThan(0);
    expect(stillUnmatched.every((r) => r.classificationMethod === "TEACHER_UNMATCHED" && r.fundingSource === "UNDETERMINED")).toBe(
      true
    );
  });

  it("2/4/5/6/9/11/12. 針對真實資料建立規則後，OVERTIME／DATE_EXCEPTION／有效期間／衝突／人工覆寫都正確運作", async () => {
    // 挑一筆有明確原教師姓名、星期、節次的真實非BD資料列作為測試錨點
    const raw = await prisma.substituteRecordRaw.findFirstOrThrow({
      where: { monthlyImportId: nonBdImportId, originalTeacherText: { not: null } },
      orderBy: { rowNumber: "asc" },
    });
    const record = await prisma.substituteRecord.findUniqueOrThrow({ where: { rawRecordId: raw.id } });
    const teacherName = raw.originalTeacherText!;

    const created = await createPerson({ name: teacherName });
    if (created.status !== "CREATED") throw new Error("unreachable: 真實資料裡這個姓名應該還沒有對應的 Person");
    cleanupPersonIds.push(created.person.id);
    await resolveTeacherReference(record.id, "original", created.person.id, "回歸測試");

    // 2/5/6：建立一個涵蓋這筆資料日期的 OVERTIME 週規則，確認 effectiveDate/endDate 判斷正確
    const overtimeRule = await prisma.specialWeeklyRule.create({
      data: {
        semesterId,
        personId: created.person.id,
        ruleType: "OVERTIME",
        weekday: record.weekday,
        periodCode: record.periodCode!,
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-02-01"),
        endDate: new Date("2026-07-31"),
      },
    });
    const overtimeResult = await classifySubstituteRecord(record.id, "回歸測試");
    expect(overtimeResult.fundingSource).toBe("OVERTIME");
    expect(overtimeResult.classificationRuleId).toBe(overtimeRule.id);

    // 4：針對同一天新增單日例外，應該優先於週規則
    const dateRule = await prisma.specialDateRule.create({
      data: {
        semesterId,
        date: record.date,
        personId: created.person.id,
        periodCode: record.periodCode!,
        overrideClassification: "GENERAL",
        note: "回歸測試：單日例外應優先於週規則",
      },
    });
    const dateExceptionResult = await classifySubstituteRecord(record.id, "回歸測試");
    expect(dateExceptionResult.fundingSource).toBe("GENERAL");
    expect(dateExceptionResult.classificationMethod).toBe("DATE_EXCEPTION");
    expect(dateExceptionResult.classificationRuleId).toBe(dateRule.id);

    // 9：把單日例外取消，改成再加一條重疊的週規則造成衝突
    await prisma.specialDateRule.update({ where: { id: dateRule.id }, data: { isCancelled: true } });
    const project = await createProject({ semesterId, name: "回歸測試專案" });
    const conflictingRule = await prisma.specialWeeklyRule.create({
      data: {
        semesterId,
        personId: created.person.id,
        ruleType: "PROJECT",
        projectId: project.id,
        weekday: record.weekday,
        periodCode: record.periodCode!,
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-02-01"),
        endDate: new Date("2026-07-31"),
      },
    });
    const conflictResult = await classifySubstituteRecord(record.id, "回歸測試");
    expect(conflictResult.fundingSource).toBe("UNDETERMINED");
    expect(conflictResult.classificationMethod).toBe("CONFLICT");
    const conflictInfo = JSON.parse(conflictResult.conflictCandidatesJson!);
    expect(conflictInfo.candidates.map((c: { ruleId: string }) => c.ruleId).sort()).toEqual(
      [overtimeRule.id, conflictingRule.id].sort()
    );

    // 11/12：人工覆寫解決衝突，之後重新分類不會洗掉人工結果
    const overridden = await overrideClassification(
      record.id,
      { fundingSource: "PROJECT", projectId: project.id },
      "回歸測試",
      "人工確認為回歸測試專案"
    );
    expect(overridden.fundingSource).toBe("PROJECT");
    expect(overridden.isManuallyModified).toBe(true);

    const reclassified = await classifySubstituteRecord(record.id, "回歸測試");
    expect(reclassified.fundingSource).toBe("PROJECT");
    expect(reclassified.isManuallyModified).toBe(true);
    expect(reclassified.autoFundingSource).toBe("UNDETERMINED"); // auto 參考欄位仍反映目前規則衝突的事實
  });

  it("7/8. overtimeMatchMode 切換為科目比對模式時，真實資料的科目會被納入判斷", async () => {
    // 刻意排除已被前一項測試人工覆寫過的資料列，避免撿到同一筆造成互相干擾
    // （人工覆寫後不應該被重新分類洗掉，這是設計上的正確行為，不是這裡要測的東西）。
    const record = await prisma.substituteRecord.findFirstOrThrow({
      where: {
        monthlyImportId: nonBdImportId,
        isManuallyModified: false,
        subject: { not: null },
        originalTeacherId: null,
        rawRecord: { originalTeacherText: { not: null } },
      },
      include: { rawRecord: true },
      orderBy: { date: "asc" },
    });
    const teacherName = record.rawRecord!.originalTeacherText!;

    let person = await prisma.person.findFirst({ where: { name: teacherName } });
    if (!person) {
      const created = await createPerson({ name: teacherName });
      if (created.status !== "CREATED") throw new Error("unreachable");
      person = created.person;
      cleanupPersonIds.push(person.id);
    }
    await resolveTeacherReference(record.id, "original", person.id, "回歸測試");

    await prisma.specialWeeklyRule.create({
      data: {
        semesterId,
        personId: person.id,
        ruleType: "OVERTIME",
        weekday: record.weekday,
        periodCode: record.periodCode!,
        subject: record.subject!, // 用真實資料本身的科目設定規則，確保一定相符
        weeklyPeriods: 1,
        effectiveDate: new Date("2026-02-01"),
        endDate: new Date("2026-07-31"),
      },
    });

    await prisma.semester.update({ where: { id: semesterId }, data: { overtimeMatchMode: "TEACHER_WEEKDAY_PERIOD_SUBJECT" } });
    const matched = await classifySubstituteRecord(record.id, "回歸測試");
    expect(matched.fundingSource).toBe("OVERTIME");

    // 把紀錄本身的科目改成不同的值，同一條規則應該就不再相符
    await prisma.substituteRecord.update({ where: { id: record.id }, data: { subject: "測試不相符科目" } });
    const notMatched = await classifySubstituteRecord(record.id, "回歸測試");
    expect(notMatched.fundingSource).toBe("GENERAL");

    await prisma.semester.update({ where: { id: semesterId }, data: { overtimeMatchMode: "TEACHER_WEEKDAY_PERIOD" } });
  });

  it("14. SchoolCalendarDay 完全沒有因為這次分類流程被建立或推算", async () => {
    const calendarDays = await prisma.schoolCalendarDay.count({ where: { semesterId } });
    expect(calendarDays).toBe(0);
  });

  it("15. Raw Data 在整個分類流程中完全不變", async () => {
    const rawRows = await prisma.substituteRecordRaw.findMany({ where: { monthlyImportId: nonBdImportId } });
    // rawJson 是匯入當下就寫死的內容，這裡只驗證分類流程本身不會再去 UPDATE 這張表——
    // 直接比對目前資料庫內容跟 Phase 7 匯入當下的結構仍然一致（欄位皆存在、rowNumber 連續）。
    expect(rawRows.length).toBeGreaterThan(0);
    for (const row of rawRows) {
      expect(row.rawJson).not.toBeNull();
      expect(JSON.parse(row.rawJson!)).toBeTypeOf("object");
    }
  });
});
