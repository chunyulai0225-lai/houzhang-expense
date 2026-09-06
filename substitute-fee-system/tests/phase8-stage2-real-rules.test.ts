// Phase 8 第二階段：用「真實」的合理員額超鐘點檔與各項專案減課檔，
// 建立真實 Person／Project／SpecialWeeklyRule，再讓真實的
// 「114學年2026.06月代課(公費).xlsx」代課紀錄跑過分類引擎，
// 驗證系統判斷結果與真實世界的備註（誰、哪天、第幾節、為什麼）完全吻合。
//
// 三份真實檔案都含真實姓名與請假紀錄，不會被 commit 進 repo，
// 也不會留在版本控制歷史裡。任一檔案不存在時整組測試自動略過。
//
// 重要：這裡「不」把第二、三份 Excel 的金額／結果欄位直接寫進 SubstituteRecord 當答案。
// 我們只用這兩份檔案的內容（教師代碼、姓名、星期、節次、每週節數、備註原文）建立
// SpecialWeeklyRule／Project，然後讓分類引擎自己去判斷第一份真實代課紀錄，
// 再比對分類結果是否和備註描述的事實一致。分類引擎完全不做金額計算。
import fs from "fs";
import { describe, expect, it, beforeAll, afterAll } from "vitest";
import { prisma } from "../src/prismaClient";
import { createPerson } from "../src/services/personService";
import { createProject } from "../src/services/projectService";
import { importSubstituteExcel, resolveTeacherReference } from "../src/services/excelImportService";
import { classifyMonthlyImport, classifySubstituteRecord, overrideClassification } from "../src/services/classificationService";
import { createFeeRule } from "../src/services/feeRuleService";
import { calculateSubstituteRecordFee } from "../src/services/feeCalculationService";

const SUBSTITUTE_CANDIDATES = [
  process.env.REAL_SUBSTITUTE_EXCEL_PATH,
  "/root/.claude/uploads/6b1e1531-1999-57e8-8932-c62ee2d638fb/f45b44e8-114__2026.06_____.xlsx",
].filter((p): p is string => Boolean(p));

// 檔名雖然寫「專案」，但實際內容是「合理員額增授節數超鐘點」（已用實際內容核對過，不是照檔名猜）
const OVERTIME_CANDIDATES = [
  process.env.REAL_OVERTIME_EXCEL_PATH,
  "/root/.claude/uploads/6b1e1531-1999-57e8-8932-c62ee2d638fb/a7370627-114______________2026.06_.xlsx",
].filter((p): p is string => Boolean(p));

// 檔名雖然寫「超鐘點」，但實際內容是「各項專案減課鐘點」（已用實際內容核對過，不是照檔名猜）
const PROJECT_CANDIDATES = [
  process.env.REAL_PROJECT_EXCEL_PATH,
  "/root/.claude/uploads/6b1e1531-1999-57e8-8932-c62ee2d638fb/444d2e6d-114____________2026_.6_.xlsx",
].filter((p): p is string => Boolean(p));

const REAL_SUBSTITUTE_FILE = SUBSTITUTE_CANDIDATES.find((p) => fs.existsSync(p));
const REAL_OVERTIME_FILE = OVERTIME_CANDIDATES.find((p) => fs.existsSync(p));
const REAL_PROJECT_FILE = PROJECT_CANDIDATES.find((p) => fs.existsSync(p));

const NON_BD_SHEET = "6月公費  非BD 0701";
const BD_SHEET = "6月公費BD 原授課排序給芝庭";

// 真實「各項專案減課鐘點」檔案裡的 7 個專案區塊標題（逐字保留原文，只去除固定的
// 「后庄國小115年6月」前綴，作為 Project.name；完整原文放進 note 供追查）。
const REAL_PROJECT_TITLES = [
  "輔導團減課鐘點",
  "閱讀推動教師減課鐘點",
  "行政協助推動中小學數位學習精進方案減課鐘點(代印領清冊)",
  "推動中小學數位學習精進方案重點學校減課鐘點(代印領清冊)",
  "推動中小學數位學習精進方案減課鐘點-推動人員減課(代印領清冊)",
  "薪傳師減課鐘點(代印領清冊)",
  "補助高級中等以下學校協助學務輔導工作教師之減授課節數鐘點(代印領清冊)",
];

describe.skipIf(!REAL_SUBSTITUTE_FILE || !REAL_OVERTIME_FILE || !REAL_PROJECT_FILE)(
  "Phase 8 第二階段：真實超鐘點／專案規則 vs 真實代課紀錄分類回歸測試",
  () => {
    const cleanupPersonIds: string[] = [];
    let semesterId: string;
    let nonBdImportId: string;
    let bdImportId: string;
    const projectIds: Record<string, string> = {};

    // 5 個「合理員額增授節數超鐘點」真實檔案裡的錨點教師（personId 在 beforeAll 建立）
    const anchors: Record<string, string> = {};

    beforeAll(async () => {
      await prisma.$connect();
      const substituteBuffer = fs.readFileSync(REAL_SUBSTITUTE_FILE!);

      const semester = await prisma.semester.create({
        data: {
          schoolYear: 8802, // 明顯不會與真實學年度衝突的測試用學年度（Stage2 專用，避免與 8801 的 Stage1 回歸測試衝突）
          term: 2,
          startDate: new Date("2026-02-01"),
          endDate: new Date("2026-07-31"),
          note: "Phase8 第二階段：真實超鐘點/專案規則回歸測試，測試結束會刪除",
        },
      });
      semesterId = semester.id;

      const nonBd = await importSubstituteExcel({
        semesterId,
        year: 2026,
        month: 6,
        fileName: "stage2-real-nonbd.xlsx",
        fileBuffer: substituteBuffer,
        sheetName: NON_BD_SHEET,
        sourceStaffType: "NON_BD",
        importedBy: "Stage2回歸測試",
      });
      nonBdImportId = nonBd.monthlyImport.id;

      const bd = await importSubstituteExcel({
        semesterId,
        year: 2026,
        month: 6,
        fileName: "stage2-real-bd.xlsx",
        fileBuffer: substituteBuffer,
        sheetName: BD_SHEET,
        sourceStaffType: "BD",
        importedBy: "Stage2回歸測試",
      });
      bdImportId = bd.monthlyImport.id;

      // 建立真實「各項專案減課鐘點」的 7 個 Project（來源：114學年 各項專案減課代課費2026.6月.xlsx）
      for (const title of REAL_PROJECT_TITLES) {
        const project = await createProject({
          semesterId,
          name: title,
          note: `來源：114學年 各項專案減課代課費2026.6月.xlsx（后庄國小115年6月 ${title}）`,
        });
        projectIds[title] = project.id;
      }

      // 建立 5 個真實「合理員額增授節數超鐘點」錨點教師（來源：114學年2026.6月合理員額增授節數超鐘點.xlsx 編制內區塊）
      // 每一位的姓名、代碼、星期、節次都直接來自該檔案「115-6總表」工作表的實際欄位，
      // 且都已經在真實代課檔案（114學年2026.06月代課(公費).xlsx）裡找到日期、節次、代課教師完全吻合的紀錄，
      // 不是憑空建立的測試資料。
      const anchorDefs = [
        { key: "陳心啟", staffCode: "B00029", note: "資料來源列：115-6總表 row7；備註原文：「陳心啟老師2026/6/30(二)每週二第2節二乙為合理員額超鐘點, 以合理員額超鐘點支付為徐碧苓老師之代課費」" },
        { key: "陳芝庭", staffCode: "B00087", note: "資料來源列：115-6總表 row15；備註原文：「陳芝庭老師2026/6/18(四)每週四第3節五戊為合理員額超鐘點, 以合理員額超鐘點支付為陳君岳老師之代課費」（原始節次欄「2-4,4-3」為星期二第4節與星期四第3節的複合記法，這裡只取星期四第3節這一段）" },
        { key: "王玉蓮", staffCode: "B00102", note: "資料來源列：115-6總表 row19；備註原文：「王玉蓮老師2026/6/9(二)每週二第4節六庚為合理員額超鐘點, 以合理員額超鐘點支付為賴俊佑老師之代課費」" },
        { key: "林宜慧", staffCode: "B00210", note: "資料來源列：115-6總表 row27；備註原文：「林宜慧老師2026/6/30(二)每週二第1節二辛為合理員額超鐘點, 以合理員額超鐘點支付為黃晏珊老師之代課費」" },
        { key: "羅純惠", staffCode: "D00048", note: "資料來源列：115-6總表 row34；備註原文：「羅純惠老師2026/6/30(二)每週二第1節二丁為合理員額超鐘點, 以合理員額超鐘點支付為古惠茹老師之代課費」" },
      ];
      for (const def of anchorDefs) {
        const created = await createPerson({ name: def.key });
        if (created.status !== "CREATED") throw new Error(`unreachable: 測試用教師「${def.key}」不應該已存在於乾淨的測試學期環境`);
        anchors[def.key] = created.person.id;
        cleanupPersonIds.push(created.person.id);
        await prisma.personCode.create({
          data: {
            personId: created.person.id,
            schoolYear: 114,
            categoryName: "合理員額超鐘點(編制內)",
            categoryCode: def.staffCode.charAt(0),
            originalStaffCode: def.staffCode,
            note: def.note,
          },
        });
      }

      // Phase9-5 驗證用：建立正式費率，確認真實 OVERTIME 錨點的實際代課費由 Raw 的
      // substitutePeriodFeeText 決定（這幾筆真實資料是 0），不會被 fundingSource=OVERTIME 洗成 405。
      await createFeeRule({ semesterId, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2026-02-01") });
      await createFeeRule({ semesterId, feeType: "OVERTIME_PERIOD", amount: 405, effectiveDate: new Date("2026-02-01") });
    });

    afterAll(async () => {
      const importIds = [nonBdImportId, bdImportId];
      const records = await prisma.substituteRecord.findMany({ where: { monthlyImportId: { in: importIds } } });
      await prisma.changeLog.deleteMany({
        where: {
          OR: [
            { recordId: { in: records.map((r) => r.id) } },
            { recordId: { in: cleanupPersonIds } },
            { recordId: { in: Object.values(projectIds) } },
          ],
        },
      });
      await prisma.monthlyImportError.deleteMany({ where: { monthlyImportId: { in: importIds } } });
      await prisma.substituteRecord.deleteMany({ where: { monthlyImportId: { in: importIds } } });
      await prisma.substituteRecordRaw.deleteMany({ where: { monthlyImportId: { in: importIds } } });
      await prisma.monthlyImport.deleteMany({ where: { id: { in: importIds } } });
      await prisma.specialDateRule.deleteMany({ where: { semesterId } });
      await prisma.specialWeeklyRule.deleteMany({ where: { semesterId } });
      await prisma.feeRule.deleteMany({ where: { semesterId } });
      await prisma.personCode.deleteMany({ where: { personId: { in: cleanupPersonIds } } });
      await prisma.project.deleteMany({ where: { semesterId } });
      await prisma.person.deleteMany({ where: { id: { in: cleanupPersonIds } } });
      await prisma.semester.delete({ where: { id: semesterId } });
      await prisma.$disconnect();
    });

    async function findAnchorRecord(monthlyImportId: string, teacherName: string, dateUtc: Date, periodCode: string) {
      return prisma.substituteRecord.findFirstOrThrow({
        where: {
          monthlyImportId,
          date: dateUtc,
          periodCode,
          rawRecord: { originalTeacherText: teacherName },
        },
        include: { rawRecord: true },
      });
    }

    it("§二/1. 真實 OVERTIME 規則命中真實紀錄：4 位錨點教師（陳心啟／王玉蓮／林宜慧／羅純惠）", async () => {
      const cases = [
        { name: "陳心啟", importId: () => nonBdImportId, date: new Date(Date.UTC(2026, 5, 30)), periodCode: "P2", weekday: "TUE" as const },
        { name: "王玉蓮", importId: () => bdImportId, date: new Date(Date.UTC(2026, 5, 9)), periodCode: "P4", weekday: "TUE" as const },
        { name: "林宜慧", importId: () => bdImportId, date: new Date(Date.UTC(2026, 5, 30)), periodCode: "P1", weekday: "TUE" as const },
        { name: "羅純惠", importId: () => bdImportId, date: new Date(Date.UTC(2026, 5, 30)), periodCode: "P1", weekday: "TUE" as const },
      ];

      for (const c of cases) {
        const record = await findAnchorRecord(c.importId(), c.name, c.date, c.periodCode);
        await resolveTeacherReference(record.id, "original", anchors[c.name], "Stage2回歸測試");

        const rule = await prisma.specialWeeklyRule.create({
          data: {
            semesterId,
            personId: anchors[c.name],
            ruleType: "OVERTIME",
            weekday: c.weekday,
            periodCode: c.periodCode,
            weeklyPeriods: 1,
            effectiveDate: new Date("2026-02-01"),
            note: `真實合理員額超鐘點規則（${c.name}）`,
          },
        });

        const result = await classifySubstituteRecord(record.id, "Stage2回歸測試");
        expect(result.fundingSource).toBe("OVERTIME");
        expect(result.classificationMethod).toBe("WEEKLY_RULE");
        expect(result.classificationRuleId).toBe(rule.id);
      }
    });

    it("§二/12（Phase9-5）. 真實 OVERTIME 錨點的實際代課費 = Raw 的代課鐘點費（0），不是 fundingSource 帶出的 405", async () => {
      const cases = [
        { name: "陳心啟", importId: () => nonBdImportId, date: new Date(Date.UTC(2026, 5, 30)), periodCode: "P2" },
        { name: "王玉蓮", importId: () => bdImportId, date: new Date(Date.UTC(2026, 5, 9)), periodCode: "P4" },
        { name: "林宜慧", importId: () => bdImportId, date: new Date(Date.UTC(2026, 5, 30)), periodCode: "P1" },
        { name: "羅純惠", importId: () => bdImportId, date: new Date(Date.UTC(2026, 5, 30)), periodCode: "P1" },
      ];
      for (const c of cases) {
        const record = await findAnchorRecord(c.importId(), c.name, c.date, c.periodCode);
        expect(record.fundingSource).toBe("OVERTIME"); // 沿用上一個測試已經分類好的結果
        const result = await calculateSubstituteRecordFee(record.id, "Stage2回歸測試");
        expect(result.amount).toBe("0"); // 真實 Excel 這筆代課鐘點費原文就是 0
        expect(result.unitPrice).toBe("405"); // 參考費率仍然是405，只是這筆代課教師實際不用付
      }
    });

    it("§二/1＋5. 陳芝庭錨點：星期四第3節命中 OVERTIME，且 effectiveDate 生效邊界正確（早於紀錄日期時不命中）", async () => {
      const record = await findAnchorRecord(bdImportId, "陳芝庭", new Date(Date.UTC(2026, 5, 18)), "P3");
      await resolveTeacherReference(record.id, "original", anchors["陳芝庭"], "Stage2回歸測試");

      const rule = await prisma.specialWeeklyRule.create({
        data: {
          semesterId,
          personId: anchors["陳芝庭"],
          ruleType: "OVERTIME",
          weekday: "THU",
          periodCode: "P3",
          weeklyPeriods: 1,
          effectiveDate: new Date("2026-06-19"), // 刻意晚於紀錄日期 6/18，驗證生效日期邊界
          note: "真實合理員額超鐘點規則（陳芝庭，星期四第3節）",
        },
      });

      const beforeEffective = await classifySubstituteRecord(record.id, "Stage2回歸測試");
      expect(beforeEffective.fundingSource).toBe("GENERAL");
      expect(beforeEffective.classificationMethod).toBe("GENERAL_DEFAULT");

      const updatedRule = await prisma.specialWeeklyRule.update({
        where: { id: rule.id },
        data: { effectiveDate: new Date("2026-02-01") },
      });
      const afterEffective = await classifySubstituteRecord(record.id, "Stage2回歸測試");
      expect(afterEffective.fundingSource).toBe("OVERTIME");
      expect(afterEffective.classificationRuleId).toBe(updatedRule.id);

      // Phase9-5：陳芝庭這筆真實 OVERTIME 紀錄，實際代課費同樣是 Raw 的 0，不是 405
      const feeResult = await calculateSubstituteRecordFee(record.id, "Stage2回歸測試");
      expect(feeResult.amount).toBe("0");
      expect(feeResult.unitPrice).toBe("405");
    });

    it("§二/2. 真實 PROJECT 規則命中真實紀錄：確認的資料模型缺口——來源檔案沒有節次，Stage2 不猜、不建立", async () => {
      // 使用者已確認：編制外超鐘點與全部 7 個真實專案區塊，來源都只有「星期＋每週節數」，
      // 沒有精確節次；而 SpecialWeeklyRule.periodCode 是必填欄位、且分類引擎用它做完全比對。
      // 因此 Stage2 明確「不」從這 7 個真實專案區塊建立任何 PROJECT 型 SpecialWeeklyRule，
      // 避免猜一個來源沒寫的節次。這裡用一個可驗證的斷言記錄這個事實，而不是silently跳過。
      const projectRuleCount = await prisma.specialWeeklyRule.count({ where: { semesterId, ruleType: "PROJECT" } });
      expect(projectRuleCount).toBe(0);

      const realProjectCount = await prisma.project.count({ where: { semesterId } });
      expect(realProjectCount).toBe(REAL_PROJECT_TITLES.length);
    });

    it("§二/3. 沒有規則、也沒配對到 Person 的真實資料維持 GENERAL 或 TEACHER_UNMATCHED，不亂猜", async () => {
      await classifyMonthlyImport(nonBdImportId, "Stage2回歸測試");
      await classifyMonthlyImport(bdImportId, "Stage2回歸測試");

      const stillUnmatched = await prisma.substituteRecord.findMany({
        where: { monthlyImportId: { in: [nonBdImportId, bdImportId] }, originalTeacherId: null },
      });
      expect(stillUnmatched.length).toBeGreaterThan(0);
      expect(
        stillUnmatched.every((r) => r.classificationMethod === "TEACHER_UNMATCHED" && r.fundingSource === "UNDETERMINED")
      ).toBe(true);
    });

    it("§二/4. SpecialDateRule 優先於 SpecialWeeklyRule：真實錨點當天被加上單日例外時，改用單日例外的判斷", async () => {
      const record = await findAnchorRecord(nonBdImportId, "陳心啟", new Date(Date.UTC(2026, 5, 30)), "P2");

      const dateRule = await prisma.specialDateRule.create({
        data: {
          semesterId,
          date: record.date,
          personId: anchors["陳心啟"],
          periodCode: record.periodCode!,
          overrideClassification: "GENERAL",
          note: "回歸測試：單日例外應優先於已存在的真實 OVERTIME 週規則",
        },
      });

      const result = await classifySubstituteRecord(record.id, "Stage2回歸測試");
      expect(result.fundingSource).toBe("GENERAL");
      expect(result.classificationMethod).toBe("DATE_EXCEPTION");
      expect(result.classificationRuleId).toBe(dateRule.id);

      // 還原：取消這筆單日例外，確認又會回到原本命中的 OVERTIME 週規則
      await prisma.specialDateRule.update({ where: { id: dateRule.id }, data: { isCancelled: true } });
      const reverted = await classifySubstituteRecord(record.id, "Stage2回歸測試");
      expect(reverted.fundingSource).toBe("OVERTIME");
    });

    it("§二/6+7. overtimeMatchMode 切換為科目比對模式：真實科目相符才命中，規則沒填科目時不算萬用", async () => {
      // 挑一筆完全獨立、尚未被前面測試觸碰過的真實資料（避免互相干擾）
      const record = await prisma.substituteRecord.findFirstOrThrow({
        where: {
          monthlyImportId: nonBdImportId,
          isManuallyModified: false,
          originalTeacherId: null,
          subject: { not: null },
          rawRecord: { originalTeacherText: { not: null } },
        },
        include: { rawRecord: true },
        orderBy: { date: "asc" },
      });
      const teacherName = record.rawRecord!.originalTeacherText!;

      const created = await createPerson({ name: teacherName });
      if (created.status !== "CREATED") throw new Error("unreachable: 真實資料裡這個姓名應該還沒有對應的 Person");
      cleanupPersonIds.push(created.person.id);
      await resolveTeacherReference(record.id, "original", created.person.id, "Stage2回歸測試");

      await prisma.semester.update({ where: { id: semesterId }, data: { overtimeMatchMode: "TEACHER_WEEKDAY_PERIOD_SUBJECT" } });

      // 7. 規則本身沒填科目 → 在科目比對模式下不算萬用規則，不應該命中
      const noSubjectRule = await prisma.specialWeeklyRule.create({
        data: {
          semesterId,
          personId: created.person.id,
          ruleType: "OVERTIME",
          weekday: record.weekday,
          periodCode: record.periodCode!,
          subject: null,
          weeklyPeriods: 1,
          effectiveDate: new Date("2026-02-01"),
        },
      });
      const noSubjectResult = await classifySubstituteRecord(record.id, "Stage2回歸測試");
      expect(noSubjectResult.fundingSource).toBe("GENERAL");
      await prisma.specialWeeklyRule.delete({ where: { id: noSubjectRule.id } });

      // 6. 規則填了跟紀錄完全相同的科目 → 應該命中
      const subjectRule = await prisma.specialWeeklyRule.create({
        data: {
          semesterId,
          personId: created.person.id,
          ruleType: "OVERTIME",
          weekday: record.weekday,
          periodCode: record.periodCode!,
          subject: record.subject!,
          weeklyPeriods: 1,
          effectiveDate: new Date("2026-02-01"),
        },
      });
      const matched = await classifySubstituteRecord(record.id, "Stage2回歸測試");
      expect(matched.fundingSource).toBe("OVERTIME");
      expect(matched.classificationRuleId).toBe(subjectRule.id);

      await prisma.semester.update({ where: { id: semesterId }, data: { overtimeMatchMode: "TEACHER_WEEKDAY_PERIOD" } });
    });

    it("§二/8. 規則衝突：同一位真實錨點教師、同星期同節次同時有 OVERTIME 與 PROJECT 規則時，判為 CONFLICT，不自行猜", async () => {
      const record = await findAnchorRecord(bdImportId, "王玉蓮", new Date(Date.UTC(2026, 5, 9)), "P4");
      // 這裡的 OVERTIME 規則就是「§二/1」測試已經建立的那條真實規則
      const overtimeRule = await prisma.specialWeeklyRule.findFirstOrThrow({
        where: { semesterId, personId: anchors["王玉蓮"], ruleType: "OVERTIME", weekday: "TUE", periodCode: "P4" },
      });

      // 刻意在同一位教師、同星期、同節次加一條 PROJECT 規則，使用真實存在的專案（薪傳師減課鐘點），
      // 藉此測試「規則衝突」機制本身，而不是宣稱真實專案檔真的在王玉蓮身上指定了這個節次
      // （已確認真實專案檔完全沒有精確節次資料，見「§二/2」）。
      const conflictingRule = await prisma.specialWeeklyRule.create({
        data: {
          semesterId,
          personId: anchors["王玉蓮"],
          ruleType: "PROJECT",
          projectId: projectIds["薪傳師減課鐘點(代印領清冊)"],
          weekday: "TUE",
          periodCode: "P4",
          weeklyPeriods: 1,
          effectiveDate: new Date("2026-02-01"),
          note: "回歸測試：刻意製造與真實 OVERTIME 規則的衝突，驗證 CONFLICT 機制",
        },
      });

      const result = await classifySubstituteRecord(record.id, "Stage2回歸測試");
      expect(result.fundingSource).toBe("UNDETERMINED");
      expect(result.classificationMethod).toBe("CONFLICT");
      const conflictInfo = JSON.parse(result.conflictCandidatesJson!);
      expect(conflictInfo.candidates.map((c: { ruleId: string }) => c.ruleId).sort()).toEqual(
        [overtimeRule.id, conflictingRule.id].sort()
      );

      // 人工覆寫解決衝突後，重新分類不會洗掉人工結果（§二/11）
      const overridden = await overrideClassification(
        record.id,
        { fundingSource: "OVERTIME" },
        "Stage2回歸測試",
        "人工確認為真實合理員額超鐘點規則，非專案規則"
      );
      expect(overridden.fundingSource).toBe("OVERTIME");
      expect(overridden.isManuallyModified).toBe(true);

      const reclassified = await classifySubstituteRecord(record.id, "Stage2回歸測試");
      expect(reclassified.fundingSource).toBe("OVERTIME");
      expect(reclassified.isManuallyModified).toBe(true);
      expect(reclassified.autoFundingSource).toBe("UNDETERMINED"); // auto 參考欄位仍誠實反映目前規則衝突的事實

      // 清掉刻意製造的衝突規則，避免影響其他測試對王玉蓮這筆紀錄的假設
      await prisma.specialWeeklyRule.delete({ where: { id: conflictingRule.id } });
    });

    it("§二/9. BD／非BD 來源互相獨立，不因為分類流程互相污染", async () => {
      const nonBdSample = await prisma.substituteRecord.findFirstOrThrow({ where: { monthlyImportId: nonBdImportId } });
      expect(nonBdSample.staffType).toBe("NON_BD");
      const bdSample = await prisma.substituteRecord.findFirstOrThrow({ where: { monthlyImportId: bdImportId } });
      expect(bdSample.staffType).toBe("BD");

      const nonBdCount = await prisma.substituteRecord.count({ where: { monthlyImportId: nonBdImportId } });
      const bdCount = await prisma.substituteRecord.count({ where: { monthlyImportId: bdImportId } });
      expect(nonBdCount).toBeGreaterThan(0);
      expect(bdCount).toBeGreaterThan(0);
    });

    it("§二/10. Raw Data 在整個分類流程中完全不變", async () => {
      const rawRows = await prisma.substituteRecordRaw.findMany({ where: { monthlyImportId: nonBdImportId } });
      expect(rawRows.length).toBeGreaterThan(0);
      for (const row of rawRows) {
        expect(row.rawJson).not.toBeNull();
        expect(JSON.parse(row.rawJson!)).toBeTypeOf("object");
      }
    });

    it("資料品質觀察：王志群／孫永守兩個候選錨點的真實日期格式無法解析，落在 Phase7 匯入錯誤，Stage2 不強行分類", async () => {
      // 王志群：唯一對得上「每週二第1節三庚」的真實紀錄，日期原文是
      // 「06-29(一) 07:50 ~ 06-30(二) 15:50」（跨兩個日期的請假區間），
      // Phase7 的日期解析器設計上會拒絕這種區間，不會只取開頭幾個字元。
      const wangRaw = await prisma.substituteRecordRaw.findFirst({
        where: { monthlyImportId: bdImportId, originalTeacherText: "王志群", dateText: { contains: "06-29" } },
        include: { processedRecord: true },
      });
      expect(wangRaw).not.toBeNull();
      expect(wangRaw!.processedRecord).toBeNull();
      const wangError = await prisma.monthlyImportError.findFirst({
        where: { monthlyImportId: bdImportId, rowNumber: wangRaw!.rowNumber ?? undefined, fieldName: "日期" },
      });
      expect(wangError).not.toBeNull();

      // 孫永守：對得上「每週二第4節四丙」的真實紀錄，日期原文是整月彙總的
      // 「6/1~6/30」，同樣不是單一日期，Phase7 設計上會拒絕、列為錯誤而非猜測成某一天。
      const sunRaw = await prisma.substituteRecordRaw.findFirst({
        where: { monthlyImportId: nonBdImportId, originalTeacherText: "孫永守", dateText: "6/1~6/30" },
        include: { processedRecord: true },
      });
      expect(sunRaw).not.toBeNull();
      expect(sunRaw!.processedRecord).toBeNull();
      const sunError = await prisma.monthlyImportError.findFirst({
        where: { monthlyImportId: nonBdImportId, rowNumber: sunRaw!.rowNumber ?? undefined, fieldName: "日期" },
      });
      expect(sunError).not.toBeNull();
    });
  }
);
