// Phase 1 示範種子資料。
// 目的：證明資料模型「不寫死學年度／金額／規則」，並示範歷史資料保留、
// 原始/處理後資料分離、單日例外優先於每週規則等關鍵設計。
// 這裡出現的所有數字（405、133、1399、700……）都只是「範例資料」，
// 不是程式邏輯，之後隨時可以在 fee_rules 新增/修改資料列來調整，不需要改程式碼。

import { PrismaClient } from "@prisma/client";

const prisma = new PrismaClient();

async function main() {
  console.log("清除舊種子資料...");
  await prisma.changeLog.deleteMany();
  await prisma.substituteRecord.deleteMany();
  await prisma.substituteRecordRaw.deleteMany();
  await prisma.monthlyLock.deleteMany();
  await prisma.monthlyImportError.deleteMany();
  await prisma.monthlyImport.deleteMany();
  await prisma.schoolCalendarDay.deleteMany();
  await prisma.specialDateRule.deleteMany();
  await prisma.specialWeeklyRule.deleteMany();
  await prisma.project.deleteMany();
  await prisma.feeRule.deleteMany();
  await prisma.semesterRole.deleteMany();
  await prisma.personCode.deleteMany();
  await prisma.bdClassificationRule.deleteMany();
  await prisma.periodSlot.deleteMany();
  await prisma.person.deleteMany();
  await prisma.semester.deleteMany();

  // ---- 節次代碼表（section 10）----
  const periodDefs = [
    { code: "EARLY_STUDY", displayName: "早自修", sortOrder: 1, isTeachingPeriod: false },
    { code: "P1", displayName: "第1節", sortOrder: 2, isTeachingPeriod: true },
    { code: "P2", displayName: "第2節", sortOrder: 3, isTeachingPeriod: true },
    { code: "P3", displayName: "第3節", sortOrder: 4, isTeachingPeriod: true },
    { code: "P4", displayName: "第4節", sortOrder: 5, isTeachingPeriod: true },
    { code: "LUNCH", displayName: "午休", sortOrder: 6, isTeachingPeriod: false },
    { code: "P5", displayName: "第5節", sortOrder: 7, isTeachingPeriod: true },
    { code: "P6", displayName: "第6節", sortOrder: 8, isTeachingPeriod: true },
    { code: "P7", displayName: "第7節", sortOrder: 9, isTeachingPeriod: true },
  ];
  for (const p of periodDefs) {
    await prisma.periodSlot.create({ data: p });
  }

  // ---- BD/非BD 判斷規則（已確認：代碼 B、D → 編制內）----
  await prisma.bdClassificationRule.createMany({
    data: [
      { codeValue: "B", isBd: true, note: "已確認規則" },
      { codeValue: "D", isBd: true, note: "已確認規則" },
    ],
  });

  // ---- 學年度／學期：刻意建立跨兩個學年度、三個學期，證明沒有寫死 115 ----
  const semesterDefs = [
    { schoolYear: 115, term: 1, startDate: "2026-08-31", endDate: "2027-01-20", status: "ENDED" as const, isCurrent: false },
    { schoolYear: 115, term: 2, startDate: "2027-02-01", endDate: "2027-07-15", status: "ACTIVE" as const, isCurrent: true },
    { schoolYear: 116, term: 1, startDate: "2027-08-30", endDate: "2028-01-19", status: "NOT_STARTED" as const, isCurrent: false },
  ];
  const semesters: Record<string, Awaited<ReturnType<typeof prisma.semester.create>>> = {};
  for (const s of semesterDefs) {
    const created = await prisma.semester.create({
      data: {
        schoolYear: s.schoolYear,
        term: s.term,
        startDate: new Date(s.startDate),
        endDate: new Date(s.endDate),
        status: s.status,
        isCurrent: s.isCurrent,
      },
    });
    semesters[`${s.schoolYear}-${s.term}`] = created;
  }
  const sem115_1 = semesters["115-1"];
  const sem115_2 = semesters["115-2"];

  // ---- 人員：包含一位已離校教師，測試歷史資料是否保留 ----
  const chenXinQi = await prisma.person.create({
    data: { name: "陳心啓", payrollCode: "T1001", enrollmentStatus: "ENROLLED" },
  });
  const xuBiLing = await prisma.person.create({
    data: { name: "徐碧苓", payrollCode: "T1002", enrollmentStatus: "ENROLLED" },
  });
  const suZhenHui = await prisma.person.create({
    data: { name: "蘇珍慧", payrollCode: "T1003", enrollmentStatus: "ENROLLED" },
  });
  const wangSomeone = await prisma.person.create({
    data: {
      name: "王○○",
      payrollCode: "T0099",
      enrollmentStatus: "NOT_ENROLLED",
      enrollDate: new Date("2015-08-01"),
      leaveDate: new Date("2026-07-31"),
      note: "115學年度離校，僅供歷史資料查詢測試",
    },
  });

  await prisma.personCode.createMany({
    data: [
      { personId: chenXinQi.id, schoolYear: 115, categoryName: "正式教師", categoryCode: "B" },
      { personId: xuBiLing.id, schoolYear: 115, categoryName: "代理教師", categoryCode: "D" },
      { personId: suZhenHui.id, schoolYear: 115, categoryName: "正式教師", categoryCode: "B" },
      { personId: wangSomeone.id, schoolYear: 114, categoryName: "正式教師", categoryCode: "B" },
    ],
  });

  await prisma.semesterRole.create({
    data: { personId: chenXinQi.id, semesterId: sem115_1.id, roleType: "SUBJECT_TEACHER", roleDetail: "自然科任" },
  });
  await prisma.semesterRole.create({
    data: { personId: wangSomeone.id, semesterId: sem115_1.id, roleType: "HOMEROOM_TEACHER", roleDetail: "六年甲班導師" },
  });

  // ---- 費用規則：同一費用類型隨時間有兩個版本，證明金額是資料而非常數 ----
  await prisma.feeRule.createMany({
    data: [
      { semesterId: sem115_1.id, feeType: "SUBSTITUTE_PERIOD", amount: 405, effectiveDate: new Date("2026-08-31") },
      { semesterId: sem115_1.id, feeType: "OVERTIME_PERIOD", amount: 405, effectiveDate: new Date("2026-08-31") },
      { semesterId: sem115_1.id, feeType: "HOMEROOM_SUBSTITUTE", amount: 133, effectiveDate: new Date("2026-08-31"), note: "第一版僅存金額，未串接自動判斷邏輯" },
      { semesterId: sem115_1.id, feeType: "DAILY_WAGE", amount: 1399, effectiveDate: new Date("2026-08-31"), note: "第一版僅存金額，未串接自動判斷邏輯" },
      { semesterId: sem115_1.id, feeType: "HALF_DAY_WAGE", amount: 700, effectiveDate: new Date("2026-08-31"), note: "第一版僅存金額，未串接自動判斷邏輯" },
      // 示範同一費用類型的新版本（例如政府調整代課鐘點費），115-2 生效
      { semesterId: sem115_2.id, feeType: "SUBSTITUTE_PERIOD", amount: 420, effectiveDate: new Date("2027-02-01"), note: "示範費用調漲版本，不覆蓋舊資料" },
    ],
  });

  // ---- 專案 ----
  const guidanceProject = await prisma.project.create({
    data: { semesterId: sem115_1.id, name: "輔導團減課" },
  });

  // ---- 每週固定規則：超鐘點（陳心啓／星期二／第2節）與專案（蘇珍慧／星期三／第5節）----
  await prisma.specialWeeklyRule.create({
    data: {
      semesterId: sem115_1.id,
      personId: chenXinQi.id,
      ruleType: "OVERTIME",
      weekday: "TUE",
      periodCode: "P2",
      weeklyPeriods: 1,
      effectiveDate: new Date("2026-08-31"),
      endDate: new Date("2027-01-20"),
    },
  });
  await prisma.specialWeeklyRule.create({
    data: {
      semesterId: sem115_1.id,
      personId: suZhenHui.id,
      ruleType: "PROJECT",
      projectId: guidanceProject.id,
      weekday: "WED",
      periodCode: "P5",
      weeklyPeriods: 1,
      effectiveDate: new Date("2026-08-31"),
      endDate: new Date("2027-01-20"),
    },
  });

  // ---- 單日例外：優先於每週規則（對應 PRD 第13、46節範例）----
  await prisma.specialDateRule.create({
    data: {
      semesterId: sem115_1.id,
      date: new Date("2026-06-30"),
      personId: chenXinQi.id,
      periodCode: "P2",
      originalClassificationNote: "一般公費",
      overrideClassification: "OVERTIME",
      note: "本日特殊安排",
    },
  });

  // ---- 學校上課日曆（示範幾天，非全月）----
  await prisma.schoolCalendarDay.createMany({
    data: [
      { semesterId: sem115_1.id, date: new Date("2026-09-01"), weekday: "TUE", isTeachingDay: true },
      { semesterId: sem115_1.id, date: new Date("2026-09-28"), weekday: "MON", isTeachingDay: false, note: "教師節補假（示範國定假日排除）" },
      { semesterId: sem115_1.id, date: new Date("2026-06-30"), weekday: "TUE", isTeachingDay: true },
    ],
  });

  // ---- 每月匯入＋原始/處理後代課紀錄（示範 raw 與 processed 分離）----
  const juneImport = await prisma.monthlyImport.create({
    data: {
      semesterId: sem115_1.id,
      year: 2026,
      month: 6,
      fileName: "公費代課明細_202606.xlsx",
      importedBy: "教學組",
      totalCount: 1,
      successCount: 1,
      errorCount: 0,
    },
  });
  const juneRaw = await prisma.substituteRecordRaw.create({
    data: {
      monthlyImportId: juneImport.id,
      rowNumber: 1,
      originalTeacherText: "陳心啓",
      dateText: "06-30(二) 09:10 ~ 10:00",
      leaveTypeText: "公假",
      periodText: "第2節",
      classText: "五年甲班",
      subjectText: "自然",
      substituteTeacherText: "徐碧苓",
      teacherCertText: "有",
      payGradeText: "D",
    },
  });
  const juneProcessed = await prisma.substituteRecord.create({
    data: {
      rawRecordId: juneRaw.id,
      monthlyImportId: juneImport.id,
      originalTeacherId: chenXinQi.id,
      substituteTeacherId: xuBiLing.id,
      date: new Date("2026-06-30"),
      weekday: "TUE",
      periodCode: "P2",
      className: "五年甲班",
      subject: "自然",
      leaveType: "公假",
      periodCount: 1,
      staffType: "NON_BD",
      fundingSource: "OVERTIME",
      unitPrice: 405,
      amount: 405,
      classificationMethod: "DATE_EXCEPTION",
      note: "對應 6/30 單日例外設定",
    },
  });

  // ---- 每月鎖定：6月已鎖定（歷史月份），9月待確認 ----
  await prisma.monthlyLock.create({
    data: {
      semesterId: sem115_1.id,
      year: 2026,
      month: 6,
      status: "LOCKED",
      lockedAt: new Date("2026-07-05"),
      lockedBy: "教學組長",
    },
  });
  await prisma.monthlyLock.create({
    data: {
      semesterId: sem115_1.id,
      year: 2026,
      month: 9,
      status: "PENDING_REVIEW",
    },
  });

  // ---- 修改紀錄：示範人工覆蓋自動判斷會留痕 ----
  await prisma.changeLog.create({
    data: {
      tableName: "substitute_records",
      recordId: juneProcessed.id,
      fieldName: "fundingSource",
      oldValue: "GENERAL",
      newValue: "OVERTIME",
      changedBy: "教學組",
      reason: "依 6/30 單日例外設定調整為合理員額增授節數超鐘點",
    },
  });

  console.log("種子資料建立完成。");
}

main()
  .catch((e) => {
    console.error(e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
