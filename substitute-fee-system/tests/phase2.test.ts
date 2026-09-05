// Phase 2 驗收測試：人員管理（人員列表／詳細資料／新增防重複／停用／PDF匯入架構）
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { prisma } from "../src/prismaClient";
import {
  createPerson,
  deactivatePerson,
  findPossiblePersonDuplicates,
  getPersonDetail,
  listPersons,
} from "../src/services/personService";
import {
  getCandidatePersonsForRow,
  listImportRows,
  resolveImportRow,
  startStaffDirectoryImport,
} from "../src/services/staffDirectoryImportService";

const cleanupPersonIds: string[] = [];

beforeAll(async () => {
  await prisma.$connect();
});

afterAll(async () => {
  await prisma.changeLog.deleteMany({ where: { recordId: { in: cleanupPersonIds } } });
  await prisma.personCode.deleteMany({ where: { personId: { in: cleanupPersonIds } } });
  await prisma.person.deleteMany({ where: { id: { in: cleanupPersonIds } } });
  await prisma.$disconnect();
});

describe("人員列表：預設只顯示在校人員，可切換查看", () => {
  it("預設（ENROLLED）不包含已離校的王○○，包含在校的陳心啓", async () => {
    const enrolled = await listPersons();
    expect(enrolled.some((p) => p.name === "王○○")).toBe(false);
    expect(enrolled.some((p) => p.name === "陳心啓")).toBe(true);
  });

  it("ALL 可以查到王○○；NOT_ENROLLED 只包含不在校人員", async () => {
    const all = await listPersons("ALL");
    expect(all.some((p) => p.name === "王○○")).toBe(true);

    const notEnrolled = await listPersons("NOT_ENROLLED");
    expect(notEnrolled.every((p) => p.enrollmentStatus === "NOT_ENROLLED")).toBe(true);
    expect(notEnrolled.some((p) => p.name === "王○○")).toBe(true);
  });
});

describe("人員詳細資料", () => {
  it("包含基本資料、代碼／身分類別、歷史學期職務與歷史代課紀錄", async () => {
    const chen = await prisma.person.findFirstOrThrow({ where: { name: "陳心啓" } });
    const detail = await getPersonDetail(chen.id);

    expect(detail.name).toBe("陳心啓");
    expect(detail.personCodes.some((c) => c.categoryCode === "B")).toBe(true);
    expect(detail.semesterRoles[0].semester.schoolYear).toBe(115);
    expect(detail.semesterRoles[0].roleDetail).toBe("自然科任");
    // Phase 1 種子資料裡，陳心啓是 6/30 那筆代課的原教師
    expect(detail.substituteRecordsAsOriginal.length).toBeGreaterThanOrEqual(1);
    expect(detail.substituteRecordsAsOriginal[0].fundingSource).toBe("OVERTIME");
  });

  it("查詢不存在的人員會丟出錯誤", async () => {
    await expect(getPersonDetail("does-not-exist")).rejects.toThrow();
  });
});

describe("重複人員保護", () => {
  it("同名人員第二次新增時會回傳待確認的既有人員清單，不會自動建立", async () => {
    const first = await createPerson({ name: "測試重複姓名甲" });
    expect(first.status).toBe("CREATED");
    if (first.status !== "CREATED") throw new Error("unreachable");
    cleanupPersonIds.push(first.person.id);

    const second = await createPerson({ name: "測試重複姓名甲" });
    expect(second.status).toBe("POSSIBLE_DUPLICATE");
    if (second.status !== "POSSIBLE_DUPLICATE") throw new Error("unreachable");
    expect(second.candidates.map((c) => c.id)).toEqual([first.person.id]);

    // 管理者確認後仍可強制建立第二個人（例如恰好真的是不同人）
    const forced = await createPerson({ name: "測試重複姓名甲" }, { forceCreate: true });
    expect(forced.status).toBe("CREATED");
    if (forced.status !== "CREATED") throw new Error("unreachable");
    cleanupPersonIds.push(forced.person.id);

    const duplicates = await findPossiblePersonDuplicates("測試重複姓名甲");
    expect(duplicates.length).toBe(2);
  });
});

describe("人員停用（離校）", () => {
  it("停用只改變狀態與離校日期，不刪除資料，並留下修改紀錄", async () => {
    const created = await createPerson({ name: "測試停用員工乙" });
    if (created.status !== "CREATED") throw new Error("unreachable");
    cleanupPersonIds.push(created.person.id);

    const updated = await deactivatePerson(created.person.id, {
      leaveDate: new Date("2026-08-01"),
      changedBy: "測試管理者",
      reason: "測試離校",
    });
    expect(updated.enrollmentStatus).toBe("NOT_ENROLLED");
    expect(updated.leaveDate?.toISOString().slice(0, 10)).toBe("2026-08-01");

    // 人員本身仍然存在（只是不在預設列表中）
    const stillExists = await prisma.person.findUnique({ where: { id: created.person.id } });
    expect(stillExists).not.toBeNull();

    const logs = await prisma.changeLog.findMany({
      where: { tableName: "persons", recordId: created.person.id },
    });
    expect(logs.length).toBe(1);
    expect(logs[0].oldValue).toBe("ENROLLED");
    expect(logs[0].newValue).toBe("NOT_ENROLLED");
    expect(logs[0].reason).toBe("測試離校");
  });
});

describe("教職員工代號 PDF 匯入架構", () => {
  it("姓名與既有人員相符時只列為候選，需管理者確認才會寫入 PersonCode", async () => {
    const wang = await prisma.person.findFirstOrThrow({ where: { name: "王○○" } });
    const codesBefore = await prisma.personCode.count({ where: { personId: wang.id } });

    const batch = await startStaffDirectoryImport({
      fileName: "教職員工代號_測試.pdf",
      importedBy: "測試",
      rows: [{ rowNumber: 1, employeeCode: "E9001", name: "王○○", departmentName: "課後照顧" }],
    });

    const row = batch.rows[0];
    expect(row.matchStatus).toBe("PENDING");
    const candidates = await getCandidatePersonsForRow(row.id);
    expect(candidates.map((c) => c.id)).toEqual([wang.id]);
    // 尚未確認前不應該有新的 PersonCode
    expect(await prisma.personCode.count({ where: { personId: wang.id } })).toBe(codesBefore);

    const resolved = await resolveImportRow(row.id, { action: "MATCH_EXISTING", personId: wang.id }, "測試管理者");
    expect(resolved.matchStatus).toBe("MATCHED_EXISTING");
    expect(resolved.matchedPersonId).toBe(wang.id);

    const newCode = await prisma.personCode.findFirst({
      where: { personId: wang.id, originalStaffCode: "E9001" },
    });
    expect(newCode?.categoryName).toBe("課後照顧");
    // categoryCode 暫不自動推導、也不要求手動輸入
    expect(newCode?.categoryCode).toBeNull();
  });

  it("找不到既有人員時不會自動建立，且不能重複處理同一列", async () => {
    const batch = await startStaffDirectoryImport({
      fileName: "教職員工代號_測試2.pdf",
      importedBy: "測試",
      rows: [{ rowNumber: 1, employeeCode: "E9002", name: "測試新進人員丙", departmentName: "增置教師" }],
    });
    const row = batch.rows[0];
    expect(await getCandidatePersonsForRow(row.id)).toEqual([]);

    const beforeCount = await prisma.person.count({ where: { name: "測試新進人員丙" } });
    expect(beforeCount).toBe(0);

    const resolved = await resolveImportRow(row.id, { action: "CREATE_NEW" }, "測試管理者");
    expect(resolved.matchStatus).toBe("CREATED_NEW");
    cleanupPersonIds.push(resolved.matchedPersonId!);

    const created = await prisma.person.findFirstOrThrow({ where: { name: "測試新進人員丙" } });
    expect(created.id).toBe(resolved.matchedPersonId);

    await expect(resolveImportRow(row.id, { action: "IGNORE" })).rejects.toThrow();
  });

  it("同一批匯入中，同一新人出現兩列：第一列建立新人，第二列可以找到第一列建立的人並配對", async () => {
    const batch = await startStaffDirectoryImport({
      fileName: "教職員工代號_測試3.pdf",
      importedBy: "測試",
      rows: [
        { rowNumber: 1, employeeCode: "E9101", name: "測試多列丁", departmentName: "課後照顧" },
        { rowNumber: 2, employeeCode: "E9102", name: "測試多列丁", departmentName: "攜手計畫" },
      ],
    });
    const [row1, row2] = batch.rows;

    // 匯入當下，這個人在系統裡完全不存在，兩列都還沒有候選人
    expect(await getCandidatePersonsForRow(row1.id)).toEqual([]);
    expect(await getCandidatePersonsForRow(row2.id)).toEqual([]);

    const resolvedRow1 = await resolveImportRow(row1.id, { action: "CREATE_NEW" }, "測試管理者");
    expect(resolvedRow1.matchStatus).toBe("CREATED_NEW");
    const newPersonId = resolvedRow1.matchedPersonId!;
    cleanupPersonIds.push(newPersonId);

    // 第一列確認建立新人後，第二列重新查詢候選人時應該立即看得到剛建立的人
    const candidatesForRow2 = await getCandidatePersonsForRow(row2.id);
    expect(candidatesForRow2.map((c) => c.id)).toEqual([newPersonId]);

    // 仍然不自動合併：由管理者明確選擇配對到剛建立的人
    const resolvedRow2 = await resolveImportRow(row2.id, { action: "MATCH_EXISTING", personId: newPersonId }, "測試管理者");
    expect(resolvedRow2.matchStatus).toBe("MATCHED_EXISTING");
    expect(resolvedRow2.matchedPersonId).toBe(newPersonId);

    // 最終只有一個 Person，底下有兩筆 PersonCode（課後照顧／攜手計畫）
    expect(await prisma.person.count({ where: { name: "測試多列丁" } })).toBe(1);
    const codes = await prisma.personCode.findMany({ where: { personId: newPersonId } });
    expect(codes.map((c) => c.categoryName).sort()).toEqual(["攜手計畫", "課後照顧"]);
  });

  it("listImportRows 會回傳每一列的即時候選清單", async () => {
    const wang = await prisma.person.findFirstOrThrow({ where: { name: "王○○" } });
    const batch = await startStaffDirectoryImport({
      fileName: "教職員工代號_測試4.pdf",
      importedBy: "測試",
      rows: [{ rowNumber: 1, employeeCode: "E9103", name: "王○○", departmentName: "攜手計畫" }],
    });
    const listed = await listImportRows(batch.id, "PENDING");
    expect(listed).toHaveLength(1);
    expect(listed[0].candidates.map((c) => c.id)).toEqual([wang.id]);

    // 清理：略過此測試列避免污染王○○的 PersonCode
    await resolveImportRow(listed[0].row.id, { action: "IGNORE" });
  });
});
