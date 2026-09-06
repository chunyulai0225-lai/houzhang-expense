// Phase 3 驗收測試：學期職務管理
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { prisma } from "../src/prismaClient";
import { createPerson } from "../src/services/personService";
import {
  assignSemesterRole,
  listRolesForPerson,
  listSemesterRoles,
  removeSemesterRole,
  updateSemesterRole,
} from "../src/services/semesterRoleService";

const cleanupPersonIds: string[] = [];
const cleanupRoleIds: string[] = [];

beforeAll(async () => {
  await prisma.$connect();
});

afterAll(async () => {
  await prisma.changeLog.deleteMany({
    where: { OR: [{ recordId: { in: cleanupRoleIds } }, { recordId: { in: cleanupPersonIds } }] },
  });
  await prisma.semesterRole.deleteMany({ where: { id: { in: cleanupRoleIds } } });
  await prisma.person.deleteMany({ where: { id: { in: cleanupPersonIds } } });
  await prisma.$disconnect();
});

describe("學期職務：新增／修改／刪除", () => {
  it("新增職務會建立 SemesterRole 並留下 ChangeLog", async () => {
    const created = await createPerson({ name: "測試職務教師戊" });
    if (created.status !== "CREATED") throw new Error("unreachable");
    cleanupPersonIds.push(created.person.id);

    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });

    const role = await assignSemesterRole(
      { personId: created.person.id, semesterId: sem.id, roleType: "HOMEROOM_TEACHER", roleDetail: "三年戊班導師" },
      "測試管理者"
    );
    cleanupRoleIds.push(role.id);

    expect(role.roleType).toBe("HOMEROOM_TEACHER");
    const logs = await prisma.changeLog.findMany({ where: { tableName: "semester_roles", recordId: role.id } });
    expect(logs.length).toBe(1);
    expect(logs[0].newValue).toBe("HOMEROOM_TEACHER:三年戊班導師");
    expect(logs[0].reason).toBe("新增學期職務");
  });

  it("同一人同一學期可以同時擔任多個職務", async () => {
    const created = await createPerson({ name: "測試職務教師己" });
    if (created.status !== "CREATED") throw new Error("unreachable");
    cleanupPersonIds.push(created.person.id);

    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });

    const role1 = await assignSemesterRole({
      personId: created.person.id,
      semesterId: sem.id,
      roleType: "HOMEROOM_TEACHER",
      roleDetail: "四年甲班導師",
    });
    const role2 = await assignSemesterRole({
      personId: created.person.id,
      semesterId: sem.id,
      roleType: "SECTION_CHIEF",
      roleDetail: "教學組長",
    });
    cleanupRoleIds.push(role1.id, role2.id);

    const roles = await listRolesForPerson(created.person.id);
    expect(roles.length).toBe(2);
    expect(roles.map((r) => r.roleType).sort()).toEqual(["HOMEROOM_TEACHER", "SECTION_CHIEF"]);
  });

  it("修改職務會逐欄位記錄修改前後", async () => {
    const created = await createPerson({ name: "測試職務教師庚" });
    if (created.status !== "CREATED") throw new Error("unreachable");
    cleanupPersonIds.push(created.person.id);

    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const role = await assignSemesterRole({
      personId: created.person.id,
      semesterId: sem.id,
      roleType: "SUBJECT_TEACHER",
      roleDetail: "體育科任",
    });
    cleanupRoleIds.push(role.id);

    const updated = await updateSemesterRole(
      role.id,
      { roleDetail: "音樂科任" },
      "測試管理者",
      "更正詳細職務"
    );
    expect(updated.roleDetail).toBe("音樂科任");

    const logs = await prisma.changeLog.findMany({
      where: { tableName: "semester_roles", recordId: role.id, fieldName: "roleDetail" },
    });
    expect(logs.length).toBe(1);
    expect(logs[0].oldValue).toBe("體育科任");
    expect(logs[0].newValue).toBe("音樂科任");
    expect(logs[0].reason).toBe("更正詳細職務");
  });

  it("刪除職務會移除資料列並在 ChangeLog 留下刪除前的內容", async () => {
    const created = await createPerson({ name: "測試職務教師辛" });
    if (created.status !== "CREATED") throw new Error("unreachable");
    cleanupPersonIds.push(created.person.id);

    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const role = await assignSemesterRole({
      personId: created.person.id,
      semesterId: sem.id,
      roleType: "ADMIN",
      roleDetail: "誤植職務",
    });

    await removeSemesterRole(role.id, "測試管理者", "指派錯誤，予以刪除");

    const stillExists = await prisma.semesterRole.findUnique({ where: { id: role.id } });
    expect(stillExists).toBeNull();

    // 這裡會有兩筆：新增時一筆、刪除時一筆，都保留可追查
    const logs = await prisma.changeLog.findMany({
      where: { tableName: "semester_roles", recordId: role.id },
      orderBy: { createdAt: "asc" },
    });
    expect(logs.length).toBe(2);
    const deleteLog = logs[1];
    expect(deleteLog.oldValue).toBe("ADMIN:誤植職務");
    expect(deleteLog.newValue).toBeNull();
    expect(deleteLog.reason).toBe("指派錯誤，予以刪除");
  });
});

describe("學期職務查詢", () => {
  it("listSemesterRoles 可以列出某學期所有職務（含 Phase 1 種子資料）", async () => {
    const sem = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 1 } });
    const roles = await listSemesterRoles(sem.id);
    expect(roles.some((r) => r.person.name === "陳心啓" && r.roleDetail === "自然科任")).toBe(true);
    expect(roles.some((r) => r.person.name === "王○○" && r.roleDetail === "六年甲班導師")).toBe(true);
  });

  it("listRolesForPerson 依學年度／學期新到舊排序", async () => {
    const chen = await prisma.person.findFirstOrThrow({ where: { name: "陳心啓" } });
    const sem115_2 = await prisma.semester.findFirstOrThrow({ where: { schoolYear: 115, term: 2 } });

    const extraRole = await assignSemesterRole({
      personId: chen.id,
      semesterId: sem115_2.id,
      roleType: "SUBJECT_TEACHER",
      roleDetail: "自然科任（115-2）",
    });
    cleanupRoleIds.push(extraRole.id);

    const roles = await listRolesForPerson(chen.id);
    expect(roles[0].semester.term).toBe(2);
    expect(roles[0].semester.schoolYear).toBe(115);
  });
});
