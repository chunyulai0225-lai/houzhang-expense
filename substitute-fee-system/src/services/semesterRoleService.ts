// Phase 3：學期職務管理
//
// 職務資料模型沿用 Phase 1 的 SemesterRole（已綁定 personId + semesterId），
// 這裡只補上管理操作（新增／修改／刪除／查詢）與對應的 ChangeLog 稽核紀錄，
// 不新增任何尚未確認的業務規則（例如職務是否互斥、是否有數量上限等）。
// 一人一學期可以同時有多筆職務（例如導師兼組長），本階段不做限制。

import type { RoleType, SemesterRole } from "@prisma/client";
import { prisma } from "../prismaClient";

export interface AssignSemesterRoleInput {
  personId: string;
  semesterId: string;
  roleType: RoleType;
  roleDetail?: string;
  note?: string;
}

// 1. 新增學期職務
export async function assignSemesterRole(
  input: AssignSemesterRoleInput,
  changedBy?: string
): Promise<SemesterRole> {
  const role = await prisma.semesterRole.create({ data: input });

  await prisma.changeLog.create({
    data: {
      tableName: "semester_roles",
      recordId: role.id,
      newValue: describeRole(role),
      changedBy,
      reason: "新增學期職務",
    },
  });

  return role;
}

export interface UpdateSemesterRoleInput {
  roleType?: RoleType;
  roleDetail?: string | null;
  note?: string | null;
}

// 2. 修改學期職務（例如更正詳細職務名稱），逐欄位記錄異動
export async function updateSemesterRole(
  roleId: string,
  changes: UpdateSemesterRoleInput,
  changedBy?: string,
  reason?: string
): Promise<SemesterRole> {
  const existing = await prisma.semesterRole.findUniqueOrThrow({ where: { id: roleId } });
  const updated = await prisma.semesterRole.update({ where: { id: roleId }, data: changes });

  const fields = Object.keys(changes) as (keyof UpdateSemesterRoleInput)[];
  for (const field of fields) {
    const oldValue = existing[field];
    const newValue = updated[field];
    if (oldValue !== newValue) {
      await prisma.changeLog.create({
        data: {
          tableName: "semester_roles",
          recordId: roleId,
          fieldName: field,
          oldValue: oldValue == null ? null : String(oldValue),
          newValue: newValue == null ? null : String(newValue),
          changedBy,
          reason,
        },
      });
    }
  }

  return updated;
}

// 3. 刪除學期職務（例如職務指派錯誤）；異動仍留在 ChangeLog 可追查
export async function removeSemesterRole(roleId: string, changedBy?: string, reason?: string): Promise<void> {
  const existing = await prisma.semesterRole.findUniqueOrThrow({ where: { id: roleId } });
  await prisma.semesterRole.delete({ where: { id: roleId } });

  await prisma.changeLog.create({
    data: {
      tableName: "semester_roles",
      recordId: roleId,
      oldValue: describeRole(existing),
      changedBy,
      reason: reason ?? "刪除學期職務",
    },
  });
}

// 4. 查詢：某學期的所有職務指派（供學期職務列表頁使用）
export async function listSemesterRoles(semesterId: string) {
  return prisma.semesterRole.findMany({
    where: { semesterId },
    include: { person: true },
    orderBy: [{ roleType: "asc" }, { person: { name: "asc" } }],
  });
}

// 5. 查詢：某人跨學年度／學期的所有職務歷史（新到舊）
export async function listRolesForPerson(personId: string) {
  const roles = await prisma.semesterRole.findMany({
    where: { personId },
    include: { semester: true },
  });
  return roles.sort((a, b) => {
    if (a.semester.schoolYear !== b.semester.schoolYear) {
      return b.semester.schoolYear - a.semester.schoolYear;
    }
    return b.semester.term - a.semester.term;
  });
}

function describeRole(role: Pick<SemesterRole, "roleType" | "roleDetail">): string {
  return role.roleDetail ? `${role.roleType}:${role.roleDetail}` : role.roleType;
}
