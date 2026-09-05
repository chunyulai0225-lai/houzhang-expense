// Phase 2：人員管理
//
// 這裡只實作「資料操作邏輯」（service 層），暫不做 UI。
// 依規格明確要求：
// - 人員不可刪除，離校只是狀態改變（deactivatePerson）。
// - 新增人員前必須先檢查同名既有人員，交由管理者確認，不自動視為同一人也不自動視為不同人。
// - 超鐘點／專案／代課計算不在本階段範圍內。

import type { EnrollmentStatus, Person } from "@prisma/client";
import { prisma } from "../prismaClient";

export type PersonListFilter = "ENROLLED" | "NOT_ENROLLED" | "ALL";

// 1. 人員列表：預設只顯示在校人員，可切換 全部／在校／不在校
export async function listPersons(filter: PersonListFilter = "ENROLLED") {
  const where = filter === "ALL" ? {} : { enrollmentStatus: filter as EnrollmentStatus };
  return prisma.person.findMany({
    where,
    orderBy: { name: "asc" },
    select: {
      id: true,
      name: true,
      payrollCode: true,
      enrollmentStatus: true,
      enrollDate: true,
      leaveDate: true,
      note: true,
    },
  });
}

// 2. 人員詳細資料：基本資料＋代碼／身分類別＋歷史學期職務＋歷史代課紀錄
// 注意：只顯示 Phase 1 已存在的資料，不在此階段計算超鐘點／專案／費用。
export async function getPersonDetail(personId: string) {
  const person = await prisma.person.findUnique({
    where: { id: personId },
    include: {
      personCodes: true,
      semesterRoles: { include: { semester: true } },
      substituteRecordsAsOriginal: true,
      substituteRecordsAsSubstitute: true,
    },
  });
  if (!person) {
    throw new Error(`Person not found: ${personId}`);
  }

  // 依學年度／學期新到舊排序，在 JS 端排序以避免依賴特定 Prisma 版本的巢狀關聯排序行為
  const semesterRoles = [...person.semesterRoles].sort((a, b) => {
    if (a.semester.schoolYear !== b.semester.schoolYear) {
      return b.semester.schoolYear - a.semester.schoolYear;
    }
    return b.semester.term - a.semester.term;
  });

  const substituteRecordsAsOriginal = [...person.substituteRecordsAsOriginal].sort(
    (a, b) => b.date.getTime() - a.date.getTime()
  );
  const substituteRecordsAsSubstitute = [...person.substituteRecordsAsSubstitute].sort(
    (a, b) => b.date.getTime() - a.date.getTime()
  );

  return { ...person, semesterRoles, substituteRecordsAsOriginal, substituteRecordsAsSubstitute };
}

// 5. 手動新增人員
export interface CreatePersonInput {
  name: string;
  payrollCode?: string;
  enrollmentStatus?: EnrollmentStatus;
  enrollDate?: Date;
  note?: string;
}

// 7. 重複人員保護：找出姓名相同的既有人員，交由管理者判斷
export async function findPossiblePersonDuplicates(name: string): Promise<Person[]> {
  const normalized = name.trim();
  return prisma.person.findMany({ where: { name: normalized } });
}

export type CreatePersonResult =
  | { status: "CREATED"; person: Person }
  | { status: "POSSIBLE_DUPLICATE"; candidates: Person[] };

// forceCreate: 管理者在看過既有同名人員後，仍確認要建立新人物時才傳 true。
export async function createPerson(
  input: CreatePersonInput,
  options: { forceCreate?: boolean } = {}
): Promise<CreatePersonResult> {
  const candidates = await findPossiblePersonDuplicates(input.name);
  if (candidates.length > 0 && !options.forceCreate) {
    return { status: "POSSIBLE_DUPLICATE", candidates };
  }

  const person = await prisma.person.create({
    data: {
      name: input.name.trim(),
      payrollCode: input.payrollCode,
      enrollmentStatus: input.enrollmentStatus ?? "ENROLLED",
      enrollDate: input.enrollDate,
      note: input.note,
    },
  });
  return { status: "CREATED", person };
}

// 6. 人員停用：不刪除資料，只將在校狀態改為不在校，並記錄離校日期與修改紀錄
export async function deactivatePerson(
  personId: string,
  params: { leaveDate: Date; changedBy?: string; reason?: string }
): Promise<Person> {
  const existing = await prisma.person.findUniqueOrThrow({ where: { id: personId } });

  const updated = await prisma.person.update({
    where: { id: personId },
    data: { enrollmentStatus: "NOT_ENROLLED", leaveDate: params.leaveDate },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "persons",
      recordId: personId,
      fieldName: "enrollmentStatus",
      oldValue: existing.enrollmentStatus,
      newValue: "NOT_ENROLLED",
      changedBy: params.changedBy,
      reason: params.reason,
    },
  });

  return updated;
}
