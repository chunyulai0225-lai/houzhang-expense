// Phase 5：專案基本資料管理
//
// 只處理專案「名稱／學期／是否啟用／備註」，不處理費用類型或單價——
// 專案的計費規則還沒確認（各專案的扣除公式、是否使用不同費率都待確認），
// 等那些確認後再另外設計，不會直接在 Project 上放一個容易被覆蓋的金額欄位。

import type { Project } from "@prisma/client";
import { prisma } from "../prismaClient";

export interface CreateProjectInput {
  semesterId: string;
  name: string;
  isActive?: boolean;
  note?: string;
}

// 1. 新增專案
export async function createProject(input: CreateProjectInput, changedBy?: string): Promise<Project> {
  const project = await prisma.project.create({
    data: {
      semesterId: input.semesterId,
      name: input.name.trim(),
      isActive: input.isActive ?? true,
      note: input.note,
    },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "projects",
      recordId: project.id,
      newValue: project.name,
      changedBy,
      reason: "新增專案",
    },
  });

  return project;
}

export interface UpdateProjectInput {
  name?: string;
  note?: string | null;
}

// 2. 修改專案基本資料（名稱、備註）
export async function updateProject(
  projectId: string,
  changes: UpdateProjectInput,
  changedBy?: string,
  reason?: string
): Promise<Project> {
  const existing = await prisma.project.findUniqueOrThrow({ where: { id: projectId } });
  const updated = await prisma.project.update({ where: { id: projectId }, data: changes });

  const fields = Object.keys(changes) as (keyof UpdateProjectInput)[];
  for (const field of fields) {
    const oldValue = existing[field];
    const newValue = updated[field];
    if (oldValue !== newValue) {
      await prisma.changeLog.create({
        data: {
          tableName: "projects",
          recordId: projectId,
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

// 3. 停用／啟用專案（優先於刪除；停用後歷史規則、歷史代課紀錄仍可查詢）
export async function setProjectActive(
  projectId: string,
  isActive: boolean,
  changedBy?: string,
  reason?: string
): Promise<Project> {
  const existing = await prisma.project.findUniqueOrThrow({ where: { id: projectId } });
  const updated = await prisma.project.update({ where: { id: projectId }, data: { isActive } });

  if (existing.isActive !== updated.isActive) {
    await prisma.changeLog.create({
      data: {
        tableName: "projects",
        recordId: projectId,
        fieldName: "isActive",
        oldValue: String(existing.isActive),
        newValue: String(updated.isActive),
        changedBy,
        reason: reason ?? (isActive ? "啟用專案" : "停用專案"),
      },
    });
  }

  return updated;
}

// 4. 查詢：某學期的專案清單，預設含停用的（供管理介面篩選）
export async function listProjects(semesterId: string, filter?: { isActive?: boolean }) {
  return prisma.project.findMany({
    where: { semesterId, ...(filter?.isActive !== undefined ? { isActive: filter.isActive } : {}) },
    orderBy: { name: "asc" },
  });
}
