// Phase 6：學校上課日曆
//
// 核心原則：以「日」為最小單位保存，月份／星期上班日數一律由 SchoolCalendarDay
// 即時聚合查詢得出，不另外存一份「9月星期一4天」之類的獨立統計資料，
// 避免兩份資料不同步。
//
// 不在這裡做的事：超鐘點/專案費用計算、代課Excel匯入/分類、代導師費/日薪/半日薪、
// BD/非BD付款判斷——這裡只負責「某一天到底是不是上課日」這件事本身。

import type { SchoolCalendarDay, Weekday } from "@prisma/client";
import { prisma } from "../prismaClient";

const WEEKDAY_BY_JS_INDEX: Weekday[] = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"];
const ALL_WEEKDAYS: Weekday[] = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"];

function computeWeekday(date: Date): Weekday {
  return WEEKDAY_BY_JS_INDEX[date.getUTCDay()];
}

// 預設值，不是最終規則：週一～週五預設上課、週六日預設非上課，
// 兩者都只是「產生資料當下的預設」，管理者可以針對任何一天個別覆蓋，
// 之後所有計算一律以 SchoolCalendarDay 實際存的 isTeachingDay 為準。
function defaultIsTeachingDay(weekday: Weekday): boolean {
  return weekday !== "SAT" && weekday !== "SUN";
}

function toUtcMidnight(date: Date): Date {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
}

function dateKey(date: Date): string {
  return date.toISOString().slice(0, 10);
}

function assertWithinSemester(date: Date, semester: { startDate: Date; endDate: Date }) {
  if (date < semester.startDate || date > semester.endDate) {
    throw new Error(
      `日期 ${dateKey(date)} 超出學期範圍（${dateKey(semester.startDate)} ~ ${dateKey(semester.endDate)}）`
    );
  }
}

export interface GenerateCalendarResult {
  createdCount: number;
  skippedCount: number;
}

// 1. 批次依學期起訖日產生每日資料。已存在的日期不會被覆蓋或重新產生，
//    所以重複執行這個函式是安全的——不會蓋掉管理者已經手動調整過的資料。
export async function generateSemesterCalendar(
  semesterId: string,
  changedBy?: string
): Promise<GenerateCalendarResult> {
  const semester = await prisma.semester.findUniqueOrThrow({ where: { id: semesterId } });
  const existing = await prisma.schoolCalendarDay.findMany({ where: { semesterId }, select: { date: true } });
  const existingDates = new Set(existing.map((d) => dateKey(d.date)));

  const daysToCreate: { semesterId: string; date: Date; weekday: Weekday; isTeachingDay: boolean }[] = [];
  const cursor = toUtcMidnight(semester.startDate);
  const end = toUtcMidnight(semester.endDate);
  while (cursor.getTime() <= end.getTime()) {
    if (!existingDates.has(dateKey(cursor))) {
      const weekday = computeWeekday(cursor);
      daysToCreate.push({
        semesterId,
        date: new Date(cursor),
        weekday,
        isTeachingDay: defaultIsTeachingDay(weekday),
      });
    }
    cursor.setUTCDate(cursor.getUTCDate() + 1);
  }

  if (daysToCreate.length > 0) {
    await prisma.schoolCalendarDay.createMany({ data: daysToCreate });
  }

  await prisma.changeLog.create({
    data: {
      tableName: "school_calendar_days",
      recordId: semesterId,
      newValue: `新增 ${daysToCreate.length} 天`,
      changedBy,
      reason: `批次產生學期日曆（略過已存在 ${existingDates.size} 天）`,
    },
  });

  return { createdCount: daysToCreate.length, skippedCount: existingDates.size };
}

export interface AddCalendarDayInput {
  semesterId: string;
  date: Date;
  isTeachingDay?: boolean;
  note?: string;
}

// 補充新增單一天（例如批次產生後才發現學期範圍調整、需要補一天）。
// 日期必須在學期範圍內，且同學期同一天不能重複。
export async function addCalendarDay(input: AddCalendarDayInput, changedBy?: string): Promise<SchoolCalendarDay> {
  const semester = await prisma.semester.findUniqueOrThrow({ where: { id: input.semesterId } });
  const date = toUtcMidnight(input.date);
  assertWithinSemester(date, semester);

  const existing = await prisma.schoolCalendarDay.findUnique({
    where: { semesterId_date: { semesterId: input.semesterId, date } },
  });
  if (existing) {
    throw new Error(`此學期的 ${dateKey(date)} 已經存在，不可重複建立`);
  }

  const weekday = computeWeekday(date);
  const created = await prisma.schoolCalendarDay.create({
    data: {
      semesterId: input.semesterId,
      date,
      weekday,
      isTeachingDay: input.isTeachingDay ?? defaultIsTeachingDay(weekday),
      note: input.note,
    },
  });

  await prisma.changeLog.create({
    data: {
      tableName: "school_calendar_days",
      recordId: created.id,
      newValue: describeDay(created),
      changedBy,
      reason: "新增單日行事曆",
    },
  });

  return created;
}

export interface UpdateCalendarDayInput {
  isTeachingDay?: boolean;
  note?: string | null;
}

// 4. 修改：是否上課日／備註。逐欄位記錄異動，供稽核追查。
export async function updateCalendarDay(
  dayId: string,
  changes: UpdateCalendarDayInput,
  changedBy?: string,
  reason?: string
): Promise<SchoolCalendarDay> {
  const existing = await prisma.schoolCalendarDay.findUniqueOrThrow({ where: { id: dayId } });
  const updated = await prisma.schoolCalendarDay.update({ where: { id: dayId }, data: changes });

  const fields = Object.keys(changes) as (keyof UpdateCalendarDayInput)[];
  for (const field of fields) {
    const oldValue = existing[field];
    const newValue = updated[field];
    if (String(oldValue) !== String(newValue)) {
      await prisma.changeLog.create({
        data: {
          tableName: "school_calendar_days",
          recordId: dayId,
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

export interface ListCalendarDaysFilter {
  year?: number;
  month?: number; // 1-12
}

// 查詢：某學期的日曆資料，可選擇只看某一年月
export async function listCalendarDays(
  semesterId: string,
  filter: ListCalendarDaysFilter = {}
): Promise<SchoolCalendarDay[]> {
  let dateFilter: { gte: Date; lte: Date } | undefined;
  if (filter.year && filter.month) {
    const start = new Date(Date.UTC(filter.year, filter.month - 1, 1));
    const end = new Date(Date.UTC(filter.year, filter.month, 0)); // 該月最後一天
    dateFilter = { gte: start, lte: end };
  }

  return prisma.schoolCalendarDay.findMany({
    where: { semesterId, ...(dateFilter ? { date: dateFilter } : {}) },
    orderBy: { date: "asc" },
  });
}

export interface WeekdaySummary {
  weekday: Weekday;
  teachingDayCount: number;
}

export interface CalendarSummary {
  byWeekday: WeekdaySummary[];
  totalTeachingDays: number;
}

function summarize(days: Pick<SchoolCalendarDay, "weekday" | "isTeachingDay">[]): CalendarSummary {
  const counts = new Map<Weekday, number>();
  let total = 0;
  for (const day of days) {
    if (day.isTeachingDay) {
      counts.set(day.weekday, (counts.get(day.weekday) ?? 0) + 1);
      total += 1;
    }
  }
  const byWeekday = ALL_WEEKDAYS.map((weekday) => ({ weekday, teachingDayCount: counts.get(weekday) ?? 0 }));
  return { byWeekday, totalTeachingDays: total };
}

// 9. 月統計：星期一~日各自的上課日數，全部由當月的 SchoolCalendarDay 即時聚合
export async function getMonthlySummary(semesterId: string, year: number, month: number): Promise<CalendarSummary> {
  const days = await listCalendarDays(semesterId, { year, month });
  return summarize(days);
}

// 10. 學期統計：整學期的上課日數
export async function getSemesterSummary(semesterId: string): Promise<CalendarSummary> {
  const days = await listCalendarDays(semesterId);
  return summarize(days);
}

function describeDay(day: Pick<SchoolCalendarDay, "date" | "weekday" | "isTeachingDay">): string {
  return `${dateKey(day.date)} ${day.weekday} ${day.isTeachingDay ? "上課" : "非上課"}`;
}
