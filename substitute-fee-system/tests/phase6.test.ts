// Phase 6 驗收測試：學校上課日曆
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { prisma } from "../src/prismaClient";
import {
  addCalendarDay,
  generateSemesterCalendar,
  getMonthlySummary,
  getSemesterSummary,
  listCalendarDays,
  updateCalendarDay,
} from "../src/services/schoolCalendarService";

const cleanupSemesterIds: string[] = [];

function weekdayOf(dateStr: string): string {
  const idx = new Date(`${dateStr}T00:00:00.000Z`).getUTCDay();
  return ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"][idx];
}

async function makeTestSemester(schoolYear: number, startDate: string, endDate: string) {
  const sem = await prisma.semester.create({
    data: { schoolYear, term: 1, startDate: new Date(startDate), endDate: new Date(endDate) },
  });
  cleanupSemesterIds.push(sem.id);
  return sem;
}

beforeAll(async () => {
  await prisma.$connect();
});

afterAll(async () => {
  await prisma.changeLog.deleteMany({ where: { recordId: { in: cleanupSemesterIds } } });
  await prisma.schoolCalendarDay.deleteMany({ where: { semesterId: { in: cleanupSemesterIds } } });
  // 需要先刪掉指向這些學期的 changeLog（用 school_calendar_days recordId 是 day id，這裡另外清一次以防遺漏）
  await prisma.semester.deleteMany({ where: { id: { in: cleanupSemesterIds } } });
  await prisma.$disconnect();
});

describe("1/2. 批次產生學期日曆，週六日預設非上課日", () => {
  it("依學期起訖日產生每一天，週一~五預設上課、週六日預設非上課", async () => {
    // 2050/09/01(四) ~ 2050/09/14(三)，共14天，涵蓋兩個完整週末
    const sem = await makeTestSemester(501, "2050-09-01", "2050-09-14");
    const result = await generateSemesterCalendar(sem.id, "測試管理者");
    expect(result.createdCount).toBe(14);
    expect(result.skippedCount).toBe(0);

    const days = await listCalendarDays(sem.id);
    expect(days.length).toBe(14);
    for (const day of days) {
      const key = day.date.toISOString().slice(0, 10);
      expect(day.weekday).toBe(weekdayOf(key));
      const expectedDefault = day.weekday !== "SAT" && day.weekday !== "SUN";
      expect(day.isTeachingDay).toBe(expectedDefault);
    }
  });

  it("重複執行不會覆蓋既有資料、也不會產生重複列", async () => {
    const sem = await makeTestSemester(502, "2050-10-01", "2050-10-07");
    await generateSemesterCalendar(sem.id);
    // 管理者手動把其中一天改掉
    const days = await listCalendarDays(sem.id);
    await updateCalendarDay(days[0].id, { isTeachingDay: !days[0].isTeachingDay }, "測試管理者", "測試手動調整");

    const second = await generateSemesterCalendar(sem.id);
    expect(second.createdCount).toBe(0);
    expect(second.skippedCount).toBe(7);

    const afterSecondRun = await prisma.schoolCalendarDay.findUnique({ where: { id: days[0].id } });
    expect(afterSecondRun?.isTeachingDay).toBe(!days[0].isTeachingDay); // 手動調整沒有被蓋掉

    const allDays = await listCalendarDays(sem.id);
    expect(allDays.length).toBe(7); // 沒有重複
  });
});

describe("3/4. 上課日狀態可人工調整", () => {
  it("週六可以改成上課日，平日可以改成非上課日", async () => {
    const sem = await makeTestSemester(503, "2050-11-01", "2050-11-07"); // 含一個週六、週日
    await generateSemesterCalendar(sem.id);
    const days = await listCalendarDays(sem.id);
    const saturday = days.find((d) => d.weekday === "SAT")!;
    const weekday = days.find((d) => d.weekday === "MON" || d.weekday === "TUE")!;

    expect(saturday.isTeachingDay).toBe(false);
    const updatedSat = await updateCalendarDay(saturday.id, { isTeachingDay: true }, "測試管理者", "補課");
    expect(updatedSat.isTeachingDay).toBe(true);

    expect(weekday.isTeachingDay).toBe(true);
    const updatedWeekday = await updateCalendarDay(weekday.id, { isTeachingDay: false }, "測試管理者", "校外教學停課");
    expect(updatedWeekday.isTeachingDay).toBe(false);
  });
});

describe("5/6. 唯一限制與範圍檢查", () => {
  it("同一學期同一天不能重複建立", async () => {
    const sem = await makeTestSemester(504, "2050-12-01", "2050-12-07");
    await addCalendarDay({ semesterId: sem.id, date: new Date("2050-12-03") });
    await expect(addCalendarDay({ semesterId: sem.id, date: new Date("2050-12-03") })).rejects.toThrow();
  });

  it("日期超出學期範圍會被拒絕", async () => {
    const sem = await makeTestSemester(505, "2051-01-01", "2051-01-07");
    await expect(addCalendarDay({ semesterId: sem.id, date: new Date("2051-02-01") })).rejects.toThrow();
    await expect(addCalendarDay({ semesterId: sem.id, date: new Date("2050-12-31") })).rejects.toThrow();
  });
});

describe("7/8/10. 月統計／學期統計，且即時反映修改", () => {
  it("月統計正確反映週一~日各自上課日數，且修改後立即更新", async () => {
    const sem = await makeTestSemester(506, "2051-03-01", "2051-03-31"); // 一整個月
    await generateSemesterCalendar(sem.id);

    const before = await getMonthlySummary(sem.id, 2051, 3);
    const mondayCountBefore = before.byWeekday.find((w) => w.weekday === "MON")!.teachingDayCount;
    expect(mondayCountBefore).toBeGreaterThan(0);
    const satCountBefore = before.byWeekday.find((w) => w.weekday === "SAT")!.teachingDayCount;
    expect(satCountBefore).toBe(0);

    // 手動把其中一個週一改成非上課日
    const days = await listCalendarDays(sem.id, { year: 2051, month: 3 });
    const oneMonday = days.find((d) => d.weekday === "MON")!;
    await updateCalendarDay(oneMonday.id, { isTeachingDay: false }, "測試管理者", "臨時停課");

    const after = await getMonthlySummary(sem.id, 2051, 3);
    const mondayCountAfter = after.byWeekday.find((w) => w.weekday === "MON")!.teachingDayCount;
    expect(mondayCountAfter).toBe(mondayCountBefore - 1);
  });

  it("學期統計加總所有月份的上課日數", async () => {
    const sem = await makeTestSemester(507, "2051-04-01", "2051-04-14");
    await generateSemesterCalendar(sem.id);
    const summary = await getSemesterSummary(sem.id);
    const total = summary.byWeekday.reduce((acc, w) => acc + w.teachingDayCount, 0);
    expect(summary.totalTeachingDays).toBe(total);
    expect(summary.totalTeachingDays).toBeGreaterThan(0);
  });
});

describe("9. 特殊日期備註", () => {
  it("可以保存備註，例如畢業後特殊狀態", async () => {
    const sem = await makeTestSemester(508, "2051-06-01", "2051-06-30");
    await generateSemesterCalendar(sem.id);
    const days = await listCalendarDays(sem.id);
    const graduationDay = days.find((d) => d.date.toISOString().slice(0, 10) === "2051-06-12")!;
    const updated = await updateCalendarDay(
      graduationDay.id,
      { note: "六年級畢業，之後無減課" },
      "測試管理者",
      "畢業典禮"
    );
    expect(updated.note).toBe("六年級畢業，之後無減課");
  });
});

describe("11. 歷史學期資料不受目前學期修改影響", () => {
  it("修改某學期的日曆不會影響另一個學期", async () => {
    const semA = await makeTestSemester(509, "2051-08-01", "2051-08-07");
    const semB = await makeTestSemester(510, "2051-08-01", "2051-08-07"); // 刻意用相同日期範圍但不同學期
    await generateSemesterCalendar(semA.id);
    await generateSemesterCalendar(semB.id);

    const daysA = await listCalendarDays(semA.id);
    await updateCalendarDay(daysA[0].id, { isTeachingDay: false, note: "只改A學期" }, "測試管理者");

    const summaryB = await getSemesterSummary(semB.id);
    const daysB = await listCalendarDays(semB.id);
    expect(daysB.every((d) => d.note !== "只改A學期")).toBe(true);
    // B 學期完全沒被動過，應該還是預設值（平日皆上課）
    const expectedTeachingDaysB = daysB.filter((d) => d.weekday !== "SAT" && d.weekday !== "SUN").length;
    expect(summaryB.totalTeachingDays).toBe(expectedTeachingDaysB);
  });
});

describe("12. ChangeLog 記錄日期狀態修改", () => {
  it("批次產生與單日修改都會留下 ChangeLog", async () => {
    const sem = await makeTestSemester(511, "2051-09-01", "2051-09-07");
    await generateSemesterCalendar(sem.id, "測試管理者");

    const generateLog = await prisma.changeLog.findFirst({
      where: { tableName: "school_calendar_days", recordId: sem.id },
    });
    expect(generateLog?.reason).toContain("批次產生學期日曆");

    const days = await listCalendarDays(sem.id);
    await updateCalendarDay(days[0].id, { isTeachingDay: !days[0].isTeachingDay }, "測試管理者", "測試修改留痕");

    const updateLog = await prisma.changeLog.findFirst({
      where: { tableName: "school_calendar_days", recordId: days[0].id, fieldName: "isTeachingDay" },
    });
    expect(updateLog?.reason).toBe("測試修改留痕");
    expect(updateLog?.oldValue).toBe(String(days[0].isTeachingDay));
    expect(updateLog?.newValue).toBe(String(!days[0].isTeachingDay));
  });
});
