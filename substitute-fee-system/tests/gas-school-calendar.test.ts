// 學校上課日曆「產生」功能測試——載入真正的 gas/*.gs 原始碼執行
// （見 tests/helpers/gasHarness.ts）。
//
// 重點：重現並驗證「按產生卻顯示已產生 0 天」這個 bug 的根本原因——目前真實的
// 115 學年度第1學期資料很可能是使用者直接在 Google Sheets 手動輸入日期，Sheets
// 會把這種儲存格內容自動辨識成「日期」型別；就算後來對整欄執行過
// setNumberFormat("@") 改成純文字格式，Sheets 也不會回頭把已經存在的儲存格內容
// 轉換成字串——用 Apps Script 的 getValues() 讀出來仍然是原生 Date 物件，不是
// "YYYY-MM-DD" 字串。原本的 api_generateSemesterCalendar() 直接把它跟字串
// 相加（`semester.startDate + "T00:00:00Z"`），Date 物件相加字串會變成一大串不是
// 日期格式的文字，new Date(...) 出來是 Invalid Date，迴圈結束條件永遠是 false，
// 一次都不會執行，所以永遠回報「已產生 0 天」。
//
// 修法：toDateOnly() 現在會先判斷是不是原生 Date 物件（或極少數情況的日期序列
// 數字），用 Utilities.formatDate() 搭配 Session.getScriptTimeZone() 正規化成
// "YYYY-MM-DD" 字串，api_generateSemesterCalendar() 一律先正規化再做日期運算。
import { describe, expect, it, beforeEach } from "vitest";
import { createGasSandbox, seedRealSemester115_1, seedSemesterWithNativeDateCells } from "./helpers/gasHarness";

function countDaysInclusive(startDate: string, endDate: string): number {
  const start = new Date(startDate + "T00:00:00Z").getTime();
  const end = new Date(endDate + "T00:00:00Z").getTime();
  return Math.round((end - start) / 86400000) + 1;
}

describe("學校上課日曆：產生功能（GAS 端，載入真正的 gas/*.gs 原始碼執行）", () => {
  describe("正常情況：Semesters 的日期已經是乾淨的 \"YYYY-MM-DD\" 字串", () => {
    let sandbox: any;
    let semester: any;

    beforeEach(() => {
      sandbox = createGasSandbox();
      semester = seedRealSemester115_1(sandbox);
    });

    it("115-1 產生後不應該是 0 天，天數要等於 startDate ~ endDate 的實際天數", () => {
      const expectedDays = countDaysInclusive(semester.startDate, semester.endDate); // 2026-08-31 ~ 2027-01-20
      const result = sandbox.api_generateSemesterCalendar({ semesterId: semester.id, changedBy: "測試" });
      expect(result.createdCount).toBe(expectedDays);
      expect(result.createdCount).toBeGreaterThan(0);
      expect(result.skippedCount).toBe(0);

      const days = sandbox.api_listCalendarDays({ semesterId: semester.id });
      expect(days).toHaveLength(expectedDays);
    });

    it("週六、週日預設 isTeachingDay=false，週一~週五預設 isTeachingDay=true", () => {
      sandbox.api_generateSemesterCalendar({ semesterId: semester.id, changedBy: "測試" });
      const days = sandbox.api_listCalendarDays({ semesterId: semester.id });
      const weekend = days.filter((d: any) => d.weekday === "SAT" || d.weekday === "SUN");
      const weekday = days.filter((d: any) => ["MON", "TUE", "WED", "THU", "FRI"].includes(d.weekday));
      expect(weekend.length).toBeGreaterThan(0);
      expect(weekday.length).toBeGreaterThan(0);
      expect(weekend.every((d: any) => d.isTeachingDay === false)).toBe(true);
      expect(weekday.every((d: any) => d.isTeachingDay === true)).toBe(true);
    });

    it("再次產生不會重複資料，第二次新增數要是 0、略過數要等於第一次建立的總天數", () => {
      const first = sandbox.api_generateSemesterCalendar({ semesterId: semester.id, changedBy: "測試" });
      const second = sandbox.api_generateSemesterCalendar({ semesterId: semester.id, changedBy: "測試" });
      expect(second.createdCount).toBe(0);
      expect(second.skippedCount).toBe(first.createdCount);
      const days = sandbox.api_listCalendarDays({ semesterId: semester.id });
      expect(days).toHaveLength(first.createdCount); // 沒有變多，也沒有變少
    });

    it("已經人工修改過的日期不會被產生功能覆蓋", () => {
      sandbox.api_generateSemesterCalendar({ semesterId: semester.id, changedBy: "測試" });
      const days = sandbox.api_listCalendarDays({ semesterId: semester.id });
      // 找一天平日，人工改成「不是上課日」（例如校慶調整），模擬管理者手動調整。
      const aWeekday = days.find((d: any) => d.isTeachingDay === true);
      sandbox.api_updateCalendarDay({ id: aWeekday.id, isTeachingDay: false, note: "校慶補假（人工調整）", changedBy: "測試" });

      // 重新產生一次，這一天不應該被改回 true，備註也不應該被清空。
      sandbox.api_generateSemesterCalendar({ semesterId: semester.id, changedBy: "測試" });
      const afterRegenerate = sandbox.api_listCalendarDays({ semesterId: semester.id }).find((d: any) => d.id === aWeekday.id);
      expect(afterRegenerate.isTeachingDay).toBe(false);
      expect(afterRegenerate.note).toBe("校慶補假（人工調整）");
    });

    it("寫入 ChangeLog", () => {
      const before = sandbox.readRows("ChangeLog").length;
      sandbox.api_generateSemesterCalendar({ semesterId: semester.id, changedBy: "測試" });
      const after = sandbox.readRows("ChangeLog").length;
      expect(after).toBeGreaterThan(before);
    });

    it("不會自作主張加入國定假日、補假、校慶——平日一律預設上課日，只有週六日預設非上課日", () => {
      sandbox.api_generateSemesterCalendar({ semesterId: semester.id, changedBy: "測試" });
      const days = sandbox.api_listCalendarDays({ semesterId: semester.id });
      // 115-1 學期（2026-08-31~2027-01-20）涵蓋國慶日(10/10)、元旦(2027-01-01)等真實假日，
      // 系統不應該自己判斷這些是假日，一律照weekday預設值，等管理者自己手動調整。
      const nationalDay = days.find((d: any) => d.date === "2026-10-10");
      const newYearsDay = days.find((d: any) => d.date === "2027-01-01");
      expect(nationalDay).toBeTruthy();
      expect(newYearsDay).toBeTruthy();
      // 2026-10-10 是星期六，2027-01-01 是星期五——用 weekday 驗證系統只看星期幾，
      // 不是用日期本身去查什麼假日清單。
      expect(nationalDay.weekday).toBe("SAT");
      expect(nationalDay.isTeachingDay).toBe(false); // 因為是週六，不是因為系統認得國慶日
      expect(newYearsDay.weekday).toBe("FRI");
      expect(newYearsDay.isTeachingDay).toBe(true); // 因為是週五，系統不會自己判斷這是元旦假期
    });
  });

  describe("重現真實 bug：Semesters 的日期是原生 Date 物件（使用者直接在 Sheets 手動輸入）", () => {
    it("修好之前會是 0 天（先證明真的重現了這個 bug，而不是憑空修一個不存在的問題）", () => {
      const sandbox = createGasSandbox();
      const semester = seedSemesterWithNativeDateCells(sandbox);
      // 直接呼叫沒有正規化過的舊寫法，證明「不正規化就會是 0 天」這個因果關係。
      const cursor = new Date((semester.startDate as any) + "T00:00:00Z");
      const end = new Date((semester.endDate as any) + "T00:00:00Z");
      expect(Number.isNaN(cursor.getTime())).toBe(true);
      expect(Number.isNaN(end.getTime())).toBe(true);
    });

    it("修好之後：即使 Semesters 的日期是原生 Date 物件，產生功能仍然能正確算出實際天數", () => {
      const sandbox = createGasSandbox();
      const semester = seedSemesterWithNativeDateCells(sandbox);
      const expectedDays = countDaysInclusive("2026-08-31", "2027-01-20");

      const result = sandbox.api_generateSemesterCalendar({ semesterId: semester.id, changedBy: "測試" });

      expect(result.createdCount).toBe(expectedDays);
      expect(result.createdCount).toBeGreaterThan(0);
      expect(result.startDate).toBe("2026-08-31");
      expect(result.endDate).toBe("2027-01-20");
    });
  });

  it("找不到學期要報錯，不會靜默回傳 0 天", () => {
    const sandbox = createGasSandbox();
    expect(() => sandbox.api_generateSemesterCalendar({ semesterId: "not-a-real-id", changedBy: "測試" })).toThrow("找不到學期");
  });
});
