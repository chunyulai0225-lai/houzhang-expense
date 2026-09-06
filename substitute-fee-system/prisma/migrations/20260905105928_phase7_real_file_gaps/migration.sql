-- AlterTable
ALTER TABLE "MonthlyImport" ADD COLUMN     "sourceSheetName" TEXT,
ADD COLUMN     "sourceStaffType" "StaffType" NOT NULL DEFAULT 'UNKNOWN';

-- AlterTable
ALTER TABLE "SubstituteRecordRaw" ADD COLUMN     "dailyOrHalfDayWageText" TEXT,
ADD COLUMN     "periodCountText" TEXT,
ADD COLUMN     "sheetName" TEXT,
ADD COLUMN     "substitutePeriodFeeText" TEXT;
