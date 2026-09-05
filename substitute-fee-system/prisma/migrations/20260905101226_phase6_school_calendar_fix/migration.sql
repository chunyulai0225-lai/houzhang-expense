-- AlterEnum
-- This migration adds more than one value to an enum.
-- With PostgreSQL versions 11 and earlier, this is not possible
-- in a single migration. This can be worked around by creating
-- multiple migrations, each migration adding only one value to
-- the enum.


ALTER TYPE "Weekday" ADD VALUE 'SAT';
ALTER TYPE "Weekday" ADD VALUE 'SUN';

-- DropForeignKey
ALTER TABLE "SchoolCalendarDay" DROP CONSTRAINT "SchoolCalendarDay_semesterId_fkey";

-- DropIndex
DROP INDEX "SchoolCalendarDay_date_key";

-- AlterTable
ALTER TABLE "SchoolCalendarDay" ALTER COLUMN "semesterId" SET NOT NULL;

-- CreateIndex
CREATE UNIQUE INDEX "SchoolCalendarDay_semesterId_date_key" ON "SchoolCalendarDay"("semesterId", "date");

-- AddForeignKey
ALTER TABLE "SchoolCalendarDay" ADD CONSTRAINT "SchoolCalendarDay_semesterId_fkey" FOREIGN KEY ("semesterId") REFERENCES "Semester"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

