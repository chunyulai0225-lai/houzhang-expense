-- CreateEnum
CREATE TYPE "SubstituteRecordEntryType" AS ENUM ('EXCEL_IMPORT', 'MANUAL_SELF_FUNDED');

-- DropForeignKey
ALTER TABLE "SubstituteRecord" DROP CONSTRAINT "SubstituteRecord_rawRecordId_fkey";

-- AlterTable
ALTER TABLE "SubstituteRecord" ADD COLUMN     "createdBy" TEXT,
ADD COLUMN     "entryType" "SubstituteRecordEntryType" NOT NULL DEFAULT 'EXCEL_IMPORT',
ADD COLUMN     "updatedBy" TEXT,
ALTER COLUMN "rawRecordId" DROP NOT NULL;

-- CreateTable
CREATE TABLE "MonthlyIssueAcknowledgement" (
    "id" TEXT NOT NULL,
    "semesterId" TEXT NOT NULL,
    "year" INTEGER NOT NULL,
    "month" INTEGER NOT NULL,
    "targetTable" TEXT NOT NULL,
    "targetId" TEXT NOT NULL,
    "reason" TEXT NOT NULL,
    "acknowledgedBy" TEXT NOT NULL,
    "acknowledgedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "MonthlyIssueAcknowledgement_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "MonthlyIssueAcknowledgement_semesterId_year_month_idx" ON "MonthlyIssueAcknowledgement"("semesterId", "year", "month");

-- CreateIndex
CREATE UNIQUE INDEX "MonthlyIssueAcknowledgement_targetTable_targetId_key" ON "MonthlyIssueAcknowledgement"("targetTable", "targetId");

-- AddForeignKey
ALTER TABLE "SubstituteRecord" ADD CONSTRAINT "SubstituteRecord_rawRecordId_fkey" FOREIGN KEY ("rawRecordId") REFERENCES "SubstituteRecordRaw"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "MonthlyIssueAcknowledgement" ADD CONSTRAINT "MonthlyIssueAcknowledgement_semesterId_fkey" FOREIGN KEY ("semesterId") REFERENCES "Semester"("id") ON DELETE RESTRICT ON UPDATE CASCADE;
