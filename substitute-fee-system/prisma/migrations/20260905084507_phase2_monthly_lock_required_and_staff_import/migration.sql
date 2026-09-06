/*
  Warnings:

  - Made the column `semesterId` on table `MonthlyLock` required. This step will fail if there are existing NULL values in that column.

*/
-- CreateEnum
CREATE TYPE "ImportRowMatchStatus" AS ENUM ('PENDING', 'MATCHED_EXISTING', 'CREATED_NEW', 'IGNORED');

-- DropForeignKey
ALTER TABLE "MonthlyLock" DROP CONSTRAINT "MonthlyLock_semesterId_fkey";

-- AlterTable
ALTER TABLE "MonthlyLock" ALTER COLUMN "semesterId" SET NOT NULL;

-- CreateTable
CREATE TABLE "StaffDirectoryImport" (
    "id" TEXT NOT NULL,
    "fileName" TEXT NOT NULL,
    "importedBy" TEXT,
    "importedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "totalCount" INTEGER NOT NULL DEFAULT 0,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "StaffDirectoryImport_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "StaffDirectoryImportRow" (
    "id" TEXT NOT NULL,
    "importId" TEXT NOT NULL,
    "rowNumber" INTEGER,
    "employeeCode" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "departmentName" TEXT,
    "schoolYear" INTEGER,
    "suggestedPersonId" TEXT,
    "matchStatus" "ImportRowMatchStatus" NOT NULL DEFAULT 'PENDING',
    "matchedPersonId" TEXT,
    "createdPersonCodeId" TEXT,
    "resolvedAt" TIMESTAMP(3),
    "resolvedBy" TEXT,
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "StaffDirectoryImportRow_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "StaffDirectoryImportRow_importId_idx" ON "StaffDirectoryImportRow"("importId");

-- CreateIndex
CREATE INDEX "StaffDirectoryImportRow_name_idx" ON "StaffDirectoryImportRow"("name");

-- CreateIndex
CREATE INDEX "StaffDirectoryImportRow_matchStatus_idx" ON "StaffDirectoryImportRow"("matchStatus");

-- AddForeignKey
ALTER TABLE "MonthlyLock" ADD CONSTRAINT "MonthlyLock_semesterId_fkey" FOREIGN KEY ("semesterId") REFERENCES "Semester"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "StaffDirectoryImportRow" ADD CONSTRAINT "StaffDirectoryImportRow_importId_fkey" FOREIGN KEY ("importId") REFERENCES "StaffDirectoryImport"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "StaffDirectoryImportRow" ADD CONSTRAINT "StaffDirectoryImportRow_suggestedPersonId_fkey" FOREIGN KEY ("suggestedPersonId") REFERENCES "Person"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "StaffDirectoryImportRow" ADD CONSTRAINT "StaffDirectoryImportRow_matchedPersonId_fkey" FOREIGN KEY ("matchedPersonId") REFERENCES "Person"("id") ON DELETE SET NULL ON UPDATE CASCADE;
