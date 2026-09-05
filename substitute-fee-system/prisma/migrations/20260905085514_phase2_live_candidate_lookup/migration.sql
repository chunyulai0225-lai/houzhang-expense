/*
  Warnings:

  - You are about to drop the column `suggestedPersonId` on the `StaffDirectoryImportRow` table. All the data in the column will be lost.

*/
-- DropForeignKey
ALTER TABLE "StaffDirectoryImportRow" DROP CONSTRAINT "StaffDirectoryImportRow_suggestedPersonId_fkey";

-- AlterTable
ALTER TABLE "StaffDirectoryImportRow" DROP COLUMN "suggestedPersonId";
