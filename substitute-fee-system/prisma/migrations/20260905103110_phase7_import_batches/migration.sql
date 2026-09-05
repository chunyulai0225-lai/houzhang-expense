-- CreateEnum
CREATE TYPE "ImportBatchStatus" AS ENUM ('ACTIVE', 'SUPERSEDED');

-- AlterTable
ALTER TABLE "MonthlyImport" ADD COLUMN     "status" "ImportBatchStatus" NOT NULL DEFAULT 'ACTIVE';

-- AlterTable
ALTER TABLE "SubstituteRecordRaw" ADD COLUMN     "originalTeacherCodeText" TEXT,
ADD COLUMN     "substituteTeacherCodeText" TEXT;

-- CreateTable
CREATE TABLE "MonthlyImportError" (
    "id" TEXT NOT NULL,
    "monthlyImportId" TEXT NOT NULL,
    "rowNumber" INTEGER,
    "fieldName" TEXT,
    "message" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "MonthlyImportError_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "MonthlyImportError_monthlyImportId_idx" ON "MonthlyImportError"("monthlyImportId");

-- AddForeignKey
ALTER TABLE "MonthlyImportError" ADD CONSTRAINT "MonthlyImportError_monthlyImportId_fkey" FOREIGN KEY ("monthlyImportId") REFERENCES "MonthlyImport"("id") ON DELETE RESTRICT ON UPDATE CASCADE;
