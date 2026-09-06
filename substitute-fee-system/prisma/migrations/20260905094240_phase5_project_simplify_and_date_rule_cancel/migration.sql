/*
  Warnings:

  - You are about to drop the column `feeTypeLabel` on the `Project` table. All the data in the column will be lost.
  - You are about to drop the column `unitPrice` on the `Project` table. All the data in the column will be lost.

*/
-- AlterTable
ALTER TABLE "Project" DROP COLUMN "feeTypeLabel",
DROP COLUMN "unitPrice";

-- AlterTable
ALTER TABLE "SpecialDateRule" ADD COLUMN     "cancelledAt" TIMESTAMP(3),
ADD COLUMN     "cancelledBy" TEXT,
ADD COLUMN     "isCancelled" BOOLEAN NOT NULL DEFAULT false;
