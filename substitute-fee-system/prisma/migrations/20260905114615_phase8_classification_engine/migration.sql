-- AlterEnum
-- This migration adds more than one value to an enum.
-- With PostgreSQL versions 11 and earlier, this is not possible
-- in a single migration. This can be worked around by creating
-- multiple migrations, each migration adding only one value to
-- the enum.


ALTER TYPE "ClassificationMethod" ADD VALUE 'CONFLICT';
ALTER TYPE "ClassificationMethod" ADD VALUE 'TEACHER_UNMATCHED';

-- AlterTable
ALTER TABLE "SubstituteRecord" ADD COLUMN     "autoClassificationMethod" "ClassificationMethod",
ADD COLUMN     "autoClassificationRuleId" TEXT,
ADD COLUMN     "autoFundingSource" "FundingSource",
ADD COLUMN     "autoProjectId" TEXT,
ADD COLUMN     "classificationRuleId" TEXT,
ADD COLUMN     "classifiedAt" TIMESTAMP(3),
ADD COLUMN     "conflictCandidatesJson" TEXT,
ADD COLUMN     "manualOverrideAt" TIMESTAMP(3),
ADD COLUMN     "manualOverrideBy" TEXT,
ADD COLUMN     "manualOverrideReason" TEXT;
