-- CreateEnum
CREATE TYPE "SemesterStatus" AS ENUM ('NOT_STARTED', 'ACTIVE', 'ENDED');

-- CreateEnum
CREATE TYPE "OvertimeMatchMode" AS ENUM ('TEACHER_WEEKDAY_PERIOD', 'TEACHER_WEEKDAY_PERIOD_SUBJECT');

-- CreateEnum
CREATE TYPE "EnrollmentStatus" AS ENUM ('ENROLLED', 'NOT_ENROLLED');

-- CreateEnum
CREATE TYPE "RoleType" AS ENUM ('HOMEROOM_TEACHER', 'SUBJECT_TEACHER', 'ADMIN', 'DIRECTOR', 'SECTION_CHIEF', 'PRINCIPAL', 'OTHER');

-- CreateEnum
CREATE TYPE "FeeType" AS ENUM ('SUBSTITUTE_PERIOD', 'OVERTIME_PERIOD', 'HOMEROOM_SUBSTITUTE', 'DAILY_WAGE', 'HALF_DAY_WAGE', 'OTHER');

-- CreateEnum
CREATE TYPE "RuleType" AS ENUM ('OVERTIME', 'PROJECT');

-- CreateEnum
CREATE TYPE "Weekday" AS ENUM ('MON', 'TUE', 'WED', 'THU', 'FRI');

-- CreateEnum
CREATE TYPE "OverrideClassification" AS ENUM ('GENERAL', 'OVERTIME', 'PROJECT');

-- CreateEnum
CREATE TYPE "StaffType" AS ENUM ('BD', 'NON_BD', 'UNKNOWN');

-- CreateEnum
CREATE TYPE "FundingSource" AS ENUM ('GENERAL', 'OVERTIME', 'PROJECT', 'UNDETERMINED');

-- CreateEnum
CREATE TYPE "ClassificationMethod" AS ENUM ('WEEKLY_RULE', 'DATE_EXCEPTION', 'GENERAL_DEFAULT', 'MANUAL_OVERRIDE');

-- CreateEnum
CREATE TYPE "MonthlyStatus" AS ENUM ('NOT_IMPORTED', 'IMPORTED', 'PENDING_REVIEW', 'CONFIRMED', 'LOCKED');

-- CreateTable
CREATE TABLE "Semester" (
    "id" TEXT NOT NULL,
    "schoolYear" INTEGER NOT NULL,
    "term" INTEGER NOT NULL,
    "status" "SemesterStatus" NOT NULL DEFAULT 'NOT_STARTED',
    "isCurrent" BOOLEAN NOT NULL DEFAULT false,
    "startDate" TIMESTAMP(3) NOT NULL,
    "endDate" TIMESTAMP(3) NOT NULL,
    "overtimeMatchMode" "OvertimeMatchMode" NOT NULL DEFAULT 'TEACHER_WEEKDAY_PERIOD',
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Semester_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Person" (
    "id" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "payrollCode" TEXT,
    "enrollmentStatus" "EnrollmentStatus" NOT NULL DEFAULT 'ENROLLED',
    "enrollDate" TIMESTAMP(3),
    "leaveDate" TIMESTAMP(3),
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Person_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PersonCode" (
    "id" TEXT NOT NULL,
    "personId" TEXT NOT NULL,
    "schoolYear" INTEGER,
    "categoryName" TEXT NOT NULL,
    "categoryCode" TEXT,
    "originalStaffCode" TEXT,
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "PersonCode_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "BdClassificationRule" (
    "id" TEXT NOT NULL,
    "codeValue" TEXT NOT NULL,
    "isBd" BOOLEAN NOT NULL,
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "BdClassificationRule_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "SemesterRole" (
    "id" TEXT NOT NULL,
    "personId" TEXT NOT NULL,
    "semesterId" TEXT NOT NULL,
    "roleType" "RoleType" NOT NULL,
    "roleDetail" TEXT,
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "SemesterRole_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "FeeRule" (
    "id" TEXT NOT NULL,
    "semesterId" TEXT NOT NULL,
    "feeType" "FeeType" NOT NULL,
    "amount" DECIMAL(65,30) NOT NULL,
    "effectiveDate" TIMESTAMP(3) NOT NULL,
    "endDate" TIMESTAMP(3),
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "FeeRule_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "PeriodSlot" (
    "id" TEXT NOT NULL,
    "code" TEXT NOT NULL,
    "displayName" TEXT NOT NULL,
    "sortOrder" INTEGER NOT NULL,
    "isTeachingPeriod" BOOLEAN NOT NULL DEFAULT true,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "PeriodSlot_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Project" (
    "id" TEXT NOT NULL,
    "semesterId" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "feeTypeLabel" TEXT NOT NULL,
    "unitPrice" DECIMAL(65,30) NOT NULL,
    "isActive" BOOLEAN NOT NULL DEFAULT true,
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Project_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "SpecialWeeklyRule" (
    "id" TEXT NOT NULL,
    "semesterId" TEXT NOT NULL,
    "personId" TEXT NOT NULL,
    "ruleType" "RuleType" NOT NULL,
    "projectId" TEXT,
    "weekday" "Weekday" NOT NULL,
    "periodCode" TEXT NOT NULL,
    "subject" TEXT,
    "weeklyPeriods" DECIMAL(65,30) NOT NULL,
    "effectiveDate" TIMESTAMP(3) NOT NULL,
    "endDate" TIMESTAMP(3),
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "SpecialWeeklyRule_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "SpecialDateRule" (
    "id" TEXT NOT NULL,
    "semesterId" TEXT NOT NULL,
    "date" TIMESTAMP(3) NOT NULL,
    "personId" TEXT NOT NULL,
    "periodCode" TEXT NOT NULL,
    "originalClassificationNote" TEXT,
    "overrideClassification" "OverrideClassification" NOT NULL,
    "projectId" TEXT,
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "SpecialDateRule_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "SchoolCalendarDay" (
    "id" TEXT NOT NULL,
    "semesterId" TEXT,
    "date" TIMESTAMP(3) NOT NULL,
    "weekday" "Weekday" NOT NULL,
    "isTeachingDay" BOOLEAN NOT NULL,
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "SchoolCalendarDay_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "MonthlyImport" (
    "id" TEXT NOT NULL,
    "semesterId" TEXT NOT NULL,
    "year" INTEGER NOT NULL,
    "month" INTEGER NOT NULL,
    "importedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "fileName" TEXT NOT NULL,
    "importedBy" TEXT,
    "totalCount" INTEGER NOT NULL DEFAULT 0,
    "successCount" INTEGER NOT NULL DEFAULT 0,
    "errorCount" INTEGER NOT NULL DEFAULT 0,
    "versionNo" INTEGER NOT NULL DEFAULT 1,
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "MonthlyImport_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "SubstituteRecordRaw" (
    "id" TEXT NOT NULL,
    "monthlyImportId" TEXT NOT NULL,
    "rowNumber" INTEGER,
    "originalTeacherText" TEXT,
    "dateText" TEXT,
    "leaveTypeText" TEXT,
    "hoursOrDaysText" TEXT,
    "periodText" TEXT,
    "classText" TEXT,
    "subjectText" TEXT,
    "substituteTeacherText" TEXT,
    "teacherCertText" TEXT,
    "payGradeText" TEXT,
    "homeroomFeeText" TEXT,
    "rawJson" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "SubstituteRecordRaw_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "SubstituteRecord" (
    "id" TEXT NOT NULL,
    "rawRecordId" TEXT NOT NULL,
    "monthlyImportId" TEXT NOT NULL,
    "originalTeacherId" TEXT,
    "substituteTeacherId" TEXT,
    "date" TIMESTAMP(3) NOT NULL,
    "weekday" "Weekday" NOT NULL,
    "periodCode" TEXT,
    "className" TEXT,
    "subject" TEXT,
    "leaveType" TEXT,
    "rawHoursOrDays" TEXT,
    "periodCount" DECIMAL(65,30),
    "staffType" "StaffType" NOT NULL DEFAULT 'UNKNOWN',
    "fundingSource" "FundingSource" NOT NULL DEFAULT 'UNDETERMINED',
    "projectId" TEXT,
    "unitPrice" DECIMAL(65,30),
    "amount" DECIMAL(65,30),
    "classificationMethod" "ClassificationMethod" NOT NULL DEFAULT 'GENERAL_DEFAULT',
    "isManuallyModified" BOOLEAN NOT NULL DEFAULT false,
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "SubstituteRecord_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "MonthlyLock" (
    "id" TEXT NOT NULL,
    "semesterId" TEXT,
    "year" INTEGER NOT NULL,
    "month" INTEGER NOT NULL,
    "status" "MonthlyStatus" NOT NULL DEFAULT 'NOT_IMPORTED',
    "lockedAt" TIMESTAMP(3),
    "lockedBy" TEXT,
    "unlockedAt" TIMESTAMP(3),
    "unlockedBy" TEXT,
    "note" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "MonthlyLock_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "ChangeLog" (
    "id" TEXT NOT NULL,
    "tableName" TEXT NOT NULL,
    "recordId" TEXT NOT NULL,
    "fieldName" TEXT,
    "oldValue" TEXT,
    "newValue" TEXT,
    "changedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "changedBy" TEXT,
    "reason" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "ChangeLog_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "Semester_isCurrent_idx" ON "Semester"("isCurrent");

-- CreateIndex
CREATE UNIQUE INDEX "Semester_schoolYear_term_key" ON "Semester"("schoolYear", "term");

-- CreateIndex
CREATE INDEX "Person_enrollmentStatus_idx" ON "Person"("enrollmentStatus");

-- CreateIndex
CREATE INDEX "PersonCode_personId_idx" ON "PersonCode"("personId");

-- CreateIndex
CREATE INDEX "PersonCode_categoryCode_idx" ON "PersonCode"("categoryCode");

-- CreateIndex
CREATE UNIQUE INDEX "BdClassificationRule_codeValue_key" ON "BdClassificationRule"("codeValue");

-- CreateIndex
CREATE INDEX "SemesterRole_personId_idx" ON "SemesterRole"("personId");

-- CreateIndex
CREATE INDEX "SemesterRole_semesterId_idx" ON "SemesterRole"("semesterId");

-- CreateIndex
CREATE INDEX "FeeRule_semesterId_feeType_effectiveDate_idx" ON "FeeRule"("semesterId", "feeType", "effectiveDate");

-- CreateIndex
CREATE UNIQUE INDEX "PeriodSlot_code_key" ON "PeriodSlot"("code");

-- CreateIndex
CREATE INDEX "Project_semesterId_idx" ON "Project"("semesterId");

-- CreateIndex
CREATE INDEX "SpecialWeeklyRule_semesterId_personId_weekday_periodCode_idx" ON "SpecialWeeklyRule"("semesterId", "personId", "weekday", "periodCode");

-- CreateIndex
CREATE INDEX "SpecialDateRule_semesterId_personId_date_periodCode_idx" ON "SpecialDateRule"("semesterId", "personId", "date", "periodCode");

-- CreateIndex
CREATE UNIQUE INDEX "SchoolCalendarDay_date_key" ON "SchoolCalendarDay"("date");

-- CreateIndex
CREATE INDEX "SchoolCalendarDay_semesterId_idx" ON "SchoolCalendarDay"("semesterId");

-- CreateIndex
CREATE INDEX "MonthlyImport_semesterId_year_month_idx" ON "MonthlyImport"("semesterId", "year", "month");

-- CreateIndex
CREATE INDEX "SubstituteRecordRaw_monthlyImportId_idx" ON "SubstituteRecordRaw"("monthlyImportId");

-- CreateIndex
CREATE UNIQUE INDEX "SubstituteRecord_rawRecordId_key" ON "SubstituteRecord"("rawRecordId");

-- CreateIndex
CREATE INDEX "SubstituteRecord_monthlyImportId_idx" ON "SubstituteRecord"("monthlyImportId");

-- CreateIndex
CREATE INDEX "SubstituteRecord_originalTeacherId_idx" ON "SubstituteRecord"("originalTeacherId");

-- CreateIndex
CREATE INDEX "SubstituteRecord_substituteTeacherId_idx" ON "SubstituteRecord"("substituteTeacherId");

-- CreateIndex
CREATE INDEX "SubstituteRecord_date_idx" ON "SubstituteRecord"("date");

-- CreateIndex
CREATE UNIQUE INDEX "MonthlyLock_year_month_key" ON "MonthlyLock"("year", "month");

-- CreateIndex
CREATE INDEX "ChangeLog_tableName_recordId_idx" ON "ChangeLog"("tableName", "recordId");

-- AddForeignKey
ALTER TABLE "PersonCode" ADD CONSTRAINT "PersonCode_personId_fkey" FOREIGN KEY ("personId") REFERENCES "Person"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SemesterRole" ADD CONSTRAINT "SemesterRole_personId_fkey" FOREIGN KEY ("personId") REFERENCES "Person"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SemesterRole" ADD CONSTRAINT "SemesterRole_semesterId_fkey" FOREIGN KEY ("semesterId") REFERENCES "Semester"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "FeeRule" ADD CONSTRAINT "FeeRule_semesterId_fkey" FOREIGN KEY ("semesterId") REFERENCES "Semester"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "Project" ADD CONSTRAINT "Project_semesterId_fkey" FOREIGN KEY ("semesterId") REFERENCES "Semester"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SpecialWeeklyRule" ADD CONSTRAINT "SpecialWeeklyRule_semesterId_fkey" FOREIGN KEY ("semesterId") REFERENCES "Semester"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SpecialWeeklyRule" ADD CONSTRAINT "SpecialWeeklyRule_personId_fkey" FOREIGN KEY ("personId") REFERENCES "Person"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SpecialWeeklyRule" ADD CONSTRAINT "SpecialWeeklyRule_projectId_fkey" FOREIGN KEY ("projectId") REFERENCES "Project"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SpecialDateRule" ADD CONSTRAINT "SpecialDateRule_semesterId_fkey" FOREIGN KEY ("semesterId") REFERENCES "Semester"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SpecialDateRule" ADD CONSTRAINT "SpecialDateRule_personId_fkey" FOREIGN KEY ("personId") REFERENCES "Person"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SpecialDateRule" ADD CONSTRAINT "SpecialDateRule_projectId_fkey" FOREIGN KEY ("projectId") REFERENCES "Project"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SchoolCalendarDay" ADD CONSTRAINT "SchoolCalendarDay_semesterId_fkey" FOREIGN KEY ("semesterId") REFERENCES "Semester"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "MonthlyImport" ADD CONSTRAINT "MonthlyImport_semesterId_fkey" FOREIGN KEY ("semesterId") REFERENCES "Semester"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SubstituteRecordRaw" ADD CONSTRAINT "SubstituteRecordRaw_monthlyImportId_fkey" FOREIGN KEY ("monthlyImportId") REFERENCES "MonthlyImport"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SubstituteRecord" ADD CONSTRAINT "SubstituteRecord_rawRecordId_fkey" FOREIGN KEY ("rawRecordId") REFERENCES "SubstituteRecordRaw"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SubstituteRecord" ADD CONSTRAINT "SubstituteRecord_monthlyImportId_fkey" FOREIGN KEY ("monthlyImportId") REFERENCES "MonthlyImport"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SubstituteRecord" ADD CONSTRAINT "SubstituteRecord_originalTeacherId_fkey" FOREIGN KEY ("originalTeacherId") REFERENCES "Person"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SubstituteRecord" ADD CONSTRAINT "SubstituteRecord_substituteTeacherId_fkey" FOREIGN KEY ("substituteTeacherId") REFERENCES "Person"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SubstituteRecord" ADD CONSTRAINT "SubstituteRecord_projectId_fkey" FOREIGN KEY ("projectId") REFERENCES "Project"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "MonthlyLock" ADD CONSTRAINT "MonthlyLock_semesterId_fkey" FOREIGN KEY ("semesterId") REFERENCES "Semester"("id") ON DELETE SET NULL ON UPDATE CASCADE;
