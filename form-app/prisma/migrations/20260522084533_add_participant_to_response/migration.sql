/*
  Warnings:

  - You are about to drop the column `participantId` on the `Form` table. All the data in the column will be lost.
  - You are about to drop the column `ratingMax` on the `Question` table. All the data in the column will be lost.
  - You are about to drop the column `ratingMin` on the `Question` table. All the data in the column will be lost.
  - You are about to drop the column `type` on the `Question` table. All the data in the column will be lost.

*/
-- DropForeignKey
ALTER TABLE "Form" DROP CONSTRAINT "Form_participantId_fkey";

-- AlterTable
ALTER TABLE "Form" DROP COLUMN "participantId";

-- AlterTable
ALTER TABLE "Question" DROP COLUMN "ratingMax",
DROP COLUMN "ratingMin",
DROP COLUMN "type";

-- AlterTable
ALTER TABLE "Response" ADD COLUMN     "participantId" TEXT;

-- DropEnum
DROP TYPE "QuestionType";
