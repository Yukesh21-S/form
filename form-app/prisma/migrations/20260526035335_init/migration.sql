/*
  Warnings:

  - You are about to drop the column `relationshipType` on the `InviteToken` table. All the data in the column will be lost.
  - You are about to drop the column `category` on the `Question` table. All the data in the column will be lost.
  - You are about to drop the column `relationshipType` on the `Response` table. All the data in the column will be lost.

*/
-- AlterTable
ALTER TABLE "InviteToken" DROP COLUMN "relationshipType";

-- AlterTable
ALTER TABLE "Question" DROP COLUMN "category";

-- AlterTable
ALTER TABLE "Response" DROP COLUMN "relationshipType";
