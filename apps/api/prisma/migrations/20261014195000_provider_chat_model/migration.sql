-- AlterTable
ALTER TABLE "Chat" ADD COLUMN "model" TEXT;

-- RedefineTables
PRAGMA defer_foreign_keys=ON;
PRAGMA foreign_keys=OFF;
CREATE TABLE "new_AiProviderConfig" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "userId" TEXT NOT NULL,
    "provider" TEXT NOT NULL,
    "baseUrl" TEXT NOT NULL,
    "encryptedApiKey" TEXT,
    "model" TEXT,
    "enabled" BOOLEAN NOT NULL DEFAULT true,
    "createdAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" DATETIME NOT NULL,
    CONSTRAINT "AiProviderConfig_userId_fkey" FOREIGN KEY ("userId") REFERENCES "User" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);
INSERT INTO "new_AiProviderConfig" ("id", "userId", "provider", "baseUrl", "encryptedApiKey", "model", "enabled", "createdAt", "updatedAt")
SELECT "id", "userId", "provider", "baseUrl", "encryptedApiKey", "model", "enabled", "createdAt", "updatedAt" FROM "AiProviderConfig";
DROP TABLE "AiProviderConfig";
ALTER TABLE "new_AiProviderConfig" RENAME TO "AiProviderConfig";
CREATE INDEX "AiProviderConfig_userId_idx" ON "AiProviderConfig"("userId");
PRAGMA foreign_keys=ON;
PRAGMA defer_foreign_keys=OFF;
