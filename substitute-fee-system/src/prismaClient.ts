import { PrismaClient } from "@prisma/client";

// 單一共用的 Prisma Client 實例，供所有 service 使用。
export const prisma = new PrismaClient();
