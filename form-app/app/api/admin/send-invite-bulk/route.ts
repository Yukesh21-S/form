import { NextResponse } from "next/server";
import { z } from "zod";
import { randomUUID } from "crypto";
import * as xlsx from "xlsx";
import prisma from "@/lib/prisma";
import { INVITE_TOKEN_LIFETIME_HOURS } from "@/lib/feedback";
import { sendInviteEmail } from "@/lib/mailer";

const formIdSchema = z.string().uuid();
const emailSchema = z.string().email();

const jsonError = (message: string, status = 400) =>
  NextResponse.json({ error: message }, { status });

function parseEmailRows(sheet: xlsx.WorkSheet) {
  const rows = xlsx.utils.sheet_to_json<string[]>(sheet, { header: 1, blankrows: false });
  if (rows.length === 0) {
    return [];
  }

  const header = rows[0].map((cell) => (typeof cell === "string" ? cell.trim().toLowerCase() : ""));
  const emailIndex = header.findIndex((value) => value === "email");

  const dataRows = emailIndex === -1 ? rows : rows.slice(1);
  const emails = dataRows
    .map((row) => {
      const cell = Array.isArray(row)
        ? row[emailIndex === -1 ? 0 : emailIndex]
        : undefined;
      return typeof cell === "string" ? cell.trim() : "";
    })
    .filter((value) => value.length > 0);

  return Array.from(new Set(emails));
}

export async function POST(request: Request) {
  const formData = await request.formData();
  const formIdRaw = formData.get("formId");
  const file = formData.get("file");

  if (!formIdRaw || typeof formIdRaw !== "string") {
    return jsonError("Missing or invalid formId", 422);
  }

  const parseFormId = formIdSchema.safeParse(formIdRaw);
  if (!parseFormId.success) {
    return jsonError(parseFormId.error.message, 422);
  }

  if (!file || !(file instanceof File)) {
    return jsonError("Missing Excel file upload under field name 'file'", 422);
  }

  const form = await prisma.form.findUnique({ where: { id: parseFormId.data } });
  if (!form) {
    return jsonError("Form not found", 404);
  }

  const buffer = await file.arrayBuffer();
  const workbook = xlsx.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    return jsonError("Uploaded file contains no worksheet", 422);
  }

  const sheet = workbook.Sheets[sheetName];
  const emails = parseEmailRows(sheet);

  if (emails.length === 0) {
    return jsonError("No email addresses found in the uploaded file", 422);
  }

  const invalidEmails = emails.filter((email) => !emailSchema.safeParse(email).success);
  if (invalidEmails.length > 0) {
    return jsonError(
      `Invalid email addresses found: ${invalidEmails.join(", ")}`,
      422
    );
  }

  const results: Array<{
    email: string;
    status: "sent" | "failed";
    error?: string;
  }> = [];

  for (const email of emails) {
    const token = randomUUID();
    const expiresAt = new Date(Date.now() + INVITE_TOKEN_LIFETIME_HOURS * 60 * 60 * 1000);

    const invite = await prisma.inviteToken.create({
      data: {
        token,
        email,
        formId: parseFormId.data,
        expiresAt,
      },
    });

    try {
      await sendInviteEmail(email, token);
      results.push({ email, status: "sent" });
    } catch (error) {
      await prisma.inviteToken.delete({ where: { id: invite.id } });
      results.push({
        email,
        status: "failed",
        error: error instanceof Error ? error.message : "Unknown error",
      });
    }
  }

  return NextResponse.json({
    formId: form.id,
    total: emails.length,
    results,
  });
}
