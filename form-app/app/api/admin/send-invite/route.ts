import { NextResponse } from "next/server";
import { z } from "zod";
import { randomUUID } from "crypto";
import prisma from "@/lib/prisma";
import { INVITE_TOKEN_LIFETIME_HOURS } from "@/lib/feedback";
import { sendInviteEmail } from "@/lib/mailer";

const inviteSchema = z.object({
  email: z.string().min(1),
  formId: z.string().uuid(),
});

const emailSchema = z.string().email();

const jsonError = (message: string, status = 400) =>
  NextResponse.json({ error: message }, { status });

export async function POST(request: Request) {
  const body = await request.json().catch(() => null);

  if (!body) {
    return jsonError("Invalid JSON payload", 400);
  }

  const parseResult = inviteSchema.safeParse(body);
  if (!parseResult.success) {
    return jsonError(parseResult.error.message, 422);
  }

  const { email: emailRaw, formId } = parseResult.data;
  const emails = emailRaw
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);

  if (emails.length === 0) {
    return jsonError("No valid email address provided", 422);
  }

  const invalidEmails = emails.filter((email) => !emailSchema.safeParse(email).success);
  if (invalidEmails.length > 0) {
    return jsonError(
      `Invalid email address(es) provided: ${invalidEmails.join(", ")}`,
      422
    );
  }

  const form = await prisma.form.findUnique({ where: { id: formId } });
  if (!form) {
    return jsonError("Form not found", 404);
  }

  const results: Array<{ email: string; status: "sent" | "failed"; error?: string }> = [];

  for (const email of emails) {
    const token = randomUUID();
    const expiresAt = new Date(Date.now() + INVITE_TOKEN_LIFETIME_HOURS * 60 * 60 * 1000);

    const invite = await prisma.inviteToken.create({
      data: {
        token,
        email,
        formId,
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
        error: error instanceof Error ? error.message : "Unknown mailer error",
      });
    }
  }

  return NextResponse.json({
    formId,
    total: emails.length,
    results,
  }, { status: 201 });
}
