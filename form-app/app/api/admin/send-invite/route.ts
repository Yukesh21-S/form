import { NextResponse } from "next/server";
import { z } from "zod";
import { randomUUID } from "crypto";

import prisma from "@/lib/prisma";

import {
  INVITE_TOKEN_LIFETIME_HOURS,
} from "@/lib/feedback";

import { sendInviteEmail } from "@/lib/mailer";

//////////////////////////////////////////////////////
// VALIDATION
//////////////////////////////////////////////////////

const inviteSchema = z.object({
  email: z.string().min(1),

  formId: z.string().uuid(),

  participantId: z.string(),
});

const emailSchema = z.string().email();

//////////////////////////////////////////////////////
// ERROR RESPONSE
//////////////////////////////////////////////////////

const jsonError = (
  message: string,
  status = 400
) =>
  NextResponse.json(
    { error: message },
    { status }
  );

//////////////////////////////////////////////////////
// POST
//////////////////////////////////////////////////////

export async function POST(
  request: Request
) {

  //////////////////////////////////////////////////////
  // PARSE BODY
  //////////////////////////////////////////////////////

  const body = await request
    .json()
    .catch(() => null);

  if (!body) {

    return jsonError(
      "Invalid JSON payload",
      400
    );
  }

  //////////////////////////////////////////////////////
  // VALIDATE BODY
  //////////////////////////////////////////////////////

  const parseResult =
    inviteSchema.safeParse(body);

  if (!parseResult.success) {

    return jsonError(
      parseResult.error.message,
      422
    );
  }

  //////////////////////////////////////////////////////
  // EXTRACT DATA
  //////////////////////////////////////////////////////

  const {
    email: emailRaw,
    formId,
    participantId,
  } = parseResult.data;

  //////////////////////////////////////////////////////
  // SPLIT EMAILS
  //////////////////////////////////////////////////////

  const emails = emailRaw
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);

  if (emails.length === 0) {

    return jsonError(
      "No valid email address provided",
      422
    );
  }

  //////////////////////////////////////////////////////
  // VALIDATE EMAILS
  //////////////////////////////////////////////////////

  const invalidEmails = emails.filter(
    (email) =>
      !emailSchema.safeParse(email)
        .success
  );

  if (invalidEmails.length > 0) {

    return jsonError(
      `Invalid email address(es) provided: ${invalidEmails.join(", ")}`,
      422
    );
  }

  //////////////////////////////////////////////////////
  // CHECK FORM EXISTS
  //////////////////////////////////////////////////////

  const form =
    await prisma.form.findUnique({
      where: {
        id: formId,
      },
    });

  if (!form) {

    return jsonError(
      "Form not found",
      404
    );
  }

  //////////////////////////////////////////////////////
  // CHECK PARTICIPANT EXISTS
  //////////////////////////////////////////////////////

  const participant =
    await prisma.participant.findUnique({
      where: {
        id: participantId,
      },
    });

  if (!participant) {

    return jsonError(
      "Participant not found",
      404
    );
  }

  //////////////////////////////////////////////////////
  // RESULTS ARRAY
  //////////////////////////////////////////////////////

  const results: Array<{
    email: string;
    status: "sent" | "failed";
    error?: string;
  }> = [];

  //////////////////////////////////////////////////////
  // SEND INVITES
  //////////////////////////////////////////////////////

  for (const email of emails) {

    //////////////////////////////////////////////////////
    // GENERATE TOKEN
    //////////////////////////////////////////////////////

    const token = randomUUID();

    //////////////////////////////////////////////////////
    // EXPIRY
    //////////////////////////////////////////////////////

    const expiresAt = new Date(
      Date.now() +
      INVITE_TOKEN_LIFETIME_HOURS *
      60 *
      60 *
      1000
    );

    //////////////////////////////////////////////////////
    // CREATE INVITE TOKEN
    //////////////////////////////////////////////////////

    const invite =
      await prisma.inviteToken.create({
        data: {

          token,

          email,

          formId,

          participantId,

          expiresAt,
        },
      });

    //////////////////////////////////////////////////////
    // SEND EMAIL
    //////////////////////////////////////////////////////

    try {

      await sendInviteEmail(
        email,
        token
      );

      results.push({
        email,
        status: "sent",
      });

    } catch (error) {

      //////////////////////////////////////////////////////
      // DELETE TOKEN IF EMAIL FAILS
      //////////////////////////////////////////////////////

      await prisma.inviteToken.delete({
        where: {
          id: invite.id,
        },
      });

      results.push({

        email,

        status: "failed",

        error:
          error instanceof Error
            ? error.message
            : "Unknown mailer error",
      });
    }
  }

  //////////////////////////////////////////////////////
  // RESPONSE
  //////////////////////////////////////////////////////

  return NextResponse.json(
    {
      success: true,

      formId,

      participantId,

      participantName:
        participant.fullName,

      total: emails.length,

      results,
    },
    {
      status: 201,
    }
  );
}