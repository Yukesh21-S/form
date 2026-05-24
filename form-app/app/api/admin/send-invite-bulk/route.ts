import { NextResponse } from "next/server";
import { z } from "zod";

import { randomUUID } from "crypto";

import * as xlsx from "xlsx";

import prisma from "@/lib/prisma";

import {
  INVITE_TOKEN_LIFETIME_HOURS,
  RELATIONSHIP_TYPES,
} from "@/lib/feedback";

import { sendInviteEmail } from "@/lib/mailer";

//////////////////////////////////////////////////////
// VALIDATION
//////////////////////////////////////////////////////

const formIdSchema = z.string().uuid();

const emailSchema = z.string().email();

const relationshipTypeSchema = z
  .enum(RELATIONSHIP_TYPES)
  .default("OTHER");

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
// PARSE EXCEL EMAILS
//////////////////////////////////////////////////////

function parseEmailRows(
  sheet: xlsx.WorkSheet
) {

  const rows =
    xlsx.utils.sheet_to_json<string[]>(
      sheet,
      {
        header: 1,
        blankrows: false,
      }
    );

  if (rows.length === 0) {
    return [];
  }

  const header = rows[0].map((cell) =>
    typeof cell === "string"
      ? cell.trim().toLowerCase()
      : ""
  );

  const emailIndex =
    header.findIndex(
      (value) => value === "email"
    );

  const dataRows =
    emailIndex === -1
      ? rows
      : rows.slice(1);

  const emails = dataRows
    .map((row) => {

      const cell = Array.isArray(row)
        ? row[
            emailIndex === -1
              ? 0
              : emailIndex
          ]
        : undefined;

      return typeof cell === "string"
        ? cell.trim()
        : "";
    })
    .filter(
      (value) => value.length > 0
    );

  return Array.from(new Set(emails));
}

//////////////////////////////////////////////////////
// POST
//////////////////////////////////////////////////////

export async function POST(
  request: Request
) {

  //////////////////////////////////////////////////////
  // FORM DATA
  //////////////////////////////////////////////////////

  const formData =
    await request.formData();

  //////////////////////////////////////////////////////
  // GET VALUES
  //////////////////////////////////////////////////////

  const formIdRaw =
    formData.get("formId");

  const participantId =
    formData.get("participantId");

  const file =
    formData.get("file");

  //////////////////////////////////////////////////////
  // RELATIONSHIP TYPE (optional, defaults to OTHER)
  //////////////////////////////////////////////////////

  const relationshipTypeRaw =
    formData.get("relationshipType");

  const parsedRelType =
    relationshipTypeSchema.safeParse(
      typeof relationshipTypeRaw === "string"
        ? relationshipTypeRaw
        : "OTHER"
    );

  const relationshipType =
    parsedRelType.success
      ? parsedRelType.data
      : "OTHER";

  //////////////////////////////////////////////////////
  // VALIDATE FORM ID
  //////////////////////////////////////////////////////

  if (
    !formIdRaw ||
    typeof formIdRaw !== "string"
  ) {

    return jsonError(
      "Missing or invalid formId",
      422
    );
  }

  //////////////////////////////////////////////////////
  // VALIDATE PARTICIPANT ID
  //////////////////////////////////////////////////////

  if (
    !participantId ||
    typeof participantId !== "string"
  ) {

    return jsonError(
      "Missing participantId",
      422
    );
  }

  //////////////////////////////////////////////////////
  // VALIDATE FORM UUID
  //////////////////////////////////////////////////////

  const parseFormId =
    formIdSchema.safeParse(
      formIdRaw
    );

  if (!parseFormId.success) {

    return jsonError(
      parseFormId.error.message,
      422
    );
  }

  //////////////////////////////////////////////////////
  // VALIDATE FILE
  //////////////////////////////////////////////////////

  if (
    !file ||
    !(file instanceof File)
  ) {

    return jsonError(
      "Missing Excel file upload under field name 'file'",
      422
    );
  }

  //////////////////////////////////////////////////////
  // CHECK FORM EXISTS
  //////////////////////////////////////////////////////

  const form =
    await prisma.form.findUnique({
      where: {
        id: parseFormId.data,
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
  // READ EXCEL FILE
  //////////////////////////////////////////////////////

  const buffer =
    await file.arrayBuffer();

  const workbook = xlsx.read(
    buffer,
    {
      type: "array",
    }
  );

  const sheetName =
    workbook.SheetNames[0];

  if (!sheetName) {

    return jsonError(
      "Uploaded file contains no worksheet",
      422
    );
  }

  const sheet =
    workbook.Sheets[sheetName];

  const emails =
    parseEmailRows(sheet);

  //////////////////////////////////////////////////////
  // VALIDATE EMAIL LIST
  //////////////////////////////////////////////////////

  if (emails.length === 0) {

    return jsonError(
      "No email addresses found in the uploaded file",
      422
    );
  }

  //////////////////////////////////////////////////////
  // INVALID EMAILS
  //////////////////////////////////////////////////////

  const invalidEmails =
    emails.filter(
      (email) =>
        !emailSchema.safeParse(email)
          .success
    );

  if (invalidEmails.length > 0) {

    return jsonError(
      `Invalid email addresses found: ${invalidEmails.join(", ")}`,
      422
    );
  }

  //////////////////////////////////////////////////////
  // RESULTS
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
    // TOKEN
    //////////////////////////////////////////////////////

    const token =
      randomUUID();

    //////////////////////////////////////////////////////
    // EXPIRY
    //////////////////////////////////////////////////////

    const expiresAt =
      new Date(
        Date.now() +
        INVITE_TOKEN_LIFETIME_HOURS *
        60 *
        60 *
        1000
      );

    //////////////////////////////////////////////////////
    // CREATE INVITE TOKEN
    //////////////////////////////////////////////////////

    //////////////////////////////////////////////////////
    // AUTO-DETECT RELATIONSHIP TYPE
    //////////////////////////////////////////////////////

    const finalRelationshipType =
      email.toLowerCase() === participant.email.toLowerCase()
        ? "SELF"
        : relationshipType;

    //////////////////////////////////////////////////////
    // CREATE INVITE TOKEN
    //////////////////////////////////////////////////////

    const invite =
      await prisma.inviteToken.create({
        data: {

          token,

          email,

          formId:
            parseFormId.data,

          participantId,

          //////////////////////////////////////////////////////
          // STORE RELATIONSHIP TYPE
          //////////////////////////////////////////////////////

          relationshipType: finalRelationshipType,

          //////////////////////////////////////////////////////

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
      // DELETE IF EMAIL FAILS
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
            : "Unknown error",
      });
    }
  }

  //////////////////////////////////////////////////////
  // RESPONSE
  //////////////////////////////////////////////////////

  return NextResponse.json({
    success: true,

    formId: form.id,

    participantId,

    participantName:
      participant.fullName,

    total: emails.length,

    results,
  });
}