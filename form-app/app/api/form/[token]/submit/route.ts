import { NextResponse } from "next/server";
import { z } from "zod";
import prisma from "@/lib/prisma";
import { getValidInviteToken } from "@/lib/feedback";

const submitSchema = z.object({
  answers: z
    .array(
      z.object({
        questionId: z.string().uuid(),
        optionId: z.string().uuid(),
      })
    )
    .min(1),
});

const jsonError = (
  message: string,
  status = 400
) =>
  NextResponse.json(
    { error: message },
    { status }
  );

export async function POST(
  request: Request,
  {
    params,
  }: {
    params: Promise<{ token: string }>;
  }
) {
  const { token } = await params;

  if (!token) {
    return jsonError(
      "Token is required",
      400
    );
  }

  const body = await request
    .json()
    .catch(() => null);

  if (!body) {
    return jsonError(
      "Invalid JSON payload",
      400
    );
  }

  const parseResult =
    submitSchema.safeParse(body);

  if (!parseResult.success) {
    return jsonError(
      parseResult.error.message,
      422
    );
  }

  const { answers } =
    parseResult.data;

  const invite =
    await getValidInviteToken(token);

  if (!invite) {
    return jsonError(
      "Invalid or expired token",
      404
    );
  }


  //////////////////////////////////////////////////////
  // VALIDATE QUESTIONS
  //////////////////////////////////////////////////////

  const questionMap = new Map<
    string,
    any
  >(
    invite.form.questions.map(
      (question: any) => [
        question.id,
        question,
      ]
    )
  );

  const questionIds =
    new Set<string>();

  for (const answer of answers) {
    if (
      !questionMap.has(
        answer.questionId
      )
    ) {
      return jsonError(
        "Answer contains invalid questionId",
        422
      );
    }

    const question =
      questionMap.get(
        answer.questionId
      )!;

    const option =
      question.options.find(
        (item: any) =>
          item.id === answer.optionId
      );

    if (!option) {
      return jsonError(
        "Answer contains invalid optionId",
        422
      );
    }

    if (
      questionIds.has(
        answer.questionId
      )
    ) {
      return jsonError(
        "Duplicate answers for the same question are not allowed",
        422
      );
    }

    questionIds.add(
      answer.questionId
    );
  }

  //////////////////////////////////////////////////////
  // CREATE RESPONSE
  //////////////////////////////////////////////////////

  const response =
    await prisma.response.create({
      data: {
        email: invite.email,

        formId: invite.formId,

        //////////////////////////////////////////////////////
        // PARTICIPANT
        //////////////////////////////////////////////////////

        participantId:
          invite.participantId,

        answers: {
          create: answers.map(
            (answer) => ({
              questionId:
                answer.questionId,

              optionId:
                answer.optionId,
            })
          ),
        },
      },
    });

  //////////////////////////////////////////////////////
  // MARK TOKEN USED
  //////////////////////////////////////////////////////

  await prisma.inviteToken.update({
    where: {
      id: invite.id,
    },

    data: {
      isUsed: true,
    },
  });

  return NextResponse.json(
    {
      success: true,
      responseId: response.id,
    },
    { status: 201 }
  );
}