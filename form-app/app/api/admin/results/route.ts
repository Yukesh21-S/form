import { NextResponse } from "next/server";
import { z } from "zod";
import prisma from "@/lib/prisma";

const querySchema = z.object({
  formId: z.string().uuid(),
});

const jsonError = (message: string, status = 400) =>
  NextResponse.json({ error: message }, { status });

export async function GET(request: Request) {
  const url = new URL(request.url);
  const formId = url.searchParams.get("formId") ?? "";

  const parseResult = querySchema.safeParse({ formId });
  if (!parseResult.success) {
    return jsonError(parseResult.error.message, 422);
  }

  const form = await prisma.form.findUnique({
    where: { id: formId },
    include: {
      questions: {
        include: {
          options: true,
        },
      },
      responses: {
        include: {
          answers: {
            include: {
              question: true,
              option: true,
            },
          },
        },
      },
      tokens: true, // ✅ ADDED (to count invites)
    },
  });

  if (!form) {
    return jsonError("Form not found", 404);
  }

  // ✅ NEW CALCULATIONS
  const totalInvited = form.tokens.length;
  const totalResponses = form.responses.length;

  const completionRate =
    totalInvited > 0
      ? Number(((totalResponses / totalInvited) * 100).toFixed(2))
      : 0;

  // ✅ QUESTION ANALYTICS
  const questions = form.questions.map((question: any) => {
    const distribution = question.options.map((option: any) => ({
      optionId: option.id,
      label: option.label,
      count: 0,
    }));

    let totalScore = 0;
    let countWithScore = 0;

    form.responses.forEach((response: any) => {
      response.answers
        .filter((answer: any) => answer.questionId === question.id)
        .forEach((answer: any) => {
          const option = answer.option;

          const bucket = distribution.find(
            (item: any) => item.optionId === option.id
          );

          if (bucket) {
            bucket.count += 1;
          }

          if (option.value !== null) {
            totalScore += option.value;
            countWithScore += 1;
          }
        });
    });

    const averageScore =
      countWithScore > 0 ? totalScore / countWithScore : null;

    return {
      questionId: question.id,
      questionText: question.text,
      averageScore,
      distribution,
    };
  });

  // ✅ FORMAT RESPONSES
  const responses = form.responses.map((response: any) => ({
    id: response.id,
    email: response.email,
    createdAt: response.createdAt.toISOString(),
    answers: response.answers.map((answer: any) => ({
      questionId: answer.questionId,
      questionText: answer.question.text,
      optionId: answer.option.id,
      optionLabel: answer.option.label,
      optionValue: answer.option.value,
    })),
  }));

  // ✅ FINAL RESPONSE
  return NextResponse.json({
    formId: form.id,
    title: form.title,

    totalInvited,       // ✅ NEW
    totalResponses,     // ✅ EXISTING
    completionRate,     // ✅ NEW

    questions,
    responses,
  });
}