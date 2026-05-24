import { NextResponse } from "next/server";
import prisma from "@/lib/prisma";

export async function GET(req: Request) {
  const url = new URL(req.url);
  const formId = url.searchParams.get("formId");

  if (!formId) {
    return NextResponse.json(
      { error: "formId required" },
      { status: 400 }
    );
  }

  try {
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
                option: true,
              },
            },
          },
        },
        tokens: true,
      },
    });

    if (!form) {
      return NextResponse.json(
        { error: "Form not found" },
        { status: 404 }
      );
    }

    const totalInvited = form.tokens.length;
    const totalResponses = form.responses.length;

    const completionRate =
      totalInvited > 0
        ? Number(((totalResponses / totalInvited) * 100).toFixed(2))
        : 0;

    // ✅ Process questions
    const questions = form.questions.map((q: any) => {
      const counts: Record<string, number> = {};

      q.options.forEach((opt: any) => {
        counts[opt.label] = 0;
      });

      let totalScore = 0;
      let count = 0;

      form.responses.forEach((res: any) => {
        res.answers.forEach((ans: any) => {
          if (ans.questionId === q.id) {
            const lbl = ans.option.label;
            counts[lbl]++;

            if (ans.option.value !== null) {
              totalScore += ans.option.value;
              count++;
            }
          }
        });
      });

      const avg = count ? totalScore / count : 0;

      return {
        questionId: q.id,
        questionText: q.text,
        averageScore: Number(avg.toFixed(2)), // ✅ single-question average
        distribution: counts,
      };
    });

    // ✅ ✅ ADD THIS (OVERALL PARTICIPANT SCORE)

    let total = 0;
    let qCount = 0;

    questions.forEach((q: any) => {
      total += q.averageScore;
      qCount++;
    });

    const overallAverage =
      qCount > 0 ? Number((total / qCount).toFixed(2)) : 0;

    // ✅ RESPONSE
    return NextResponse.json({
      formId,
      title: form.title,
      totalInvited,
      totalResponses,
      completionRate,

      questions,

      overallAverage, // ✅ ADDDED (Participant overall score)

    });

  } catch (error) {
    console.error(error);
    return NextResponse.json(
      { error: "Server error" },
      { status: 500 }
    );
  }
}