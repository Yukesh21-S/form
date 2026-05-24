import { NextResponse } from "next/server";
import prisma from "@/lib/prisma";

export async function GET(req: Request) {
  const url = new URL(req.url);

  //////////////////////////////////////////////////////
  // QUERY PARAMS
  //////////////////////////////////////////////////////

  const formId =
    url.searchParams.get("formId");

  const participantId =
    url.searchParams.get("participantId");

  if (!formId) {
    return NextResponse.json(
      { error: "formId required" },
      { status: 400 }
    );
  }

  //////////////////////////////////////////////////////
  // FETCH FORM
  //////////////////////////////////////////////////////

  const form =
    await prisma.form.findUnique({
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

          orderBy: {
            createdAt: "desc",
          },
        },
      },
    });

  if (!form) {
    return NextResponse.json(
      { error: "Form not found" },
      { status: 404 }
    );
  }

  //////////////////////////////////////////////////////
  // FILTER RESPONSES
  //////////////////////////////////////////////////////

  const filteredResponses =
    participantId
      ? form.responses.filter(
          (response: any) =>
            response.participantId ===
            participantId
        )
      : form.responses;

  //////////////////////////////////////////////////////
  // ANALYTICS
  //////////////////////////////////////////////////////

  const results =
    form.questions.map(
      (question: any) => {
        let total = 0;

        let validResponses = 0;

        const distribution: Record<
          string,
          number
        > = {
          Rarely: 0,
          Sometimes: 0,
          Often: 0,
          Always: 0,
          "Insufficient Exposure": 0,
        };

        //////////////////////////////////////////////////////
        // LOOP RESPONSES
        //////////////////////////////////////////////////////

        filteredResponses.forEach(
          (response: any) => {
            response.answers.forEach(
              (answer: any) => {
                if (
                  answer.questionId ===
                  question.id
                ) {
                  const option =
                    answer.option;

                  if (!option) return;

                  //////////////////////////////////////////////////////
                  // DISTRIBUTION
                  //////////////////////////////////////////////////////

                  distribution[
                    option.label
                  ]++;

                  //////////////////////////////////////////////////////
                  // SCORE
                  //////////////////////////////////////////////////////

                  if (
                    option.value !== null
                  ) {
                    total +=
                      option.value;

                    validResponses++;
                  }
                }
              }
            );
          }
        );

        //////////////////////////////////////////////////////
        // AVERAGE
        //////////////////////////////////////////////////////

        const rawAverage =
          validResponses > 0
            ? total /
              validResponses
            : 0;

        const roundedScore =
          Math.round(rawAverage);

        //////////////////////////////////////////////////////
        // BAND
        //////////////////////////////////////////////////////

        let band = "";

        if (roundedScore > 85) {
          band =
            "Consistently observed";
        } else if (
          roundedScore >= 70
        ) {
          band =
            "Moderately observed";
        } else {
          band =
            "Inconsistently observed";
        }

        //////////////////////////////////////////////////////
        // RETURN
        //////////////////////////////////////////////////////

        return {
          questionId: question.id,

          question: question.text,

          distribution,

          totalScore: total,

          validResponses,

          rawAverage: Number(
            rawAverage.toFixed(2)
          ),

          roundedScore,

          band,
        };
      }
    );

  //////////////////////////////////////////////////////
  // RESPONSE
  //////////////////////////////////////////////////////

  return NextResponse.json({
    formId,

    participantId,

    totalResponses:
      filteredResponses.length,

    results,
  });
}