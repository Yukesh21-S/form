import { NextResponse } from "next/server";
import prisma from "@/lib/prisma";

export async function GET(req: Request) {
  const url = new URL(req.url);

  //////////////////////////////////////////////////////
  // QUERY PARAMS
  //////////////////////////////////////////////////////

  const formId = url.searchParams.get("formId");

  const participantId = url.searchParams.get("participantId");

  if (!formId) {
    return NextResponse.json(
      { error: "formId required" },
      { status: 400 }
    );
  }

  //////////////////////////////////////////////////////
  // FETCH FORM
  //////////////////////////////////////////////////////

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
  // FILTER BY PARTICIPANT
  //////////////////////////////////////////////////////

  const filteredResponses = participantId
    ? form.responses.filter(
        (r: any) => r.participantId === participantId
      )
    : form.responses;

  //////////////////////////////////////////////////////
  // ALL RESPONSES (enterprise benchmark)
  //////////////////////////////////////////////////////

  const allResponses = form.responses;

  //////////////////////////////////////////////////////
  // SELF vs OTHERS split using relationshipType
  //////////////////////////////////////////////////////

  const selfResponses = filteredResponses.filter(
    (r: any) => r.relationshipType === "SELF"
  );

  const othersResponses = filteredResponses.filter(
    (r: any) => r.relationshipType !== "SELF"
  );

  //////////////////////////////////////////////////////
  // HELPER: compute analytics per question set
  //////////////////////////////////////////////////////

  function computeResults(
    questions: any[],
    responses: any[]
  ) {
    return questions.map((question: any) => {
      let total = 0;
      let validResponses = 0;

      const distribution: Record<string, number> = {
        Rarely: 0,
        Sometimes: 0,
        Often: 0,
        Always: 0,
        "Insufficient Exposure": 0,
      };

      responses.forEach((response: any) => {
        response.answers.forEach((answer: any) => {
          if (answer.questionId === question.id) {
            const option = answer.option;

            if (!option) return;

            distribution[option.label]++;

            if (option.value !== null) {
              total += option.value;
              validResponses++;
            }
          }
        });
      });

      const rawAverage =
        validResponses > 0 ? total / validResponses : 0;

      const roundedScore = Math.round(rawAverage);

      let band = "";

      if (roundedScore > 85) {
        band = "Consistently observed";
      } else if (roundedScore >= 70) {
        band = "Moderately observed";
      } else {
        band = "Inconsistently observed";
      }

      return {
        questionId: question.id,

        question: question.text,

        category: question.category ?? null,

        distribution,

        totalScore: total,

        validResponses,

        rawAverage: Number(rawAverage.toFixed(2)),

        roundedScore,

        band,
      };
    });
  }

  //////////////////////////////////////////////////////
  // RESULTS
  //////////////////////////////////////////////////////

  const results = computeResults(form.questions, filteredResponses);

  //////////////////////////////////////////////////////
  // ENTERPRISE BENCHMARK
  //////////////////////////////////////////////////////

  const enterpriseResults = computeResults(form.questions, allResponses);

  //////////////////////////////////////////////////////
  // SELF vs OTHERS results
  //////////////////////////////////////////////////////

  const selfResults = computeResults(form.questions, selfResponses);

  const othersResults = computeResults(form.questions, othersResponses);

  //////////////////////////////////////////////////////
  // RESPONSE BREAKDOWN BY TYPE
  //////////////////////////////////////////////////////

  const responsesByType: Record<string, number> = {};

  filteredResponses.forEach((r: any) => {
    const type = r.relationshipType ?? "UNKNOWN";
    responsesByType[type] = (responsesByType[type] ?? 0) + 1;
  });

  //////////////////////////////////////////////////////
  // COMPETENCY GROUPING
  //////////////////////////////////////////////////////

  const leadingSelf = results.filter(
    (q) => q.category === "Leading Self"
  );

  const leadingOthers = results.filter(
    (q) => q.category === "Leading Others"
  );

  const enterpriseLeadingSelf = enterpriseResults.filter(
    (q) => q.category === "Leading Self"
  );

  const enterpriseLeadingOthers = enterpriseResults.filter(
    (q) => q.category === "Leading Others"
  );

  //////////////////////////////////////////////////////
  // RESPONSE
  //////////////////////////////////////////////////////

  return NextResponse.json({
    formId,

    participantId,

    totalResponses: filteredResponses.length,

    responsesByType,

    results,

    //////////////////////////////////////////////////////
    // COMPETENCY GROUPS
    //////////////////////////////////////////////////////

    competencyGroups: {
      leadingSelf,
      leadingOthers,
    },

    //////////////////////////////////////////////////////
    // ENTERPRISE BENCHMARK
    //////////////////////////////////////////////////////

    enterpriseBenchmark: {
      results: enterpriseResults,

      competencyGroups: {
        leadingSelf: enterpriseLeadingSelf,
        leadingOthers: enterpriseLeadingOthers,
      },
    },

    //////////////////////////////////////////////////////
    // SELF vs OTHERS (uses relationshipType field)
    //////////////////////////////////////////////////////

    selfVsOthers: {
      self: selfResults,
      others: othersResults,
      selfResponseCount: selfResponses.length,
      othersResponseCount: othersResponses.length,
    },
  });
}