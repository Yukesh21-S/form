import { NextRequest } from "next/server";
import pptxgen from "pptxgenjs";
import prisma from "@/lib/prisma";

//////////////////////////////////////////////////////
// COLORS
//////////////////////////////////////////////////////

const COLORS = {
  bg: "F3F3F3",

  green: "00A000",

  orange: "E67E22",

  red: "E05A6A",

  darkRed: "B00000",

  brightRed: "E53935",

  beige: "B8AA97",

  lightPink: "F3D6D8",

  lightBeige: "E7DED0",

  lightGreen: "DDEEDB",

  grayHeader: "D9D9D9",

  line: "CFCFCF",

  text: "222222",
};

//////////////////////////////////////////////////////
// LAYOUT
//////////////////////////////////////////////////////

const ROW_HEIGHT = 0.31;

const START_Y = 1.38;

const QUESTION_X = 0.52;

const QUESTION_W = 3.55;

const BAR_X = 4.15;

const BAR_W = 2.55;

const COUNT_START_X = 7.05;

const COUNT_W = 0.92;

//////////////////////////////////////////////////////
// SCORE COLOR
//////////////////////////////////////////////////////

function getScoreColor(score: number) {
  if (score > 85) {
    return COLORS.green;
  }

  if (score >= 70) {
    return COLORS.orange;
  }

  return COLORS.red;
}

//////////////////////////////////////////////////////
// ROUTE
//////////////////////////////////////////////////////

export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url);

  const formId = searchParams.get("formId");

  const participantId =
    searchParams.get("participantId");

  if (!formId) {
    return Response.json(
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
        },
      },
    });

  if (!form) {
    return Response.json(
      { error: "Form not found" },
      { status: 404 }
    );
  }

  //////////////////////////////////////////////////////
  // FILTER PARTICIPANT
  //////////////////////////////////////////////////////

  const filteredResponses =
    participantId
      ? form.responses.filter(
          (r: any) =>
            r.participantId ===
            participantId
        )
      : form.responses;

  //////////////////////////////////////////////////////
  // ANALYTICS
  //////////////////////////////////////////////////////

  const analytics =
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

                  distribution[
                    option.label
                  ]++;

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

        const rawAverage =
          validResponses > 0
            ? total /
              validResponses
            : 0;

        const roundedScore =
          Math.round(rawAverage);

        return {
          question: question.text,

          roundedScore,

          distribution,
        };
      }
    );

  //////////////////////////////////////////////////////
  // PPT
  //////////////////////////////////////////////////////

  const pptx = new pptxgen();

  pptx.layout = "LAYOUT_WIDE";

  pptx.author = "ChatGPT";

  pptx.subject = "360 Feedback";

  pptx.company = "Ericsson";

  //////////////////////////////////////////////////////
  // SLIDE 1
  //////////////////////////////////////////////////////

  const slide1 = pptx.addSlide();

  slide1.background = {
    color: COLORS.bg,
  };

  //////////////////////////////////////////////////////
  // TITLE
  //////////////////////////////////////////////////////

  slide1.addText(
    "Overall results",
    {
      x: 0.38,
      y: 0.22,

      w: 4.5,
      h: 0.5,

      fontFace: "Segoe UI",

      fontSize: 30,

      color: "111111",

      margin: 0,
    }
  );

  slide1.addText(
    "How consistently micro-behaviors are seen in practice by your respondents.",
    {
      x: 0.4,
      y: 0.88,

      w: 7,
      h: 0.2,

      fontFace: "Segoe UI",

      fontSize: 11,

      color: "444444",

      margin: 0,
    }
  );

  //////////////////////////////////////////////////////
  // LEGEND
  //////////////////////////////////////////////////////

  const legends = [
    {
      color: COLORS.green,

      text: "Consistently observed (>85%)",

      x: 6.9,
    },

    {
      color: COLORS.orange,

      text: "Moderately observed (70-85%)",

      x: 9.0,
    },

    {
      color: COLORS.red,

      text: "Inconsistently observed (<70%)",

      x: 11.25,
    },
  ];

  legends.forEach((l) => {
    slide1.addShape(
      pptx.ShapeType.rect,
      {
        x: l.x,
        y: 0.22,

        w: 0.18,
        h: 0.18,

        fill: {
          color: l.color,
        },

        line: {
          color: l.color,
        },
      }
    );

    slide1.addText(l.text, {
      x: l.x + 0.24,
      y: 0.18,

      w: 1.8,
      h: 0.3,

      fontFace: "Segoe UI",

      fontSize: 8.5,

      color: COLORS.text,

      margin: 0,
    });
  });

  //////////////////////////////////////////////////////
  // CONTENT
  //////////////////////////////////////////////////////

  let leftY = 1.45;

  let rightY = 1.45;

  analytics.forEach(
    (item, index) => {
      const isLeft =
        index < 10;

      const currentY = isLeft
        ? leftY
        : rightY;

      const textX = isLeft
        ? 0.38
        : 6.65;

      const barX = isLeft
        ? 3.82
        : 10.08;

      const scoreX = isLeft
        ? 5.15
        : 11.42;

      //////////////////////////////////////////////////////
      // QUESTION
      //////////////////////////////////////////////////////

      slide1.addText(
        item.question,
        {
          x: textX,
          y: currentY,

          w: 3.1,
          h: 0.35,

          fontFace: "Segoe UI",

          fontSize: 10,

          color: COLORS.text,

          fit: "shrink",

          margin: 0,

          breakLine: false,
        }
      );

      //////////////////////////////////////////////////////
      // LINE
      //////////////////////////////////////////////////////

      slide1.addShape(
        pptx.ShapeType.line,
        {
          x: barX,
          y: currentY - 0.02,

          w: 0,
          h: 0.38,

          line: {
            color: "8E8E8E",
            pt: 1,
          },
        }
      );

      //////////////////////////////////////////////////////
      // BAR
      //////////////////////////////////////////////////////

      slide1.addShape(
        pptx.ShapeType.rect,
        {
          x: barX,
          y: currentY + 0.1,

          w:
            item.roundedScore /
            100,

          h: 0.16,

          fill: {
            color: getScoreColor(
              item.roundedScore
            ),
          },

          line: {
            color: getScoreColor(
              item.roundedScore
            ),
          },
        }
      );

      //////////////////////////////////////////////////////
      // SCORE
      //////////////////////////////////////////////////////

      slide1.addText(
        `${item.roundedScore}%`,
        {
          x: scoreX,
          y: currentY + 0.02,

          w: 0.5,
          h: 0.2,

          fontFace: "Segoe UI",

          fontSize: 14,

          bold: false,

          color: COLORS.text,

          margin: 0,
        }
      );

      if (isLeft) {
        leftY += 0.48;
      } else {
        rightY += 0.48;
      }
    }
  );

  //////////////////////////////////////////////////////
  // SLIDE 2
  //////////////////////////////////////////////////////

  const slide2 = pptx.addSlide();

  slide2.background = {
    color: COLORS.bg,
  };

  //////////////////////////////////////////////////////
  // TITLE
  //////////////////////////////////////////////////////

  slide2.addText(
    "Results distribution",
    {
      x: 0.15,
      y: 0.2,

      w: 4,
      h: 0.5,

      fontFace: "Segoe UI",

      fontSize: 30,

      color: "111111",

      margin: 0,
    }
  );

  //////////////////////////////////////////////////////
  // TABLE HEADER
  //////////////////////////////////////////////////////

  slide2.addShape(
    pptx.ShapeType.rect,
    {
      x: 0.05,
      y: 0.98,

      w: 11.35,
      h: 0.34,

      fill: {
        color: COLORS.grayHeader,
      },

      line: {
        color: COLORS.grayHeader,
      },
    }
  );

  slide2.addText("Results", {
    x: 0.22,
    y: 1.06,

    fontFace: "Segoe UI",

    fontSize: 8.5,

    bold: true,

    margin: 0,
  });

  slide2.addText("Distribution", {
    x: 5.1,
    y: 1.06,

    fontFace: "Segoe UI",

    fontSize: 8.5,

    bold: true,

    margin: 0,
  });

  //////////////////////////////////////////////////////
  // HEADERS
  //////////////////////////////////////////////////////

  const headers = [
    {
      label: "Rarely",

      x: 7.02,

      color: COLORS.darkRed,
    },

    {
      label: "Sometimes",

      x: 7.97,

      color: COLORS.brightRed,
    },

    {
      label: "Often",

      x: 8.92,

      color: COLORS.beige,
    },

    {
      label: "Always",

      x: 9.87,

      color: COLORS.green,
    },
  ];

  headers.forEach((h) => {
    slide2.addShape(
      pptx.ShapeType.rect,
      {
        x: h.x,
        y: 0.98,

        w: 0.95,
        h: 0.34,

        fill: {
          color: h.color,
        },

        line: {
          color: h.color,
        },
      }
    );

    slide2.addText(h.label, {
      x: h.x,
      y: 1.06,

      w: 0.95,
      h: 0.1,

      align: "center",

      fontFace: "Segoe UI",

      fontSize: 8,

      bold: true,

      color: "FFFFFF",

      margin: 0,
    });
  });

  //////////////////////////////////////////////////////
  // ROWS
  //////////////////////////////////////////////////////

  let y = START_Y;

  analytics.forEach((item) => {
    //////////////////////////////////////////////////////
    // ROW LINE
    //////////////////////////////////////////////////////

    slide2.addShape(
      pptx.ShapeType.line,
      {
        x: 0.05,
        y: y + ROW_HEIGHT,

        w: 11.35,
        h: 0,

        line: {
          color: COLORS.line,
          pt: 0.5,
        },
      }
    );

    //////////////////////////////////////////////////////
    // QUESTION
    //////////////////////////////////////////////////////

    slide2.addText(
      item.question,
      {
        x: QUESTION_X,
        y: y + 0.02,

        w: QUESTION_W,
        h: ROW_HEIGHT,

        fontFace: "Segoe UI",

        fontSize: 8,

        color: COLORS.text,

        fit: "shrink",

        valign: "middle",

        margin: 0,
      }
    );

    //////////////////////////////////////////////////////
    // TOTAL
    //////////////////////////////////////////////////////

    const total =
      item.distribution.Rarely +
      item.distribution.Sometimes +
      item.distribution.Often +
      item.distribution.Always;

    //////////////////////////////////////////////////////
    // BAR BACKGROUND
    //////////////////////////////////////////////////////

    slide2.addShape(
      pptx.ShapeType.rect,
      {
        x: BAR_X,
        y: y + 0.08,

        w: BAR_W,
        h: 0.16,

        fill: {
          color: "EEEEEE",
        },

        line: {
          color: "EEEEEE",
        },
      }
    );

    //////////////////////////////////////////////////////
    // SEGMENTS
    //////////////////////////////////////////////////////

    const segments = [
      {
        value:
          item.distribution.Rarely,

        color: COLORS.darkRed,
      },

      {
        value:
          item.distribution.Sometimes,

        color: COLORS.brightRed,
      },

      {
        value:
          item.distribution.Often,

        color: COLORS.beige,
      },

      {
        value:
          item.distribution.Always,

        color: COLORS.green,
      },
    ];

    let segmentX = BAR_X;

    segments.forEach((s) => {
      const width =
        total > 0
          ? (s.value / total) *
            BAR_W
          : 0;

      slide2.addShape(
        pptx.ShapeType.rect,
        {
          x: segmentX,
          y: y + 0.08,

          w: width,
          h: 0.16,

          fill: {
            color: s.color,
          },

          line: {
            color: s.color,
          },
        }
      );

      segmentX += width;
    });

    //////////////////////////////////////////////////////
    // COUNT BACKGROUND
    //////////////////////////////////////////////////////

    const bgColumns = [
      {
        x: 7.02,
        color: COLORS.lightPink,
      },

      {
        x: 7.97,
        color: COLORS.lightPink,
      },

      {
        x: 8.92,
        color: COLORS.lightBeige,
      },

      {
        x: 9.87,
        color: COLORS.lightGreen,
      },
    ];

    bgColumns.forEach((c) => {
      slide2.addShape(
        pptx.ShapeType.rect,
        {
          x: c.x,
          y: y,

          w: 0.95,
          h: ROW_HEIGHT,

          fill: {
            color: c.color,
          },

          line: {
            color: c.color,
          },
        }
      );
    });

    //////////////////////////////////////////////////////
    // COUNTS
    //////////////////////////////////////////////////////

    const counts = [
      item.distribution.Rarely,

      item.distribution.Sometimes,

      item.distribution.Often,

      item.distribution.Always,
    ];

    counts.forEach(
      (count, idx) => {
        slide2.addText(
          `${count}`,
          {
            x:
              COUNT_START_X +
              idx * COUNT_W,

            y: y + 0.09,

            w: COUNT_W,
            h: 0.1,

            align: "center",

            fontFace: "Segoe UI",

            fontSize: 8.5,

            color: COLORS.text,

            margin: 0,
          }
        );
      }
    );

    y += ROW_HEIGHT;
  });
//////////////////////////////////////////////////////
// SLIDE 3 - INSUFFICIENT EXPOSURES
//////////////////////////////////////////////////////

const slide3 = pptx.addSlide();

slide3.background = {
  color: "F2F2F2",
};

//////////////////////////////////////////////////////
// TITLE
//////////////////////////////////////////////////////

slide3.addText(
  "Insufficient Exposures",
  {
    x: 0.25,
    y: 0.05,

    w: 5.8,
    h: 0.55,

    fontFace: "Segoe UI",

    fontSize: 28,

    color: "1A1A1A",

    bold: false,

    margin: 0,
  }
);

//////////////////////////////////////////////////////
// TABLE HEADER
//////////////////////////////////////////////////////

slide3.addShape(
  pptx.ShapeType.rect,
  {
    x: 0.15,
    y: 0.82,

    w: 7.1,
    h: 0.3,

    fill: {
      color: "D0D0D0",
    },

    line: {
      color: "D0D0D0",
    },
  }
);

slide3.addText(
  "Results",
  {
    x: 0.22,
    y: 0.89,

    w: 1,
    h: 0.1,

    fontFace: "Segoe UI",

    fontSize: 8.5,

    bold: true,

    color: "111111",

    margin: 0,
  }
);

slide3.addText(
  "Count of Insufficient Exposure",
  {
    x: 4.15,
    y: 0.89,

    w: 2.4,
    h: 0.1,

    fontFace: "Segoe UI",

    fontSize: 8.5,

    bold: true,

    color: "111111",

    margin: 0,
  }
);

//////////////////////////////////////////////////////
// TABLE ROWS
//////////////////////////////////////////////////////

let tableY = 1.12;

analytics.forEach((item) => {

  const count =
    item.distribution[
      "Insufficient Exposure"
    ] || 0;

  //////////////////////////////////////////////////////
  // ROW LINE
  //////////////////////////////////////////////////////

  slide3.addShape(
    pptx.ShapeType.line,
    {
      x: 0.15,
      y: tableY + 0.27,

      w: 7.1,
      h: 0,

      line: {
        color: "C8C8C8",
        pt: 0.5,
      },
    }
  );

  //////////////////////////////////////////////////////
  // QUESTION TEXT
  //////////////////////////////////////////////////////

  slide3.addText(
    item.question,
    {
      x: 0.22,
      y: tableY + 0.03,

      w: 3.8,
      h: 0.2,

      fontFace: "Segoe UI",

      fontSize: 8.2,

      color: "222222",

      margin: 0,

      fit: "shrink",
    }
  );

  //////////////////////////////////////////////////////
  // LEFT VALUE
  //////////////////////////////////////////////////////

  slide3.addText(
    `${count}`,
    {
      x: 4.2,
      y: tableY + 0.03,

      w: 0.2,
      h: 0.1,

      fontFace: "Consolas",

      fontSize: 9,

      color: "222222",

      margin: 0,
    }
  );

  //////////////////////////////////////////////////////
  // BLUE BAR
  //////////////////////////////////////////////////////

  if (count > 0) {

    const barWidth =
      count * 0.32;

    slide3.addShape(
      pptx.ShapeType.rect,
      {
        x: 4.18,
        y: tableY + 0.16,

        w: barWidth,
        h: 0.1,

        fill: {
          color: "1D6FD8",
        },

        line: {
          color: "1D6FD8",
        },
      }
    );

    //////////////////////////////////////////////////////
    // RIGHT VALUE
    //////////////////////////////////////////////////////

    slide3.addText(
      `${count}`,
      {
        x: 4.18 + barWidth + 0.04,
        y: tableY + 0.11,

        w: 0.2,
        h: 0.1,

        fontFace: "Segoe UI",

        fontSize: 9,

        color: "444444",

        margin: 0,
      }
    );
  }

  tableY += 0.29;
});

//////////////////////////////////////////////////////
// NOTE HEADER
//////////////////////////////////////////////////////

slide3.addShape(
  pptx.ShapeType.rect,
  {
    x: 7.8,
    y: 3.95,

    w: 4.35,
    h: 0.3,

    fill: {
      color: "A5A5A5",
    },

    line: {
      color: "A5A5A5",
    },
  }
);

slide3.addText(
  "Note",
  {
    x: 7.92,
    y: 4.02,

    w: 0.5,
    h: 0.1,

    fontFace: "Segoe UI",

    fontSize: 9,

    bold: true,

    color: "FFFFFF",

    margin: 0,
  }
);

//////////////////////////////////////////////////////
// NOTE BODY
//////////////////////////////////////////////////////

slide3.addShape(
  pptx.ShapeType.rect,
  {
    x: 7.8,
    y: 4.25,

    w: 4.35,
    h: 2.65,

    fill: {
      color: "EFEFEF",
    },

    line: {
      color: "EFEFEF",
    },
  }
);

slide3.addText(
  `Insufficient Exposure is what a rater selects when they feel they have not seen sufficient evidence from you on this behavior to provide a rating. These answers are set aside and they do not lower your score.

A high count can itself be a signal as it can mean a behavior is not be visible to the people around you, even if you feel you are demonstrating it.

If you have areas with high insufficient exposure, we recommend you consider why you are receiving this feedback and what you could do to address this.`,
  {
    x: 7.95,
    y: 4.38,

    w: 3.95,
    h: 2.25,

    fontFace: "Segoe UI",

    fontSize: 8.8,

    color: "222222",

    valign: "top",

    breakLine: true,

    fit: "shrink",

    margin: 0,
  }
); 
  //////////////////////////////////////////////////////
  // EXPORT
  //////////////////////////////////////////////////////

  const pptxData =
    await pptx.write({
      outputType:
        "nodebuffer",
    });
  
  const uint8 =
    pptxData instanceof
    Uint8Array
      ? pptxData
      : new Uint8Array(
          pptxData as ArrayBuffer
        );

  const arrayBuffer =
    uint8.buffer.slice(
      uint8.byteOffset,
      uint8.byteOffset +
        uint8.byteLength
    ) as ArrayBuffer;

  return new Response(
    arrayBuffer,
    {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",

        "Content-Disposition":
          `attachment; filename=report-${formId}.pptx`,
      },
    }
  );
}