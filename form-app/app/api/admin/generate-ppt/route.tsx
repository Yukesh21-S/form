import { NextRequest } from "next/server";
import pptxgen from "pptxgenjs";
import prisma from "@/lib/prisma";
import fs from "fs";
import path from "path";
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
  blue: "1D6FD8",
  teal: "00897B",
  orange2: "E67E22",
};

//////////////////////////////////////////////////////
// LAYOUT CONSTANTS
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
  if (score > 85) return COLORS.green;
  if (score >= 70) return COLORS.orange;
  return COLORS.red;
}

//////////////////////////////////////////////////////
// ANALYTICS HELPER
//////////////////////////////////////////////////////

function computeAnalytics(questions: any[], responses: any[]) {
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

    return {
      question: question.text,
      category: question.category ?? null,
      roundedScore,
      rawAverage: Number(rawAverage.toFixed(2)),
      distribution,
      validResponses,
    };
  });
}

//////////////////////////////////////////////////////
// ROUTE
//////////////////////////////////////////////////////

export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url);

  const formId = searchParams.get("formId");
  const participantId = searchParams.get("participantId");

  if (!formId) {
    return Response.json(
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
      questions: { include: { options: true } },
      responses: {
        include: {
          answers: { include: { option: true } },
        },
      },
    },
  });

  if (!form) {
    return Response.json({ error: "Form not found" }, { status: 404 });
  }

  //////////////////////////////////////////////////////
  // RESPONSE POOLS
  //////////////////////////////////////////////////////

  // All responses targeting this participant (or all if no participantId)
  const filteredResponses = participantId
    ? form.responses.filter(
      (r: any) => r.participantId === participantId
    )
    : form.responses;

  // All responses across the form (enterprise benchmark)
  const allResponses = form.responses;

  // SELF = responses where relationshipType === "SELF"
  const selfResponses = filteredResponses.filter(
    (r: any) => r.relationshipType === "SELF"
  );

  // OTHERS = everything that is not SELF
  const othersResponses = filteredResponses.filter(
    (r: any) => r.relationshipType !== "SELF"
  );

  //////////////////////////////////////////////////////
  // COMPUTE ANALYTICS
  //////////////////////////////////////////////////////

  const analytics = computeAnalytics(form.questions, filteredResponses);
  const selfAnalytics = computeAnalytics(form.questions, selfResponses);
  const managerAnalytics = computeAnalytics(form.questions, filteredResponses.filter(r => r.relationshipType === 'MANAGER'));
  const peerAnalytics = computeAnalytics(form.questions, filteredResponses.filter(r => r.relationshipType === 'PEER'));
  const directReportAnalytics = computeAnalytics(form.questions, filteredResponses.filter(r => r.relationshipType === 'DIRECT_REPORT'));
  const otherRoleAnalytics = computeAnalytics(form.questions, filteredResponses.filter(r => r.relationshipType === 'OTHER'));
  const othersAnalytics = computeAnalytics(form.questions, othersResponses);
  const enterpriseAnalytics = computeAnalytics(form.questions, allResponses);

  //////////////////////////////////////////////////////
  // COMPETENCY GROUPS
  //////////////////////////////////////////////////////

  const leadingSelf = analytics.filter((q) => q.category === "Leading Self");
  const leadingOthers = analytics.filter((q) => q.category === "Leading Others");

  const enterpriseLeadingSelf = enterpriseAnalytics.filter((q) => q.category === "Leading Self");
  const enterpriseLeadingOthers = enterpriseAnalytics.filter((q) => q.category === "Leading Others");

  // NEW: Specifically for the blue line in Slide 4 (SELF ONLY)
  const selfLeadingSelf = selfAnalytics.filter((q) => q.category === "Leading Self");
  const selfLeadingOthers = selfAnalytics.filter((q) => q.category === "Leading Others");

  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "ChatGPT";
  pptx.subject = "360 Feedback";
  pptx.company = "Ericsson";

const slideIntro = pptx.addSlide();
slideIntro.background = { color: "111111" };

// TITLE
slideIntro.addText("What is a 360 diagnostic", {
  x: 0.4, y: 0.25, w: 9, h: 0.9,
  fontFace: "Aptos", fontSize: 40, bold: false, color: "FFFFFF", margin: 0,
});

// ─── DIAGRAM IMAGE (replaces all the arrow/circle drawing code) ────────────


const diagramPath = path.join(process.cwd(), "public", "360_diagram.png");
const diagramBase64 = fs.readFileSync(diagramPath).toString("base64");

slideIntro.addImage({
  data: "image/png;base64," + diagramBase64,
  x: 0.2,
  y: 1.2,
  w: 5.8,
  h: 5.8,
});

// ─── RIGHT SIDE TEXT ───────────────────────────────────────────────────────
const TX = 7.0;
const TW = 5.9;

slideIntro.addText(
  "A 360 diagnostic is a powerful development tool. It brings together feedback from your manager, peers, direct reports and other stakeholders to show how consistently your leadership behaviors are experienced in practice.",
  { x: TX, y: 1.6, w: TW, h: 1.5, fontFace: "Aptos", fontSize: 15, color: "FFFFFF", breakLine: true, margin: 0, valign: "top" }
);
slideIntro.addText(
  "No single perspective tells the full story. The value of a 360 is that it gives you a broader view of your leadership impact: what others see clearly, what may be less visible, and where your own view may differ from the experience of those around you.",
  { x: TX, y: 3.3, w: TW, h: 1.5, fontFace: "Aptos", fontSize: 15, color: "FFFFFF", breakLine: true, margin: 0, valign: "top" }
);
slideIntro.addText(
  "The micro-behaviors covered in this report reflect observable leadership behaviors relevant to your organization's context. This means the feedback is focused on what people see you do day to day — not on personality, intent or potential. The report is intended to support reflection, coaching and practical development planning.",
  { x: TX, y: 4.9, w: TW, h: 2.0, fontFace: "Aptos", fontSize: 15, color: "FFFFFF", breakLine: true, margin: 0, valign: "top" }
);
//  ------------------------------------------------------------

 const slideHowCreated = pptx.addSlide();
  slideHowCreated.background = { color: "2D2D2D" };

  slideHowCreated.addText("How this report was\ncreated", {
    x: 0.4, y: 0.25, w: 6, h: 0.75,
    fontFace: "Segoe UI", fontSize: 44, bold: false, color: "FFFFFF", margin: 0,
  });

  // Blue banner notification (top right)
  slideHowCreated.addShape(pptx.ShapeType.rect, {
    x: 8.0, y: 0.3, w: 2.5, h: 0.55,
    fill: { color: "1E56DB" }, line: { color: "1E56DB" },
  });
  slideHowCreated.addText("Micro-behaviors not yet\nfinalized", {
    x: 8.0, y: 0.3, w: 2.5, h: 0.55,
    fontFace: "Segoe UI", fontSize: 7.5, bold: true, color: "FFFFFF",
    align: "center", valign: "middle", margin: 0,
  });

  // LEFT SECTION - BACKGROUND
  slideHowCreated.addText("Background", {
    x: 0.4, y: 1.1, w: 2.3, h: 0.35,
    fontFace: "Segoe UI", fontSize: 12, bold: true, color: "FFFFFF", margin: 0,
  });
  slideHowCreated.addText(
    "You nominated colleagues across manager, peer, direct report and broader stakeholder groups to give a balanced view of your leadership.\n\nEach rater assessed you against the micro-behaviors on the right. Micro-behaviors are deliberately observable and specific interpretations of your organization's leadership framework.",
    { x: 0.4, y: 1.5, w: 2.3, h: 2.2, fontFace: "Segoe UI", fontSize: 8.5, color: "FFFFFF", valign: "top", breakLine: true, margin: 0 }
  );

  // VERTICAL DIVIDER LINE
  // In slideHowCreated — fix the vertical divider
slideHowCreated.addShape(pptx.ShapeType.line, {
  x: 3.0, y: 1.1, w: 0.001, h: 5.0,  // ← tiny epsilon, not 0
  line: { color: "555555", pt: 1.5 },
});

  // BLUE CHEVRON DECORATION (left-center)
  slideHowCreated.addShape(pptx.ShapeType.rect, {
    x: 2.7, y: 2.7, w: 0.2, h: 0.6,
    fill: { color: "1E56DB" }, line: { color: "1E56DB" }, rotate: 45,
  });

  // MIDDLE & RIGHT SECTIONS - MICRO-BEHAVIORS
  slideHowCreated.addText("Micro-Behaviors", {
    x: 3.3, y: 1.1, w: 6.2, h: 0.35,
    fontFace: "Segoe UI", fontSize: 12, bold: true, color: "FFFFFF", margin: 0,
  });

  const leftBehaviors = [
    "Takes ownership of outcomes end-to-end",
    "Prioritizes highest-value work",
    "Makes timely decisions",
    "Questions assumptions and ways of working",
    "Adapts based on feedback",
    "Runs continuous experiments instead of waiting for perfect solutions",
    "Uses external signals to guide decisions",
    "Proactively seeks feedback to improve impact",
    "Follows through on commitments",
    "Remains constructive under pressure",
    "Sets direction and priorities",
  ];

  let behaviorY = 1.55;
  leftBehaviors.forEach((behavior) => {
    slideHowCreated.addText(behavior, {
      x: 3.3, y: behaviorY, w: 3.0, h: 0.28,
      fontFace: "Segoe UI", fontSize: 8.5, color: "82B7E8",
      valign: "top", breakLine: true, margin: 0,
    });
    behaviorY += 0.28;
  });

  const rightBehaviors = [
    "Facilitates collective momentum across teams",
    "Removes obstacles to progress",
    "Summarizes what was said to confirm understanding",
    "Develops others for long-term impact",
    "Connects work to vision and strategy",
    "Gives others ownership within boundaries",
    "Creates environment for others to thrive and perform",
    "Stops work that no longer adds value",
    "Engages multiple perspectives",
    "Seeks input from affected stakeholders before moving forward",
    "Sets high standards for performance",
  ];

  let rightBehaviorY = 1.55;
  rightBehaviors.forEach((behavior) => {
    slideHowCreated.addText(behavior, {
      x: 6.5, y: rightBehaviorY, w: 3.0, h: 0.28,
      fontFace: "Segoe UI", fontSize: 8.5, color: "82B7E8",
      valign: "top", breakLine: true, margin: 0,
    });
    rightBehaviorY += 0.28;
  });

  // ---------------------------------------------------------

  const slideHowToUse = pptx.addSlide();
  slideHowToUse.background = { color: "111111" };
 
  const HTU_CARD_X:   number[] = [0.3, 3.52, 6.74, 9.96];
  const HTU_CARD_W:   number   = 3.07;
  const HTU_HEADER_Y: number   = 1.5;
  const HTU_HEADER_H: number   = 0.78;
  const HTU_BODY_Y:   number   = HTU_HEADER_Y + HTU_HEADER_H;
  const HTU_BODY_H:   number   = 3.9;
  const HTU_BANNER_Y: number   = 6.28;
  const HTU_BANNER_H: number   = 0.88;
  const HTU_HEADER_BG: string  = "2D6DB5";
  const HTU_BODY_BG:   string  = "E8E8E8";
 
  slideHowToUse.addText("How to use this report", {
    x: 0.3, y: 0.1, w: 9.5, h: 0.9,
    fontFace: "Aptos", fontSize: 40, color: "FFFFFF", bold: false, margin: 0,
  });
 
  const htuHeaders: string[] = [
    "What is this report?",
    "How is it calculated?",
    "What should you do with\nthe insights?",
    "What will happen with\nthe results?",
  ];
 
  htuHeaders.forEach((text, i) => {
    slideHowToUse.addShape(pptx.ShapeType.roundRect, {
      x: HTU_CARD_X[i], y: HTU_HEADER_Y, w: HTU_CARD_W, h: HTU_HEADER_H,
      fill: { color: HTU_HEADER_BG }, line: { color: HTU_HEADER_BG, pt: 0 }, rectRadius: 0.06,
    });
    slideHowToUse.addText(text, {
      x: HTU_CARD_X[i] + 0.13, y: HTU_HEADER_Y,
      w: HTU_CARD_W - 0.26, h: HTU_HEADER_H,
      fontFace: "Aptos", fontSize: 13, color: "FFFFFF",
      bold: false, valign: "middle", margin: 0, breakLine: true,
    });
  });
 
  const htuBodies: string[] = [
    "This report shows how your colleagues experience your leadership day to day.\n\nIt is based on what they see you do — not on opinions about who you are.\n\nUse it to reflect, plan your development and guide coaching conversations.",
    "Your colleagues rated how often they see each leadership behavior in practice.\n\nFor every micro-behavior, we average their answers into a single percentage score based on the frequency observed.\n\nHigher scores are better.",
    "Read through the report and notice the behaviors people see less often.\n\nPick two or three to work on.\n\nTurn them into a personal development plan.",
    "This report is for development purposes only. It will not be used for selection, promotion, performance management or compensation decisions.\n\nDistribution will be limited to you and your appropriate HR contact. It will not be shared with your line manager unless you choose to do so.",
  ];
 
  htuBodies.forEach((text, i) => {
    slideHowToUse.addShape(pptx.ShapeType.roundRect, {
      x: HTU_CARD_X[i], y: HTU_BODY_Y, w: HTU_CARD_W, h: HTU_BODY_H,
      fill: { color: HTU_BODY_BG }, line: { color: HTU_BODY_BG, pt: 0 }, rectRadius: 0.06,
    });
    slideHowToUse.addText(text, {
      x: HTU_CARD_X[i] + 0.15, y: HTU_BODY_Y + 0.15,
      w: HTU_CARD_W - 0.3, h: HTU_BODY_H - 0.3,
      fontFace: "Aptos", fontSize: 12, color: "111111",
      valign: "top", margin: 0, breakLine: true,
    });
  });
 
  slideHowToUse.addShape(pptx.ShapeType.rect, {
    x: 0.3, y: HTU_BANNER_Y, w: 12.73, h: HTU_BANNER_H,
    fill: { color: "555555" }, line: { color: "555555", pt: 0 },
  });
  slideHowToUse.addText(
    "Consider this report to contain insights on how you are experienced — and not a judgement on who you are.",
    { x: 0.3, y: HTU_BANNER_Y, w: 12.73, h: HTU_BANNER_H, fontFace: "Aptos", fontSize: 14, color: "FFFFFF", align: "center", valign: "middle", margin: 0 }
  );

  // ----------------------------------------------------------

  
  //////////////////////////////////////////////////////
  // SLIDE 1 — OVERALL RESULTS
  //////////////////////////////////////////////////////

  const slide1 = pptx.addSlide();
  slide1.background = { color: COLORS.bg };

  slide1.addText("Overall results", {
    x: 0.38, y: 0.22, w: 4.5, h: 0.5,
    fontFace: "Segoe UI", fontSize: 30, color: "111111", margin: 0,
  });

  slide1.addText(
    "How consistently micro-behaviors are seen in practice by your respondents.",
    {
      x: 0.4, y: 0.88, w: 7, h: 0.2,
      fontFace: "Segoe UI", fontSize: 11, color: "444444", margin: 0,
    }
  );

  // Legend
  const legends = [
    { color: COLORS.green, text: "Consistently observed (>85%)", x: 6.9 },
    { color: COLORS.orange, text: "Moderately observed (70-85%)", x: 9.0 },
    { color: COLORS.red, text: "Inconsistently observed (<70%)", x: 11.25 },
  ];

  legends.forEach((l) => {
    slide1.addShape(pptx.ShapeType.rect, {
      x: l.x, y: 0.22, w: 0.18, h: 0.18,
      fill: { color: l.color }, line: { color: l.color },
    });
    slide1.addText(l.text, {
      x: l.x + 0.24, y: 0.18, w: 1.8, h: 0.3,
      fontFace: "Segoe UI", fontSize: 8.5, color: COLORS.text, margin: 0,
    });
  });

  let leftY = 1.45;
  let rightY = 1.45;

  analytics.forEach((item, index) => {
    const isLeft = index < 10;
    const currentY = isLeft ? leftY : rightY;
    const textX = isLeft ? 0.38 : 6.65;
    const barX = isLeft ? 3.82 : 10.08;
    const scoreX = isLeft ? 5.15 : 11.42;

    slide1.addText(item.question, {
      x: textX, y: currentY, w: 3.1, h: 0.35,
      fontFace: "Segoe UI", fontSize: 10, color: COLORS.text,
      fit: "shrink", margin: 0, breakLine: false,
    });

  slide1.addShape(pptx.ShapeType.line, {
  x: barX, y: currentY - 0.02, w: 0.001, h: 0.38,  // ← not w: 0
  line: { color: "8E8E8E", pt: 1 },
});

    const BAR_MAX_W = 1.15;
    const barFillW = (item.rawAverage / 100) * BAR_MAX_W;

    slide1.addShape(pptx.ShapeType.rect, {
      x: barX, y: currentY + 0.1,
      w: barFillW, h: 0.16,
      fill: { color: getScoreColor(item.roundedScore) },
      line: { color: getScoreColor(item.roundedScore) },
    });

    slide1.addText(`${item.roundedScore}%`, {
      x: scoreX, y: currentY + 0.02, w: 0.5, h: 0.2,
      fontFace: "Segoe UI", fontSize: 14,
      bold: false, color: COLORS.text, margin: 0,
    });

    if (isLeft) leftY += 0.48; else rightY += 0.48;
  });

  //////////////////////////////////////////////////////
  // SLIDE 2 — DISTRIBUTION
  //////////////////////////////////////////////////////

  const slide2 = pptx.addSlide();
  slide2.background = { color: COLORS.bg };

  slide2.addText("Results distribution", {
    x: 0.15, y: 0.2, w: 4, h: 0.5,
    fontFace: "Segoe UI", fontSize: 30, color: "111111", margin: 0,
  });

  slide2.addShape(pptx.ShapeType.rect, {
    x: 0.05, y: 0.98, w: 11.35, h: 0.34,
    fill: { color: COLORS.grayHeader }, line: { color: COLORS.grayHeader },
  });

  slide2.addText("Results", {
    x: 0.22, y: 1.06, fontFace: "Segoe UI", fontSize: 8.5, bold: true, margin: 0,
  });

  slide2.addText("Distribution", {
    x: 5.1, y: 1.06, fontFace: "Segoe UI", fontSize: 8.5, bold: true, margin: 0,
  });

  const headers = [
    { label: "Rarely", x: 7.02, color: COLORS.darkRed },
    { label: "Sometimes", x: 7.97, color: COLORS.brightRed },
    { label: "Often", x: 8.92, color: COLORS.beige },
    { label: "Always", x: 9.87, color: COLORS.green },
  ];

  headers.forEach((h) => {
    slide2.addShape(pptx.ShapeType.rect, {
      x: h.x, y: 0.98, w: 0.95, h: 0.34,
      fill: { color: h.color }, line: { color: h.color },
    });
    slide2.addText(h.label, {
      x: h.x, y: 1.06, w: 0.95, h: 0.1, align: "center",
      fontFace: "Segoe UI", fontSize: 8, bold: true, color: "FFFFFF", margin: 0,
    });
  });

  let y = START_Y;

  analytics.forEach((item) => {
    slide2.addShape(pptx.ShapeType.line, {
      x: 0.05, y: y + ROW_HEIGHT, w: 11.35, h: 0,
      line: { color: COLORS.line, pt: 0.5 },
    });

    slide2.addText(item.question, {
      x: QUESTION_X, y: y + 0.02, w: QUESTION_W, h: ROW_HEIGHT,
      fontFace: "Segoe UI", fontSize: 8, color: COLORS.text,
      fit: "shrink", valign: "middle", margin: 0,
    });

    const total =
      item.distribution.Rarely +
      item.distribution.Sometimes +
      item.distribution.Often +
      item.distribution.Always;

    slide2.addShape(pptx.ShapeType.rect, {
      x: BAR_X, y: y + 0.08, w: BAR_W, h: 0.16,
      fill: { color: "EEEEEE" }, line: { color: "EEEEEE" },
    });

    const segments = [
      { value: item.distribution.Rarely, color: COLORS.darkRed },
      { value: item.distribution.Sometimes, color: COLORS.brightRed },
      { value: item.distribution.Often, color: COLORS.beige },
      { value: item.distribution.Always, color: COLORS.green },
    ];

    let segmentX = BAR_X;
    segments.forEach((s) => {
      const width = total > 0 ? (s.value / total) * BAR_W : 0;
      slide2.addShape(pptx.ShapeType.rect, {
        x: segmentX, y: y + 0.08, w: width, h: 0.16,
        fill: { color: s.color }, line: { color: s.color },
      });
      segmentX += width;
    });

    const bgColumns = [
      { x: 7.02, color: COLORS.lightPink },
      { x: 7.97, color: COLORS.lightPink },
      { x: 8.92, color: COLORS.lightBeige },
      { x: 9.87, color: COLORS.lightGreen },
    ];

    bgColumns.forEach((c) => {
      slide2.addShape(pptx.ShapeType.rect, {
        x: c.x, y, w: 0.95, h: ROW_HEIGHT,
        fill: { color: c.color }, line: { color: c.color },
      });
    });

    const counts = [
      item.distribution.Rarely,
      item.distribution.Sometimes,
      item.distribution.Often,
      item.distribution.Always,
    ];

    counts.forEach((count, idx) => {
      slide2.addText(`${count}`, {
        x: COUNT_START_X + idx * COUNT_W, y: y + 0.09, w: COUNT_W, h: 0.1,
        align: "center", fontFace: "Segoe UI", fontSize: 8.5,
        color: COLORS.text, margin: 0,
      });
    });

    y += ROW_HEIGHT;
  });

  //////////////////////////////////////////////////////
  // SLIDE 3 — INSUFFICIENT EXPOSURES
  //////////////////////////////////////////////////////

  const slide3 = pptx.addSlide();
  slide3.background = { color: "F2F2F2" };

  slide3.addText("Insufficient Exposures", {
    x: 0.25, y: 0.05, w: 5.8, h: 0.55,
    fontFace: "Segoe UI", fontSize: 28, color: "1A1A1A", bold: false, margin: 0,
  });

  slide3.addShape(pptx.ShapeType.rect, {
    x: 0.15, y: 0.82, w: 7.1, h: 0.3,
    fill: { color: "D0D0D0" }, line: { color: "D0D0D0" },
  });

  slide3.addText("Results", {
    x: 0.22, y: 0.89, w: 1, h: 0.1,
    fontFace: "Segoe UI", fontSize: 8.5, bold: true, color: "111111", margin: 0,
  });

  slide3.addText("Count of Insufficient Exposure", {
    x: 4.15, y: 0.89, w: 2.4, h: 0.1,
    fontFace: "Segoe UI", fontSize: 8.5, bold: true, color: "111111", margin: 0,
  });

  let tableY = 1.12;

  analytics.forEach((item) => {
    const count = item.distribution["Insufficient Exposure"] || 0;

  // slide3 — fix zero-height separator lines
slide3.addShape(pptx.ShapeType.line, {
  x: 0.15, y: tableY + 0.27, w: 7.1, h: 0.001,  // ← not h: 0
  line: { color: "C8C8C8", pt: 0.5 },
});

    slide3.addText(item.question, {
      x: 0.22, y: tableY + 0.03, w: 3.8, h: 0.2,
      fontFace: "Segoe UI", fontSize: 8.2, color: "222222", margin: 0, fit: "shrink",
    });

    slide3.addText(`${count}`, {
      x: 4.2, y: tableY + 0.03, w: 0.2, h: 0.1,
      fontFace: "Consolas", fontSize: 9, color: "222222", margin: 0,
    });

    if (count > 0) {
      const barWidth = count * 0.32;
      slide3.addShape(pptx.ShapeType.rect, {
        x: 4.18, y: tableY + 0.16, w: barWidth, h: 0.1,
        fill: { color: "1D6FD8" }, line: { color: "1D6FD8" },
      });
      slide3.addText(`${count}`, {
        x: 4.18 + barWidth + 0.04, y: tableY + 0.11, w: 0.2, h: 0.1,
        fontFace: "Segoe UI", fontSize: 9, color: "444444", margin: 0,
      });
    }

    tableY += 0.29;
  });

  slide3.addShape(pptx.ShapeType.rect, {
    x: 7.8, y: 3.95, w: 4.35, h: 0.3,
    fill: { color: "A5A5A5" }, line: { color: "A5A5A5" },
  });

  slide3.addText("Note", {
    x: 7.92, y: 4.02, w: 0.5, h: 0.1,
    fontFace: "Segoe UI", fontSize: 9, bold: true, color: "FFFFFF", margin: 0,
  });

  slide3.addShape(pptx.ShapeType.rect, {
    x: 7.8, y: 4.25, w: 4.35, h: 2.65,
    fill: { color: "EFEFEF" }, line: { color: "EFEFEF" },
  });

  slide3.addText(
    `Insufficient Exposure is what a rater selects when they feel they have not seen sufficient evidence from you on this behavior to provide a rating. These answers are set aside and they do not lower your score.\n\nA high count can itself be a signal as it can mean a behavior is not be visible to the people around you, even if you feel you are demonstrating it.\n\nIf you have areas with high insufficient exposure, we recommend you consider why you are receiving this feedback and what you could do to address this.`,
    {
      x: 7.95, y: 4.38, w: 3.95, h: 2.25,
      fontFace: "Segoe UI", fontSize: 8.8, color: "222222",
      valign: "top", breakLine: true, fit: "shrink", margin: 0,
    }
  );

  //////////////////////////////////////////////////////
  // SLIDE 4 — RADAR BENCHMARK
  // Renders only if categorised questions exist
  //////////////////////////////////////////////////////

  if (leadingSelf.length > 0 || leadingOthers.length > 0) {

    const slide4 = pptx.addSlide();
    slide4.background = { color: "FFFFFF" };

    slide4.addText("Benchmarking", {
      x: 0.38, y: 0.15, w: 6, h: 0.5,
      fontFace: "Segoe UI Semibold", fontSize: 32, color: "111111", margin: 0,
    });

    const PANEL_W = 6.2;
    const PANEL_H = 5.2;
    const PANEL_Y = 0.9;
    const HEADER_H = 0.35;
    const COLOR_HEADER = "1C2B4B";
    const COLOR_CONTENT = "F7F9FB";

    // LEADING SELF PANEL
    if (leadingSelf.length > 0) {
      slide4.addShape(pptx.ShapeType.rect, {
        x: 0.35, y: PANEL_Y, w: PANEL_W, h: HEADER_H,
        fill: { color: COLOR_HEADER },
      });
      slide4.addText("Leading Self", {
        x: 0.45, y: PANEL_Y, w: 4, h: HEADER_H,
        fontFace: "Segoe UI", fontSize: 11, bold: true, color: "FFFFFF", margin: 0,
      });
      slide4.addShape(pptx.ShapeType.rect, {
        x: 0.35, y: PANEL_Y + HEADER_H, w: PANEL_W, h: PANEL_H - HEADER_H,
        fill: { color: COLOR_CONTENT },
      });

      slide4.addChart(
        pptx.ChartType.radar,
        [
          {
            name: "Participant",
            labels: selfLeadingSelf.map((q) => q.question),
            values: selfLeadingSelf.map((q, idx) => {
              return q.validResponses > 0 ? q.rawAverage : leadingSelf[idx].rawAverage;
            }),
          },
          {
            name: "Enterprise Avg",
            labels: enterpriseLeadingSelf.map((q) => q.question),
            values: enterpriseLeadingSelf.map((q) => q.rawAverage),
          },
        ],
        {
          x: 0.45, y: PANEL_Y + HEADER_H + 0.1, w: 6.0, h: 4.6,
          radarStyle: "marker",
          lineDataSymbol: "circle",
          lineDataSymbolSize: 4,
          chartColors: ["1D6FD8", "00897B"],
          lineSmooth: false,
          showLegend: false,
          showTitle: false,
          valAxisMaxVal: 100,
          valAxisMinVal: 0,
          catAxisLabelFontSize: 6.5,
          valAxisLabelFontSize: 6,
        }
      );
    }

    // LEADING OTHERS PANEL
    if (leadingOthers.length > 0) {
      const xOffset = 6.8;
      slide4.addShape(pptx.ShapeType.rect, {
        x: xOffset, y: PANEL_Y, w: PANEL_W, h: HEADER_H,
        fill: { color: COLOR_HEADER },
      });
      slide4.addText("Leading Others", {
        x: xOffset + 0.08, y: PANEL_Y, w: 4, h: HEADER_H,
        fontFace: "Segoe UI", fontSize: 11, bold: true, color: "FFFFFF", margin: 0,
      });
      slide4.addShape(pptx.ShapeType.rect, {
        x: xOffset, y: PANEL_Y + HEADER_H, w: PANEL_W, h: PANEL_H - HEADER_H,
        fill: { color: COLOR_CONTENT },
      });

      slide4.addChart(
        pptx.ChartType.radar,
        [
          {
            name: "Participant",
            labels: selfLeadingOthers.map((q) => q.question),
            values: selfLeadingOthers.map((q, idx) => {
              return q.validResponses > 0 ? q.rawAverage : leadingOthers[idx].rawAverage;
            }),
          },
          {
            name: "Enterprise Avg",
            labels: enterpriseLeadingOthers.map((q) => q.question),
            values: enterpriseLeadingOthers.map((q) => q.rawAverage),
          },
        ],
        {
          x: xOffset + 0.1, y: PANEL_Y + HEADER_H + 0.1, w: 6.0, h: 4.6,
          radarStyle: "marker",
          lineDataSymbol: "circle",
          lineDataSymbolSize: 4,
          chartColors: ["1D6FD8", "00897B"],
          lineSmooth: false,
          showLegend: false,
          showTitle: false,
          valAxisMaxVal: 100,
          valAxisMinVal: 0,
          catAxisLabelFontSize: 6.5,
          valAxisLabelFontSize: 6,
        }
      );
    }

    // LEGEND AT BOTTOM
    const LEGEND_BOTTOM_Y = 6.4;
    slide4.addText("Legend:", { x: 4.5, y: LEGEND_BOTTOM_Y, w: 0.8, h: 0.2, fontFace: "Segoe UI", fontSize: 9, bold: true, align: "right" });

    slide4.addShape(pptx.ShapeType.line, { x: 5.4, y: LEGEND_BOTTOM_Y + 0.1, w: 0.25, h: 0, line: { color: "1D6FD8", pt: 1.5 } });
    slide4.addText("Participant", { x: 5.7, y: LEGEND_BOTTOM_Y, w: 1, h: 0.2, fontFace: "Segoe UI", fontSize: 9 });

    slide4.addShape(pptx.ShapeType.line, { x: 6.8, y: LEGEND_BOTTOM_Y + 0.1, w: 0.25, h: 0, line: { color: "00897B", pt: 1.5 } });
    slide4.addText("Enterprise Average", { x: 7.1, y: LEGEND_BOTTOM_Y, w: 1.5, h: 0.2, fontFace: "Segoe UI", fontSize: 9 });
  }

  //////////////////////////////////////////////////////
  // SLIDE 5 — SELF / OTHER
  //////////////////////////////////////////////////////

  const slide5 = pptx.addSlide();
  slide5.background = { color: "FFFFFF" };

  slide5.addText("Self / other", {
    x: 0.38, y: 0.15, w: 8, h: 0.5,
    fontFace: "Segoe UI Semibold", fontSize: 32, color: "111111", margin: 0,
  });

  slide5.addText("This page shows where your self-view differs from how colleagues experience your micro-behaviors.", {
    x: 0.38, y: 0.7, w: 8, h: 0.3,
    fontFace: "Segoe UI", fontSize: 11, color: "555555", margin: 0,
  });

  // Three cards at top
  const cardW = 2.8;
  const cardH = 1.2;
  const cardStartY = 1.15;
  const cardGap = 0.25;
  const startX = 0.4;

  const cards = [
    {
      header: "Colleagues see less often",
      subheader: "Potential blind spots",
      bullets: [
        "Develops others for long-term impact",
        "Connects work to vision and strategy",
        "Sets high standards for performance",
      ],
    },
    {
      header: "Broadly aligned",
      subheader: "Self-view and colleague feedback are similar",
      bullets: [
        "Questions assumptions and ways of working",
        "Adapts based on feedback",
        "Uses external signals to guide decisions",
      ],
    },
    {
      header: "Colleagues see more often",
      subheader: "Potential strengths",
      bullets: [
        "Seeks input from affected stakeholders before moving forward",
        "Sets direction and priorities",
        "Facilitates collective momentum across teams",
      ],
    },
  ];

  cards.forEach((card, idx) => {
    const cardX = startX + idx * (cardW + cardGap);

    // Blue header
    slide5.addShape(pptx.ShapeType.rect, {
      x: cardX,
      y: cardStartY,
      w: cardW,
      h: 0.35,
      fill: { color: "2563EB" },
      line: { color: "2563EB" },
    });

    slide5.addText(card.header, {
      x: cardX + 0.1,
      y: cardStartY + 0.05,
      w: cardW - 0.2,
      h: 0.25,
      fontFace: "Segoe UI",
      fontSize: 11,
      bold: true,
      color: "FFFFFF",
      valign: "middle",
      margin: 0,
    });

    // Subheader
    slide5.addText(card.subheader, {
      x: cardX + 0.1,
      y: cardStartY + 0.4,
      w: cardW - 0.2,
      h: 0.25,
      fontFace: "Segoe UI",
      fontSize: 9,
      bold: true,
      color: "1D6FD8",
      valign: "top",
      margin: 0,
    });

    // Bullets
    const bulletText = card.bullets.map((b) => "• " + b).join("\n");
    slide5.addText(bulletText, {
      x: cardX + 0.1,
      y: cardStartY + 0.7,
      w: cardW - 0.2,
      h: 1.5,
      fontFace: "Segoe UI",
      fontSize: 8.5,
      color: "333333",
      valign: "top",
      breakLine: true,
      margin: 0,
    });
  });

  //////////////////////////////////////////////////////
  // SLIDE 6 — QUALITATIVE FEEDBACK
  //////////////////////////////////////////////////////

  const slide6 = pptx.addSlide();
  slide6.background = { color: "FFFFFF" };

  slide6.addText("Qualitative feedback", {
    x: 0.38, y: 0.15, w: 8, h: 0.5,
    fontFace: "Segoe UI Semibold", fontSize: 32, color: "111111", margin: 0,
  });

  slide6.addText("This page summarises key themes from the qualitative feedback your respondents provided.", {
    x: 0.38, y: 0.7, w: 8, h: 0.3,
    fontFace: "Segoe UI", fontSize: 11, color: "555555", margin: 0,
  });

  // Two cards side-by-side
  const qualCardW = 4.4;
  const qualCardH = 4.8;
  const qualCardStartY = 1.15;
  const qualCardGap = 0.4;
  const qualStartX = 0.4;

  const qualCards = [
    {
      header: "Perceived areas of strength",
      content: "You are perceived as operating across boundaries with credibility - colleagues describe this as your most distinctive contribution.\n\nSeen as bringing the right stakeholders in early, which is felt to reduce rework and accelerate decisions.\n\nPerceived as trusted across functions, including on difficult topics.\n\nExperienced as setting direction with clarity - colleagues note they leave meetings clear on what was agreed and who owns what.\n\nDescribed as thinking at the system level, not only within your own area.",
    },
    {
      header: "Perceived areas of improvement",
      content: "Your support is perceived as genuine but reactive. Development of others appears to be the clearest opportunity for improvement.\n\nFeedback is felt to follow a request rather than being offered proactively.\n\nRecognised potential is not always seen as translated into a clear stretch or development plan.\n\nPriorities are perceived as set clearly, but follow-through on standards is reported as less visible.\n\nLow-value work is seen as continuing past its useful life.",
    },
  ];

  qualCards.forEach((card, idx) => {
    const cardX = qualStartX + idx * (qualCardW + qualCardGap);

    // Blue header
    slide6.addShape(pptx.ShapeType.rect, {
      x: cardX,
      y: qualCardStartY,
      w: qualCardW,
      h: 0.4,
      fill: { color: "2563EB" },
      line: { color: "2563EB" },
    });

    slide6.addText(card.header, {
      x: cardX + 0.15,
      y: qualCardStartY + 0.08,
      w: qualCardW - 0.3,
      h: 0.25,
      fontFace: "Segoe UI",
      fontSize: 13,
      bold: true,
      color: "FFFFFF",
      valign: "middle",
      margin: 0,
    });

    // Content
    slide6.addText(card.content, {
      x: cardX + 0.15,
      y: qualCardStartY + 0.5,
      w: qualCardW - 0.3,
      h: qualCardH - 0.55,
      fontFace: "Segoe UI",
      fontSize: 10,
      color: "333333",
      valign: "top",
      breakLine: true,
      margin: 0,
    });
  });

  //////////////////////////////////////////////////////
  // SLIDE 7 — COACHING QUADRANT (commented out for now)
  //////////////////////////////////////////////////////
  // slide6.background = { color: "FFFFFF" };

  // slide6.addText("360 results: Coaching quadrant", {
  //   x: 0.38, y: 0.15, w: 8, h: 0.6,
  //   fontFace: "Segoe UI Semibold", fontSize: 32, color: "1C2B4B", margin: 0,
  // });

  // const THRESHOLD = 70;
  // const quadrants = {
  //   underRecognized: [] as string[],
  //   sharedDevelopment: [] as string[],
  //   sharedStrengths: [] as string[],
  //   blindSpots: [] as string[]
  // };

  // analytics.forEach((item, idx) => {
  //   const s = selfAnalytics[idx]?.rawAverage || 0;
  //   const o = othersAnalytics[idx]?.rawAverage || 0;
  //   const qText = item.question;

  //   if (s >= THRESHOLD && o >= THRESHOLD) quadrants.sharedStrengths.push(qText);
  //   else if (s >= THRESHOLD && o < THRESHOLD) quadrants.blindSpots.push(qText);
  //   else if (s < THRESHOLD && o >= THRESHOLD) quadrants.underRecognized.push(qText);
  //   else quadrants.sharedDevelopment.push(qText);
  // });

  // const BOX_W = 4.3;
  // const BOX_H = 2.4;
  // const GAP = 0.05;
  // const QUADRANT_START_X = 0.5;
  // const QUADRANT_START_Y = 1.3;

  // const renderQuadrant = (title: string, desc: string, items: string[], x: number, y: number, bgColor: string) => {
  //   slide6.addShape(pptx.ShapeType.rect, {
  //     x, y, w: BOX_W, h: BOX_H,
  //     fill: { color: bgColor }, line: { color: "EEEEEE", pt: 0.5 }
  //   });
  //   slide6.addText(title, {
  //     x: x + 0.1, y: y + 0.1, w: BOX_W - 0.2, h: 0.25,
  //     fontFace: "Segoe UI", fontSize: 13, bold: true, color: "333333"
  //   });
  //   slide6.addText(desc, {
  //     x: x + 0.1, y: y + 0.35, w: BOX_W - 0.2, h: 0.35,
  //     fontFace: "Segoe UI", fontSize: 11, bold: true, color: "1D6FD8"
  //   });

  //   const bulletText = items.slice(0, 4).map(it => `• ${it}`).join("\n");
  //   slide6.addText(bulletText, {
  //     x: x + 0.1, y: y + 0.8, w: BOX_W - 0.2, h: 1.5,
  //     fontFace: "Segoe UI", fontSize: 9.5, color: "333333", valign: "top"
  //   });
  // };

  // renderQuadrant(
  //   "Under-recognized strengths",
  //   "Colleagues see these more consistently than you do.",
  //   quadrants.underRecognized,
  //   QUADRANT_START_X, QUADRANT_START_Y, "EFEFEF"
  // );
  // renderQuadrant(
  //   "Shared development areas",
  //   "You and colleagues both see these as less consistent.",
  //   quadrants.sharedDevelopment,
  //   QUADRANT_START_X + BOX_W + GAP, QUADRANT_START_Y, "FFFFFF"
  // );
  // renderQuadrant(
  //   "Shared strengths",
  //   "You think you show these often, and colleagues broadly agree.",
  //   quadrants.sharedStrengths,
  //   QUADRANT_START_X, QUADRANT_START_Y + BOX_H + GAP, "FFFFFF"
  // );
  // renderQuadrant(
  //   "Potential blind spots",
  //   "You think you show these often, but colleagues do not see them consistently.",
  //   quadrants.blindSpots,
  //   QUADRANT_START_X + BOX_W + GAP, QUADRANT_START_Y + BOX_H + GAP, "EFEFEF"
  // );

  // const labelStyle = { fontFace: "Segoe UI", fontSize: 10, color: "333333", italic: true };
  // slide6.addText("You consider consistent", { x: 0.1, y: QUADRANT_START_Y + BOX_H - 0.2, w: 0.3, h: 2, rotate: 270, ...labelStyle });
  // slide6.addText("You don't consider consistent", { x: 0.1, y: QUADRANT_START_Y - 0.2, w: 0.3, h: 2, rotate: 270, ...labelStyle });
  // slide6.addText("Colleagues see it consistently", { x: QUADRANT_START_X, y: QUADRANT_START_Y + (BOX_H * 2) + 0.1, w: BOX_W, h: 0.3, align: "center", ...labelStyle });
  // slide6.addText("Colleagues don't see it consistently", { x: QUADRANT_START_X + BOX_W, y: QUADRANT_START_Y + (BOX_H * 2) + 0.1, w: BOX_W, h: 0.3, align: "center", ...labelStyle });

  // slide6.addShape(pptx.ShapeType.rect, {
  //   x: 7.5, y: 6.8, w: 1.8, h: 0.3,
  //   fill: { color: "FFFFFF" }, line: { color: "B00000", pt: 1 },
  // });
  // slide6.addText(`${form.responses[0]?.participant?.fullName || "PARTICIPANT"} & ${new Date().toLocaleDateString()}`, {
  //   x: 7.55, y: 6.85, w: 1.7, h: 0.2,
  //   fontFace: "Segoe UI", fontSize: 7, color: "B00000", bold: true, align: "center",
  // });

  //////////////////////////////////////////////////////
  // EXPORT
  //////////////////////////////////////////////////////

  const pptxData = await pptx.write({ outputType: "nodebuffer" });

  const uint8 =
    pptxData instanceof Uint8Array
      ? pptxData
      : new Uint8Array(pptxData as ArrayBuffer);

  const arrayBuffer = uint8.buffer.slice(
    uint8.byteOffset,
    uint8.byteOffset + uint8.byteLength
  ) as ArrayBuffer;

  return new Response(arrayBuffer, {
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "Content-Disposition":
        `attachment; filename=report-${formId}.pptx`,
    },
  });
}