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
          participant: true, // Added
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

  // SELF = responses where respondent email == participant email
  const selfResponses = filteredResponses.filter(
    (r: any) => !!r.participant?.email && r.email.toLowerCase() === r.participant.email.toLowerCase()
  );

  // OTHERS = everything else
  const othersResponses = filteredResponses.filter(
    (r: any) => !r.participant?.email || r.email.toLowerCase() !== r.participant.email.toLowerCase()
  );

  //////////////////////////////////////////////////////
  // COMPUTE ANALYTICS
  //////////////////////////////////////////////////////

  const analytics = computeAnalytics(form.questions, filteredResponses);
  const selfAnalytics = computeAnalytics(form.questions, selfResponses);
  const othersAnalytics = computeAnalytics(form.questions, othersResponses);
  const enterpriseAnalytics = computeAnalytics(form.questions, allResponses);

  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Antigravity";
  pptx.subject = "360 Feedback";
  pptx.company = "360 Feedback Report";

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

  const slideHowCreated = pptx.addSlide();
  slideHowCreated.background = { color: "111111" }; // Dark theme matched to screenshot

  // TITLES
  slideHowCreated.addText("How this report was\ncreated", {
    x: 0.5, y: 0.4, w: 8, h: 1.2,
    fontFace: "Segoe UI", fontSize: 50, bold: false, color: "FFFFFF", margin: 0,
  });


  // Blue banner notification (top right)
  slideHowCreated.addShape(pptx.ShapeType.roundRect, {
    x: 7.7, y: 0.5, w: 2.8, h: 0.65,
    fill: { color: "1E56DB" }, line: { color: "1E56DB" },
    rectRadius: 0.12
  });
  slideHowCreated.addText("Micro-behaviors not yet\nfinalized", {
    x: 7.7, y: 0.5, w: 2.8, h: 0.65,
    fontFace: "Segoe UI", fontSize: 9.5, bold: true, color: "FFFFFF",
    align: "center", valign: "middle", margin: 0,
  });

  // LEFT SECTION - BACKGROUND
  slideHowCreated.addText("Background", {
    x: 0.5, y: 1.8, w: 2.5, h: 0.4,
    fontFace: "Segoe UI", fontSize: 16, bold: true, color: "FFFFFF", margin: 0,
  });
  slideHowCreated.addText(
    "You nominated colleagues across manager, peer, direct report and broader stakeholder groups to give a balanced view of your leadership.\n\nEach rater assessed you against the micro-behaviors on the right. Micro-behaviors are deliberately observable and specific interpretations of the leadership framework.",
    { x: 0.5, y: 2.3, w: 2.8, h: 3.5, fontFace: "Segoe UI", fontSize: 11, color: "FFFFFF", valign: "top", breakLine: true, margin: 0 }
  );


  // VERTICAL DIVIDER LINE
  slideHowCreated.addShape(pptx.ShapeType.line, {
    x: 3.7, y: 1.8, w: 0.001, h: 5.2,
    line: { color: "444444", pt: 1.2 },
  });

  // BLUE CHEVRON DECORATION
  slideHowCreated.addShape(pptx.ShapeType.chevron, {
    x: 3.55, y: 4.15, w: 0.3, h: 0.4,
    fill: { color: "1E56DB" }, line: { color: "1E56DB" },
  });

  // MIDDLE & RIGHT SECTIONS - MICRO-BEHAVIORS
  slideHowCreated.addText("Micro-Behaviors", {
    x: 4.3, y: 1.8, w: 6.2, h: 0.4,
    fontFace: "Segoe UI", fontSize: 16, bold: true, color: "FFFFFF", margin: 0,
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

  let behaviorY = 2.3;
  const lineSpacing = 0.42;

  leftBehaviors.forEach((behavior) => {
    slideHowCreated.addText(behavior, {
      x: 4.3, y: behaviorY, w: 3.8, h: 0.35,
      fontFace: "Segoe UI", fontSize: 10.5, color: "82B7E8",
      valign: "top", breakLine: true, margin: 0,
    });
    behaviorY += lineSpacing;
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

  let rightBehaviorY = 2.3;
  rightBehaviors.forEach((behavior) => {
    slideHowCreated.addText(behavior, {
      x: 8.5, y: rightBehaviorY, w: 4.0, h: 0.35,
      fontFace: "Segoe UI", fontSize: 10.5, color: "82B7E8",
      valign: "top", breakLine: true, margin: 0,
    });
    rightBehaviorY += lineSpacing;
  });

  // ---------------------------------------------------------

  const slideHowToUse = pptx.addSlide();
  slideHowToUse.background = { color: "111111" };

  const HTU_CARD_X: number[] = [0.3, 3.52, 6.74, 9.96];
  const HTU_CARD_W: number = 3.07;
  const HTU_HEADER_Y: number = 1.5;
  const HTU_HEADER_H: number = 0.78;
  const HTU_BODY_Y: number = HTU_HEADER_Y + HTU_HEADER_H;
  const HTU_BODY_H: number = 3.9;
  const HTU_BANNER_Y: number = 6.28;
  const HTU_BANNER_H: number = 0.88;
  const HTU_HEADER_BG: string = "2D6DB5";
  const HTU_BODY_BG: string = "E8E8E8";

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

  const SPLIT_INDEX = Math.ceil(analytics.length / 2);
  let leftY = 1.05;
  let rightY = 1.05;
  const Y_GAP = 0.40;

  analytics.forEach((item, index) => {
    const isLeft = index < SPLIT_INDEX;
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
      x: scoreX, y: currentY + 0.02, w: 0.8, h: 0.2, // Widened to 0.8 to prevent wrapping
      fontFace: "Segoe UI", fontSize: 13,
      bold: false, color: COLORS.text, margin: 0,
    });

    if (isLeft) leftY += Y_GAP; else rightY += Y_GAP;
  });

  //////////////////////////////////////////////////////
  // SLIDE 2 — DISTRIBUTION
  //////////////////////////////////////////////////////

  //////////////////////////////////////////////////////
  // SLIDE 2 — RESULTS DISTRIBUTION
  //////////////////////////////////////////////////////

  const slide2 = pptx.addSlide();
  slide2.background = { color: "FFFFFF" };


  slide2.addText("Results distribution", {
    x: 0.38, y: 0.15, w: 6, h: 0.5,
    fontFace: "Segoe UI", fontSize: 32, color: "111111", margin: 0,
  });

  // Blue Notification Box - Centered above Distribution column
  slide2.addShape(pptx.ShapeType.rect, {
    x: 4.8, y: 0.15, w: 2.5, h: 0.35,
    fill: { color: "1E56DB" }
  });
  slide2.addText("Micro-behaviors not yet finalized", {
    x: 4.8, y: 0.15, w: 2.5, h: 0.35,
    fontFace: "Segoe UI", fontSize: 8.5, bold: true, color: "FFFFFF",
    align: "center", valign: "middle", margin: 0,
  });

  // TABLE CONFIG - Following template layout
  const TABLE_X = 0.38;
  const TABLE_Y = 0.75;
  const HEADER_H = 0.36;
  const COL_W = [4.6, 4.6, 0.8, 0.8, 0.8, 0.8];
  const COL_X = [
    TABLE_X,
    TABLE_X + COL_W[0],
    TABLE_X + COL_W[0] + COL_W[1],
    TABLE_X + COL_W[0] + COL_W[1] + 0.8,
    TABLE_X + COL_W[0] + COL_W[1] + 1.6,
    TABLE_X + COL_W[0] + COL_W[1] + 2.4
  ];
  const ROW_H = 0.20;
  const TOTAL_W = 12.4;

  // Header Backgrounds
  // Results & Distribution part (Light Gray)
  slide2.addShape(pptx.ShapeType.rect, { x: COL_X[0], y: TABLE_Y, w: COL_W[0] + COL_W[1], h: HEADER_H, fill: { color: "E7E7E7" } });

  const headers = [
    { text: "Results", x: COL_X[0], w: COL_W[0], align: "left", color: "111111", bold: true },
    { text: "Distribution", x: COL_X[1], w: COL_W[1], align: "center", color: "111111", bold: true },
    { text: "Rarely", x: COL_X[2], w: COL_W[2], bg: "9E0B0F", color: "FFFFFF" },
    { text: "Sometimes", x: COL_X[3], w: COL_W[3], bg: "FF0000", color: "FFFFFF" },
    { text: "Often", x: COL_X[4], w: COL_W[4], bg: "AD9380", color: "FFFFFF" },
    { text: "Always", x: COL_X[5], w: COL_W[5], bg: "008032", color: "FFFFFF" },
  ];

  headers.forEach(h => {
    if (h.bg) {
      slide2.addShape(pptx.ShapeType.rect, { x: h.x, y: TABLE_Y, w: h.w, h: HEADER_H, fill: { color: h.bg } });
    }
    slide2.addText(h.text, {
      x: h.x, y: TABLE_Y, w: h.w, h: HEADER_H,
      fontFace: "Segoe UI", fontSize: h.text === "Sometimes" ? 8 : 9, bold: h.bold || false, color: h.color,
      align: (h.align || "center") as any, valign: "middle", margin: h.align === "left" ? 0.1 : 0
    });
  });

  // COLUMN STRIPES (Very Light)
  const STRIPE_H = analytics.length * ROW_H;
  slide2.addShape(pptx.ShapeType.rect, { x: COL_X[2], y: TABLE_Y + HEADER_H, w: COL_W[2], h: STRIPE_H, fill: { color: "F9EBEB" } });
  slide2.addShape(pptx.ShapeType.rect, { x: COL_X[3], y: TABLE_Y + HEADER_H, w: COL_W[3], h: STRIPE_H, fill: { color: "FFF0F0" } });
  slide2.addShape(pptx.ShapeType.rect, { x: COL_X[4], y: TABLE_Y + HEADER_H, w: COL_W[4], h: STRIPE_H, fill: { color: "F5F5F5" } });
  slide2.addShape(pptx.ShapeType.rect, { x: COL_X[5], y: TABLE_Y + HEADER_H, w: COL_W[5], h: STRIPE_H, fill: { color: "F0F9F0" } });

  const CHART_CENTER_X = COL_X[1] + (COL_W[1] / 2);

  // ROWS
  othersAnalytics.forEach((item, idx) => {
    const y = TABLE_Y + HEADER_H + idx * ROW_H;
    if (y > 7.3) return;

    // Full row line
    slide2.addShape(pptx.ShapeType.line, {
      x: COL_X[0], y: y + ROW_H, w: TOTAL_W, h: 0,
      line: { color: "CCCCCC", pt: 0.5 }
    });

    // Content
    slide2.addText(item.question, {
      x: COL_X[0] + 0.1, y: y, w: COL_W[0] - 0.2, h: ROW_H,
      fontFace: "Segoe UI", fontSize: 8.0, color: "111111", align: "left", valign: "middle", margin: 0,
      fit: "shrink"
    });

    const dist = item.distribution;
    const counts = [dist.Rarely, dist.Sometimes, dist.Often, dist.Always];
    counts.forEach((c, cIdx) => {
      slide2.addText(c.toString(), {
        x: COL_X[2 + cIdx], y: y, w: COL_W[2 + cIdx], h: ROW_H,
        fontFace: "Segoe UI", fontSize: 9, color: "111111", align: "center", valign: "middle", margin: 0
      });
    });

    // Bars - Normalized to stay within column
    const total = (dist.Rarely || 0) + (dist.Sometimes || 0) + (dist.Often || 0) + (dist.Always || 0);
    if (total > 0) {
      // Divergent logic: Rarely/Sometimes go Left, Often/Always go Right
      // Max possible width for one side is COL_W[1] / 2 - padding
      const HALF_BAR_MAX_W = (COL_W[1] / 2) - 0.1; // 2.2 units
      const scale = HALF_BAR_MAX_W / total;

      const wR = (dist.Rarely || 0) * scale;
      const wS = (dist.Sometimes || 0) * scale;
      const wO = (dist.Often || 0) * scale;
      const wA = (dist.Always || 0) * scale;

      const barY = y + (ROW_H * 0.2);
      const barH = ROW_H * 0.6;

      if (wR > 0) slide2.addShape(pptx.ShapeType.rect, { x: CHART_CENTER_X - wS - wR, y: barY, w: wR, h: barH, fill: { color: "9E0B0F" }, line: { pt: 0 } });
      if (wS > 0) slide2.addShape(pptx.ShapeType.rect, { x: CHART_CENTER_X - wS, y: barY, w: wS, h: barH, fill: { color: "FF0000" }, line: { pt: 0 } });
      if (wO > 0) slide2.addShape(pptx.ShapeType.rect, { x: CHART_CENTER_X, y: barY, w: wO, h: barH, fill: { color: "AD9380" }, line: { pt: 0 } });
      if (wA > 0) slide2.addShape(pptx.ShapeType.rect, { x: CHART_CENTER_X + wO, y: barY, w: wA, h: barH, fill: { color: "008032" }, line: { pt: 0 } });
    }
  });

  // Strong center axis
  slide2.addShape(pptx.ShapeType.line, {
    x: CHART_CENTER_X, y: TABLE_Y + HEADER_H, w: 0, h: STRIPE_H,
    line: { color: "111111", pt: 1 }
  });

  // Footer - Box style from template
  slide2.addShape(pptx.ShapeType.rect, {
    x: 10.4, y: 7.3, w: 2.5, h: 0.18,
    fill: { color: "FFFFFF" }, line: { color: "B00000", pt: 1 },
  });
  slide2.addText(`${form.responses[0]?.participant?.fullName || "PARTICIPANT"} & ${new Date().toLocaleDateString()}`, {
    x: 10.4, y: 7.3, w: 2.5, h: 0.18,
    fontFace: "Segoe UI", fontSize: 6.5, color: "B00000", bold: true, align: "center",
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

  let tableY = 1.05;

  const PIX_ROW_H = 0.20;

  // Vertical Separator Line for the table
  slide3.addShape(pptx.ShapeType.line, {
    x: 4.15, y: 0.82, w: 0.001, h: (analytics.length * PIX_ROW_H) + 0.3,
    line: { color: "A0A0A0", pt: 1 }
  });

  analytics.forEach((item) => {
    const count = item.distribution["Insufficient Exposure"] || 0;

    // slide3 — fix zero-height separator lines
    slide3.addShape(pptx.ShapeType.line, {
      x: 0.15, y: tableY + PIX_ROW_H, w: 7.1, h: 0.001,
      line: { color: "C8C8C8", pt: 0.5 },
    });

    slide3.addText(item.question, {
      x: 0.22, y: tableY + 0.01, w: 3.8, h: 0.18,
      fontFace: "Segoe UI", fontSize: 7.5, color: "333333", margin: 0, fit: "shrink",
    });

    // Zero-state center count
    if (count === 0) {
      slide3.addText(`0`, {
        x: 4.2, y: tableY + 0.01, w: 0.2, h: 0.18,
        fontFace: "Segoe UI", fontSize: 8, color: "333333", margin: 0,
        align: "center"
      });
    }

    if (count > 0) {
      const barWidth = count * 0.32;
      slide3.addShape(pptx.ShapeType.rect, {
        x: 4.18, y: tableY + 0.10, w: barWidth, h: 0.04,
        fill: { color: "0055D4" }, line: { pt: 0 },
      });
      slide3.addText(`${count}`, {
        x: 4.18 + barWidth + 0.04, y: tableY + 0.01, w: 0.3, h: 0.18,
        fontFace: "Segoe UI", fontSize: 7.5, color: "333333", margin: 0,
      });
    }

    tableY += PIX_ROW_H;
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
  // SLIDE 5 — SELF / OTHER
  //////////////////////////////////////////////////////

  const slide5 = pptx.addSlide();
  slide5.background = { color: "FFFFFF" };


  slide5.addText("Self / other", {
    x: 0.38, y: 0.15, w: 8, h: 0.5,
    fontFace: "Segoe UI Semibold", fontSize: 40, color: "111111", margin: 0,
  });

  slide5.addText("This page shows where your self-view differs from how colleagues experience your micro-behaviors.", {
    x: 0.38, y: 0.75, w: 10, h: 0.3,
    fontFace: "Segoe UI", fontSize: 11, color: "111111", margin: 0,
  });

  const cardW = 3.9;
  const cardStartY = 1.4;
  const cardGap = 0.25;
  const startX = 0.4;

  const cards = [
    {
      header: "Colleagues see less\noften",
      subheader: "Potential blind spots",
      bullets: [
        "Develops others for long-term impact",
        "Connects work to vision and strategy",
        "Sets high standards for performance",
      ],
    },
    {
      header: "Broadly aligned",
      subheader: "Self-view and colleague\nfeedback are similar",
      bullets: [
        "Questions assumptions and\nways of working",
        "Adapts based on feedback",
        "Uses external signals to\nguide decisions",
      ],
    },
    {
      header: "Colleagues see more\noften",
      subheader: "Potential strengths",
      bullets: [
        "Seeks input from affected\nstakeholders before moving\nforward",
        "Sets direction and priorities",
        "Facilitates collective\nmomentum across teams",
      ],
    },
  ];

  cards.forEach((card, idx) => {
    const cardX = startX + idx * (cardW + cardGap);
    const HEADER_H = 0.75;
    const BODY_H = 4.3;

    // Header (Blue)
    slide5.addShape(pptx.ShapeType.rect, {
      x: cardX, y: cardStartY, w: cardW, h: HEADER_H,
      fill: { color: "1E56DB" }, line: { color: "1E56DB", pt: 0 }
    });

    slide5.addText(card.header + "\n" + card.subheader, {
      x: cardX + 0.15, y: cardStartY, w: cardW - 0.3, h: HEADER_H,
      fontFace: "Segoe UI", fontSize: 12, bold: true, color: "FFFFFF",
      valign: "middle", align: "left", margin: 0, breakLine: true
    });

    // Body (Gray)
    slide5.addShape(pptx.ShapeType.rect, {
      x: cardX, y: cardStartY + HEADER_H, w: cardW, h: BODY_H,
      fill: { color: "F2F2F2" }, line: { color: "F2F2F2", pt: 0 }
    });

    // Bullets
    const bulletText = card.bullets.map((b) => "• " + b).join("\n\n");
    slide5.addText(bulletText, {
      x: cardX + 0.15,
      y: cardStartY + HEADER_H + 0.2,
      w: cardW - 0.3,
      h: BODY_H - 0.4,
      fontFace: "Segoe UI",
      fontSize: 10.5,
      color: "111111",
      valign: "top",
      breakLine: true,
      margin: 0,
    });
  });

  slide5.addText("Additional detail in appendix", {
    x: 0.38, y: 7.1, w: 4, h: 0.3,
    fontFace: "Segoe UI", fontSize: 11, color: "111111", margin: 0,
  });

  //////////////////////////////////////////////////////
  // SLIDE 6 — QUALITATIVE FEEDBACK
  //////////////////////////////////////////////////////

  const slide6 = pptx.addSlide();
  slide6.background = { color: "FFFFFF" };


  slide6.addText("Qualitative feedback", {
    x: 0.38, y: 0.15, w: 8, h: 0.5,
    fontFace: "Segoe UI Semibold", fontSize: 40, color: "111111", margin: 0,
  });

  slide6.addText("This page summarises key themes from the qualitative feedback your respondents provided.", {
    x: 0.38, y: 0.75, w: 10, h: 0.3,
    fontFace: "Segoe UI", fontSize: 11, color: "111111", margin: 0,
  });

  const qualCardW = 5.9;
  const qualCardStartY = 1.4;
  const qualCardGap = 0.3;
  const qualStartX = 0.4;

  const qualCards = [
    {
      header: "Perceived areas of strength",
      description: "You are perceived as operating across boundaries with credibility - colleagues describe this as your most distinctive contribution.",
      bullets: [
        "Seen as bringing the right stakeholders in early, which is felt to reduce rework and accelerate decisions.",
        "Perceived as trusted across functions, including on difficult topics.",
        "Experienced as setting direction with clarity - colleagues note they leave meetings clear on what was agreed and who owns what.",
        "Described as thinking at the system level, not only within your own area.",
      ],
    },
    {
      header: "Perceived areas of improvement",
      description: "Your support is perceived as genuine but reactive. Development of others appears to be the clearest opportunity for improvement.",
      bullets: [
        "Feedback is felt to follow a request rather than being offered proactively.",
        "Recognised potential is not always seen as translated into a clear stretch or development plan.",
        "Priorities are perceived as set clearly, but follow-through on standards is reported as less visible.",
        "Low-value work is seen as continuing past its useful life.",
      ],
    },
  ];

  qualCards.forEach((card, idx) => {
    const cardX = qualStartX + idx * (qualCardW + qualCardGap);
    const HEADER_H = 0.45;
    const BODY_H = 5.3;

    // Header (Blue)
    slide6.addShape(pptx.ShapeType.rect, {
      x: cardX, y: qualCardStartY, w: qualCardW, h: HEADER_H,
      fill: { color: "1E56DB" }, line: { color: "1E56DB", pt: 0 }
    });

    slide6.addText(card.header, {
      x: cardX + 0.15, y: qualCardStartY, w: qualCardW - 0.3, h: HEADER_H,
      fontFace: "Segoe UI", fontSize: 13, bold: true, color: "FFFFFF",
      valign: "middle", align: "left", margin: 0,
    });

    // Body (Gray)
    slide6.addShape(pptx.ShapeType.rect, {
      x: cardX, y: qualCardStartY + HEADER_H, w: qualCardW, h: BODY_H,
      fill: { color: "F2F2F2" }, line: { color: "F2F2F2", pt: 0 }
    });

    // Content: Paragraph then bullets
    const bulletText = card.bullets.map((b) => "•  " + b).join("\n\n");
    const fullContent = card.description + "\n\n" + bulletText;

    slide6.addText(fullContent, {
      x: cardX + 0.15,
      y: qualCardStartY + HEADER_H + 0.2,
      w: qualCardW - 0.3,
      h: BODY_H - 0.4,
      fontFace: "Segoe UI",
      fontSize: 10.5,
      color: "111111",
      valign: "top",
      breakLine: true,
      margin: 0,
    });
  });

  slide6.addText("Additional detail in appendix", {
    x: 0.38, y: 7.1, w: 4, h: 0.3,
    fontFace: "Segoe UI", fontSize: 11, color: "111111", margin: 0,
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
  // APPENDIX — HOW THE SCORE IS CALCULATED
  //////////////////////////////////////////////////////

  const slideScoreCalc = pptx.addSlide();
  slideScoreCalc.background = { color: "FFFFFF" };

  slideScoreCalc.addText("How the score is calculated", {
    x: 0.38, y: 0.2, w: 9, h: 0.6,
    fontFace: "Segoe UI Semibold", fontSize: 32, color: "111111", margin: 0,
  });

  slideScoreCalc.addText("From rater responses to a single percentage score, in three steps.", {
    x: 0.4, y: 0.75, w: 9, h: 0.3,
    fontFace: "Segoe UI", fontSize: 12, color: "555555", margin: 0,
  });

  // THREE CARDS TOP
  const CARD_W = 4.1;
  const CARD_H = 3.5;
  const CARD_START_Y = 1.15;
  const CARD_HEADER_H = 0.55;
  const CARD_GAP = 0.25;
  const CARD_COLOR_HEADER = "2C8E8F"; // Teal
  const CARD_COLOR_BG = "F2F4F7";

  const calcSteps = [
    {
      title: "Step 1. Each response is weighted",
      content: "Each rater chooses how often they see a behavior. Each choice carries a weight:\n\nAlways         →  100% of the time\nOften            →  75% of the time\nSometimes   →  50% of the time\nRarely           →  25% of the time"
    },
    {
      title: "Step 2. Average across raters",
      content: "For each behavior, take the mean of the assigned percentage values:\n\nscore   =   sum(v)  /  n\n\nv  =   percentage value\nn  =   raters who answered"
    },
    {
      title: "Step 3. Apply the threshold band",
      content: "Round the score to the nearest whole percentage, then place it in a band:\n\n> 85 %\nConsistently observed\n\n70 - 85 %\nModerately observed\n\n< 70 %\nInconsistently observed"
    }
  ];

  calcSteps.forEach((step, idx) => {
    const x = 0.4 + idx * (CARD_W + CARD_GAP);

    // Header
    slideScoreCalc.addShape(pptx.ShapeType.rect, {
      x, y: CARD_START_Y, w: CARD_W, h: CARD_HEADER_H,
      fill: { color: CARD_COLOR_HEADER }
    });
    slideScoreCalc.addText(step.title, {
      x: x + 0.15, y: CARD_START_Y, w: CARD_W - 0.3, h: CARD_HEADER_H,
      fontFace: "Segoe UI", fontSize: 12.5, bold: true, color: "FFFFFF", valign: "middle", margin: 0
    });

    // Body
    slideScoreCalc.addShape(pptx.ShapeType.rect, {
      x, y: CARD_START_Y + CARD_HEADER_H, w: CARD_W, h: CARD_H - CARD_HEADER_H,
      fill: { color: CARD_COLOR_BG }
    });
    slideScoreCalc.addText(step.content, {
      x: x + 0.15, y: CARD_START_Y + CARD_HEADER_H + 0.2, w: CARD_W - 0.3, h: CARD_H - CARD_HEADER_H - 0.4,
      fontFace: "Segoe UI", fontSize: 11, color: "333333", valign: "top", margin: 0, breakLine: true
    });
  });

  // EXAMPLE PANEL AT BOTTOM
  const EX_Y = 4.9;
  const EX_H = 1.9;
  slideScoreCalc.addShape(pptx.ShapeType.rect, {
    x: 0.4, y: EX_Y, w: 12.8, h: EX_H,
    fill: { color: "FFFFFF" }, line: { color: "D9DEE5", pt: 1 }
  });
  slideScoreCalc.addText("EXAMPLE", {
    x: 0.6, y: EX_Y + 0.15, w: 2, h: 0.2,
    fontFace: "Segoe UI", fontSize: 10, bold: true, color: "1C2B4B"
  });
  slideScoreCalc.addText(
    "Behavior: Makes timely decisions   •   n = 12 raters answered   •   illustrative\n\nCounts:  1 Rarely  ·  1 Sometimes  ·  5 Often  ·  5 Always\nScore  =  (1·25 + 1·50 + 5·75 + 5·100) / 12\n            =  (25 + 50 + 375 + 500) / 12  =  950 / 12   ≈   79.2 %\nRounded to 79 %",
    {
      x: 0.6, y: EX_Y + 0.5, w: 9, h: 1.2,
      fontFace: "Segoe UI", fontSize: 11, color: "333333", margin: 0
    }
  );

  // Score Badge in Example
  slideScoreCalc.addShape(pptx.ShapeType.roundRect, {
    x: 9.8, y: EX_Y + 0.55, w: 2.8, h: 1.0,
    fill: { color: "DAA520" }, rectRadius: 0.1
  });
  slideScoreCalc.addText("79 %", {
    x: 9.8, y: EX_Y + 0.6, w: 2.8, h: 0.5,
    fontFace: "Segoe UI Semibold", fontSize: 24, color: "FFFFFF", align: "center", margin: 0
  });
  slideScoreCalc.addText("Moderately observed", {
    x: 9.8, y: EX_Y + 1.1, w: 2.8, h: 0.2,
    fontFace: "Segoe UI", fontSize: 9.5, color: "FFFFFF", align: "center", margin: 0
  });

  slideScoreCalc.addText("Note: Insufficient Exposure responses are excluded from n; they are never counted as zero. Scores are rounded to the nearest whole percent before bands are applied.", {
    x: 0.4, y: 6.85, w: 12.8, h: 0.2,
    fontFace: "Segoe UI", fontSize: 8.5, italic: true, color: "666666", margin: 0
  });

  // Footer for Score Calc
  slideScoreCalc.addShape(pptx.ShapeType.rect, {
    x: 10.4, y: 7.2, w: 2.5, h: 0.3,
    fill: { color: "FFFFFF" }, line: { color: "B00000", pt: 1 },
  });
  slideScoreCalc.addText(`${form.responses[0]?.participant?.fullName || "PARTICIPANT"} & ${new Date().toLocaleDateString()}`, {
    x: 10.45, y: 7.25, w: 2.4, h: 0.2,
    fontFace: "Segoe UI", fontSize: 7, color: "B00000", bold: true, align: "center",
  });

  //////////////////////////////////////////////////////
  // APPENDIX — FAQ
  //////////////////////////////////////////////////////

  const slideFAQ = pptx.addSlide();
  slideFAQ.background = { color: "FFFFFF" };

  slideFAQ.addText("FAQ", {
    x: 0.38, y: 0.2, w: 9, h: 0.8,
    fontFace: "Segoe UI Semibold", fontSize: 48, color: "111111", margin: 0,
  });

  const FAQ_W = 6.2;
  const FAQ_H = 1.6;
  const FAQ_GAP_X = 0.35;
  const FAQ_GAP_Y = 0.25;
  const FAQ_START_X = 0.4;
  const FAQ_START_Y = 1.3;

  const faqs = [
    {
      q: "Can I find out who said what?",
      a: "No. Individual responses are confidential and reported only in aggregate. Where a rater group is very small (e.g., one or two direct reports), results may be combined with another group to protect anonymity. The verbatim quotes are also presented without attribution."
    },
    {
      q: "What is a personal development plan?",
      a: "A short, practical document you build with your coach to turn this feedback into action. It names two or three behaviors to work on, what good looks like for each, and the specific actions, stretch assignments or conversations you'll use over the next 90 days to make progress visible."
    },
    {
      q: "How is the score calculated?",
      a: "For each behavior, raters choose how often they see it: always, often, sometimes, rarely. We average their answers into a single score, then show the share of time the behavior is seen."
    },
    {
      q: "What counts as a meaningful gap in self / others?",
      a: "Under 10 points is noise. Above 10 is worth a conversation. The most interesting gaps are usually where you rate yourself higher than your colleagues do."
    },
    {
      q: "What does \"Insufficient Exposure\" mean?",
      a: "A rater chooses this when they haven't worked closely enough with you to judge a behavior fairly. These answers are left out of the score and never counted as low. High Insufficient Exposure scores can itself be a signal: people may not be seeing you act in that area."
    },
    {
      q: "Who sees my report?",
      a: "Your report is shared only with you and HR. It is not shared with your manager or anyone else without your explicit consent. It is not used in performance, pay or promotion decisions. However, we do encourage you to use this report to facilitate development discussions with key stakeholders."
    },
  ];

  faqs.forEach((faq, idx) => {
    const col = idx % 2;
    const row = Math.floor(idx / 2);
    const x = FAQ_START_X + col * (FAQ_W + FAQ_GAP_X);
    const y = FAQ_START_Y + row * (FAQ_H + FAQ_GAP_Y);

    // Box
    slideFAQ.addShape(pptx.ShapeType.rect, {
      x, y, w: FAQ_W, h: FAQ_H,
      fill: { color: "E9ECF1" }, // Light blue-gray
      rectRadius: 0.05
    });

    // Q Icon
    const ICON_SIZE = 0.28;
    slideFAQ.addShape(pptx.ShapeType.ellipse, {
      x: x + 0.15, y: y + 0.15, w: ICON_SIZE, h: ICON_SIZE,
      fill: { color: "1C2B4B" }
    });
    slideFAQ.addText("Q", {
      x: x + 0.15, y: y + 0.15, w: ICON_SIZE, h: ICON_SIZE,
      fontFace: "Segoe UI", fontSize: 10, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0
    });

    // Question
    slideFAQ.addText(faq.q, {
      x: x + 0.5, y: y + 0.15, w: FAQ_W - 0.7, h: 0.3,
      fontFace: "Segoe UI", fontSize: 12, bold: true, color: "1C2B4B", valign: "top", margin: 0
    });

    // Answer
    slideFAQ.addText(faq.a, {
      x: x + 0.5, y: y + 0.55, w: FAQ_W - 0.7, h: 0.9,
      fontFace: "Segoe UI", fontSize: 9.5, color: "333333", valign: "top", margin: 0, breakLine: true
    });
  });

  // Footer for FAQ
  slideFAQ.addShape(pptx.ShapeType.rect, {
    x: 10.4, y: 7.2, w: 2.5, h: 0.3,
    fill: { color: "FFFFFF" }, line: { color: "B00000", pt: 1 },
  });
  slideFAQ.addText(`${form.responses[0]?.participant?.fullName || "PARTICIPANT"} & ${new Date().toLocaleDateString()}`, {
    x: 10.45, y: 7.25, w: 2.4, h: 0.2,
    fontFace: "Segoe UI", fontSize: 7, color: "B00000", bold: true, align: "center",
  });

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