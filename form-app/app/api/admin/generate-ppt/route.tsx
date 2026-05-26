import { NextRequest } from "next/server";
import pptxgen from "pptxgenjs";
import prisma from "@/lib/prisma";
import fs from "fs";
import path from "path";

import verbatimQuotesData, {
  coachingQuadrantData,
  leadershipImpactData,
  personalDevelopmentPlanData,
  positiveImpactData,
  resultsAtGlanceData,
  verbatimThemesData,
} from "../form-participants/datas";

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
// ERICSSON BRANDING HELPERS
//////////////////////////////////////////////////////

function getEricssonLogoSvg(color: string) {
  const hexColor = color.startsWith("#") ? color : `#${color}`;
  const dPath = "M20.76 1.593A2.36 2.36 0 0 0 19.572.225c-.54-.27-1.188-.336-2.256.02L5.187 4.29c-1.068.357-1.548.795-1.818 1.338a2.36 2.36 0 0 0 1.059 3.174c.54.27 1.188.336 2.256-.021l12.129-4.044c1.068-.354 1.548-.795 1.818-1.338a2.35 2.35 0 0 0 .13-1.806zm0 7.485a2.36 2.36 0 0 0-1.188-1.368c-.54-.27-1.188-.336-2.256.021L5.187 11.775c-1.068.357-1.548.795-1.818 1.338a2.36 2.36 0 0 0 1.059 3.174c.54.27 1.188.336 2.256-.021l12.129-4.041c1.068-.357 1.548-.795 1.818-1.341a2.35 2.35 0 0 0 .13-1.806zm0 7.488a2.36 2.36 0 0 0-1.188-1.368c-.54-.27-1.188-.336-2.256.021L5.187 19.263c-1.068.357-1.548.795-1.818 1.338a2.36 2.36 0 0 0 1.059 3.174c.54.27 1.188.336 2.256-.02l12.129-4.045c1.068-.354 1.548-.795 1.818-1.338a2.35 2.35 0 0 0 .13-1.806z";
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="${hexColor}" d="${dPath}"/></svg>`;
}

function getEricssonLogoBase64(color: string) {
  const svg = getEricssonLogoSvg(color);
  return "data:image/svg+xml;base64," + Buffer.from(svg).toString("base64");
}

function addLogoToSlide(slide: any, isDarkBg: boolean) {
  const color = isDarkBg ? "FFFFFF" : "000000";
  const logoBase64 = getEricssonLogoBase64(color);

  // Add Logo symbol
  slide.addImage({
    data: logoBase64,
    x: 12.45,
    y: 0.22,
    w: 0.35,
    h: 0.35,
  });

  // Add "ERICSSON" text underneath it
  slide.addText("ERICSSON", {
    x: 12.15,
    y: 0.59,
    w: 0.95,
    h: 0.15,
    fontFace: "Segoe UI",
    fontSize: 6.5,
    bold: true,
    color: color,
    align: "center",
    margin: 0,
  });
}

//////////////////////////////////////////////////////
// GLOBAL FOOTER HELPER
//////////////////////////////////////////////////////

function addFooterToSlide(
  slide: any,
  pptx: any,
  participantName: string,
  date: string,
  pageNumber: number,
  x: number = 10.4,
  y: number = 7.3,
  w: number = 2.5,
  h: number = 0.18
) {
  // ── Left: "Leadership Assessments  |  Ericsson  |  Page N" ───────────────
  slide.addText(`Leadership Assessments  |  Ericsson  |  Page ${pageNumber}`, {
    x: 0.25,
    y: y,
    w: 6.0,
    h: h,
    fontFace: "Segoe UI",
    fontSize: 7,
    color: "888888",
    margin: 0,
    valign: "middle",
    align: "left",
  });

  // ── Right: red participant name & date box ────────────────────────────────
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: "FFFFFF" },
    line: { color: "B00000", pt: 1 },
  });

  slide.addText(`${participantName} & ${date}`, {
    x: x + 0.05,
    y: y + 0.02,
    w: w - 0.1,
    h: h - 0.04,
    fontFace: "Segoe UI",
    fontSize: 6.5,
    color: "B00000",
    bold: true,
    align: "center",
    valign: "middle",
    margin: 0,
  });
}

//////////////////////////////////////////////////////
// SLIDE INTERFACE
//////////////////////////////////////////////////////

interface SlideParams {
  pptx: pptxgen;
  participantName: string;
  footerDate: string;
  pageNumber: number;
  analytics: any[];
  selfAnalytics: any[];
  othersAnalytics: any[];
  enterpriseAnalytics: any[];
  diagramBase64?: string;
}

//////////////////////////////////////////////////////
// MODULAR SLIDE CREATION FUNCTIONS
//////////////////////////////////////////////////////

export function createSlideIntro(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber, diagramBase64 } = params;
  const slideIntro = pptx.addSlide();
  slideIntro.background = { color: "111111" };
  addLogoToSlide(slideIntro, true);

  // TITLE
  slideIntro.addText("What is a 360 diagnostic", {
    x: 0.4, y: 0.25, w: 9, h: 0.9,
    fontFace: "Aptos", fontSize: 40, bold: false, color: "FFFFFF", margin: 0,
  });

  if (diagramBase64) {
    slideIntro.addImage({
      data: "image/png;base64," + diagramBase64,
      x: 0.2,
      y: 1.2,
      w: 5.8,
      h: 5.8,
    });
  }

  // RIGHT SIDE TEXT
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

  addFooterToSlide(slideIntro, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideHowCreated(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slideHowCreated = pptx.addSlide();
  slideHowCreated.background = { color: "111111" };
  addLogoToSlide(slideHowCreated, true);

  // TITLES
  slideHowCreated.addText("How this report was created", {
    x: 0.5, y: 0.4, w: 8, h: 1.2,
    fontFace: "Segoe UI", fontSize: 40, bold: false, color: "FFFFFF", margin: 0,
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

  addFooterToSlide(slideHowCreated, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideHowToUse(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slideHowToUse = pptx.addSlide();
  slideHowToUse.background = { color: "111111" };
  addLogoToSlide(slideHowToUse, true);

  const HTU_CARD_X: number[] = [0.3, 3.52, 6.74, 9.96];
  const HTU_CARD_W: number = 3.07;
  const HTU_HEADER_Y: number = 1.5;
  const HTU_HEADER_H: number = 0.78;
  const HTU_BODY_Y: number = HTU_HEADER_Y + HTU_HEADER_H;
  const HTU_BODY_H: number = 3.9;
  const HTU_BANNER_Y: number = 6.28;
  const HTU_BANNER_H: number = 0.5;
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

  addFooterToSlide(slideHowToUse, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideOverallResults(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber, analytics } = params;
  const slide1 = pptx.addSlide();
  slide1.background = { color: COLORS.bg };

  slide1.addText("Overall results", {
    x: 0.38, y: 0.22, w: 2.8, h: 0.5,
    fontFace: "Segoe UI", fontSize: 30, color: "111111", margin: 0,
  });

  slide1.addText(
    "How consistently micro-behaviors are seen in practice by your respondents.",
    {
      x: 0.38, y: 0.88, w: 7, h: 0.2,
      fontFace: "Segoe UI", fontSize: 11, color: "444444", margin: 0,
    }
  );

  // Top center-right blue badges
  slide1.addShape(pptx.ShapeType.roundRect, {
    x: 3.3, y: 0.12, w: 1.7, h: 0.55,
    fill: { color: "1D6FD8" }, line: { color: "1D6FD8" },
    rectRadius: 0.12
  });
  slide1.addText("Micro-behaviors not yet\nfinalized", {
    x: 3.3, y: 0.12, w: 1.7, h: 0.55,
    fontFace: "Segoe UI", fontSize: 8.5, bold: true, color: "FFFFFF",
    align: "center", valign: "middle", margin: 0,
  });

  slide1.addShape(pptx.ShapeType.roundRect, {
    x: 5.1, y: 0.12, w: 2.1, h: 0.35,
    fill: { color: "1D6FD8" }, line: { color: "1D6FD8" },
    rectRadius: 0.12
  });
  slide1.addText("ERICSSON TO DEFINE THRESHOLDS", {
    x: 5.1, y: 0.12, w: 2.1, h: 0.35,
    fontFace: "Segoe UI", fontSize: 8.0, bold: true, color: "FFFFFF",
    align: "center", valign: "middle", margin: 0,
  });

  // Legend
  const legends = [
    { color: COLORS.green, text: "Consistently\nobserved (>85%)", x: 7.3 },
    { color: "D96B27", text: "Moderately\nobserved (85-70%)", x: 8.8 },
    { color: COLORS.red, text: "Inconsistently\nobserved (<70%)", x: 10.4 },
  ];

  legends.forEach((l) => {
    slide1.addShape(pptx.ShapeType.rect, {
      x: l.x, y: 0.22, w: 0.16, h: 0.16,
      fill: { color: l.color }, line: { color: l.color },
    });
    slide1.addText(l.text, {
      x: l.x + 0.22, y: 0.14, w: 1.2, h: 0.35,
      fontFace: "Segoe UI", fontSize: 7.5, color: COLORS.text, margin: 0,
      valign: "middle", breakLine: true,
    });
  });

  addLogoToSlide(slide1, false);

  const SPLIT_INDEX = Math.ceil(analytics.length / 2);
  let leftY = 1.22;
  let rightY = 1.22;
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

slide1.addShape(pptx.ShapeType.rect, {
  x: barX,
  y: currentY - 0.01,

  w: 0.008,
  h: 0.42,

  fill: { color: "9E9E9E" },
  line: {
    color: "9E9E9E",
    pt: 0,
  },
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
      x: scoreX, y: currentY + 0.02, w: 0.8, h: 0.2,
      fontFace: "Segoe UI", fontSize: 13,
      bold: false, color: COLORS.text, margin: 0,
    });

    if (isLeft) leftY += Y_GAP; else rightY += Y_GAP;
  });

  slide1.addText("N.B. Scores are rounded to the nearest whole percentage before thresholds are applied.", {
    x: 0.38, y: 6.82, w: 8.0, h: 0.18,
    fontFace: "Segoe UI", fontSize: 7.5, color: "555555", margin: 0,
  });

  addFooterToSlide(slide1, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideResultsDistribution(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber, othersAnalytics } = params;
  const slide2 = pptx.addSlide();
  slide2.background = { color: "FFFFFF" };
  addLogoToSlide(slide2, false);

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

  // TABLE CONFIG
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
  const STRIPE_H = othersAnalytics.length * ROW_H;
  slide2.addShape(pptx.ShapeType.rect, { x: COL_X[2], y: TABLE_Y + HEADER_H, w: COL_W[2], h: STRIPE_H, fill: { color: "F9EBEB" } });
  slide2.addShape(pptx.ShapeType.rect, { x: COL_X[3], y: TABLE_Y + HEADER_H, w: COL_W[3], h: STRIPE_H, fill: { color: "FFF0F0" } });
  slide2.addShape(pptx.ShapeType.rect, { x: COL_X[4], y: TABLE_Y + HEADER_H, w: COL_W[4], h: STRIPE_H, fill: { color: "F5F5F5" } });
  slide2.addShape(pptx.ShapeType.rect, { x: COL_X[5], y: TABLE_Y + HEADER_H, w: COL_W[5], h: STRIPE_H, fill: { color: "F0F9F0" } });

  const CHART_CENTER_X = COL_X[1] + (COL_W[1] / 2);

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

    // Bars
    const total = (dist.Rarely || 0) + (dist.Sometimes || 0) + (dist.Often || 0) + (dist.Always || 0);
    if (total > 0) {
      const HALF_BAR_MAX_W = (COL_W[1] / 2) - 0.1;
      const scale = HALF_BAR_MAX_W / total;

      const wR = (dist.Rarely || 0) * scale;
      const wS = (dist.Sometimes || 0) * scale;
      const wO = (dist.Often || 0) * scale;
      const wA = (dist.Always || 0) * scale;

      const barY = y + (ROW_H * 0.2);
      const barH = ROW_H * 0.6;

      if (wR > 0) slide2.addShape(pptx.ShapeType.rect, { x: CHART_CENTER_X - wS - wO - wR, y: barY, w: wR, h: barH, fill: { color: "9E0B0F" }, line: { pt: 0 } });
      if (wS > 0) slide2.addShape(pptx.ShapeType.rect, { x: CHART_CENTER_X - wO - wS, y: barY, w: wS, h: barH, fill: { color: "FF0000" }, line: { pt: 0 } });
      if (wO > 0) slide2.addShape(pptx.ShapeType.rect, { x: CHART_CENTER_X - wO, y: barY, w: wO, h: barH, fill: { color: "AD9380" }, line: { pt: 0 } });
      if (wA > 0) slide2.addShape(pptx.ShapeType.rect, { x: CHART_CENTER_X, y: barY, w: wA, h: barH, fill: { color: "008032" }, line: { pt: 0 } });
    }
  });

  // Strong center axis
  slide2.addShape(pptx.ShapeType.line, {
    x: CHART_CENTER_X, y: TABLE_Y + HEADER_H, w: 0, h: STRIPE_H,
    line: { color: "111111", pt: 1 }
  });

  addFooterToSlide(slide2, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideInsufficientExposures(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber, analytics } = params;
  const slide3 = pptx.addSlide();
  slide3.background = { color: "FFFFFF" };
  addLogoToSlide(slide3, false);

  slide3.addText("Insufficient Exposures", {
    x: 0.25, y: 0.08, w: 8.0, h: 0.55,
    fontFace: "Segoe UI", fontSize: 28, color: "1A1A1A", bold: false, margin: 0,
  });

  const TABLE_LEFT = 0.25;
  const TABLE_W = 6.55;
  const COL1_W = 3.55;
  const COL2_W = TABLE_W - COL1_W;
  const COL2_X = TABLE_LEFT + COL1_W;
  const HEADER_Y = 0.72;
  const HEADER_H = 0.28;
  const DATA_START_Y = HEADER_Y + HEADER_H;   // 1.00"

  // ── Dynamic row height: fit ALL rows between header and footer ───────────
  // Footer sits at y ≈ 6.82"; leave 0.05" breathing room above it.
  const FOOTER_TOP = 6.82;
  const AVAILABLE_H = FOOTER_TOP - DATA_START_Y;           // 5.82"
  const ROW_H = analytics.length > 0
    ? Math.min(0.255, AVAILABLE_H / analytics.length)      // shrink only when needed
    : 0.255;

  // Header row background
  slide3.addShape(pptx.ShapeType.rect, {
    x: TABLE_LEFT, y: HEADER_Y, w: TABLE_W, h: HEADER_H,
    fill: { color: "D0D0D0" }, line: { color: "D0D0D0" },
  });

  // Header text col 1
  slide3.addText("Results", {
    x: TABLE_LEFT + 0.06, y: HEADER_Y, w: COL1_W, h: HEADER_H,
    fontFace: "Segoe UI", fontSize: 8.5, bold: true, color: "111111",
    margin: 0, valign: "middle", align: "left",
  });

  // Header text col 2
  slide3.addText("Count of Insufficient Exposure", {
    x: COL2_X, y: HEADER_Y, w: COL2_W, h: HEADER_H,
    fontFace: "Segoe UI", fontSize: 8.5, bold: true, color: "111111",
    margin: 0, valign: "middle", align: "left",
  });

  // Vertical divider between col 1 and col 2
  slide3.addShape(pptx.ShapeType.line, {
    x: COL2_X, y: HEADER_Y,
    w: 0.001, h: HEADER_H + (analytics.length * ROW_H),
    line: { color: "A0A0A0", pt: 0.75 },
  });

  // Outer border of entire table
  slide3.addShape(pptx.ShapeType.rect, {
    x: TABLE_LEFT, y: HEADER_Y,
    w: TABLE_W, h: HEADER_H + (analytics.length * ROW_H),
    fill: { type: "none" }, line: { color: "B0B0B0", pt: 0.75 },
  });

  // ── Data rows ─────────────────────────────────────────────────────────────
  analytics.forEach((item, idx) => {
    const count = item.distribution["Insufficient Exposure"] || 0;
    const rowY = DATA_START_Y + idx * ROW_H;
    const isEven = idx % 2 === 0;

    if (isEven) {
      slide3.addShape(pptx.ShapeType.rect, {
        x: TABLE_LEFT, y: rowY, w: TABLE_W, h: ROW_H,
        fill: { color: "F9F9F9" }, line: { pt: 0 },
      });
    }

    // Row bottom border
    slide3.addShape(pptx.ShapeType.line, {
      x: TABLE_LEFT, y: rowY + ROW_H,
      w: TABLE_W, h: 0.001,
      line: { color: "D0D0D0", pt: 0.5 },
    });

    // Question label
    slide3.addText(item.question, {
      x: TABLE_LEFT + 0.06, y: rowY + 0.01, w: COL1_W - 0.08, h: ROW_H - 0.02,
      fontFace: "Segoe UI", fontSize: 7.8, color: "222222",
      margin: 0, valign: "middle", fit: "shrink",
    });

    if (count === 0) {
      slide3.addText("0", {
        x: COL2_X + 0.08, y: rowY + 0.01, w: 0.5, h: ROW_H - 0.02,
        fontFace: "Segoe UI", fontSize: 7.8, color: "333333",
        margin: 0, valign: "middle", align: "left",
      });
    } else {
      const BAR_X = COL2_X + 0.08;
      const BAR_H = Math.min(0.09, ROW_H * 0.45);     // bar height scales with row
      const BAR_Y = rowY + (ROW_H - BAR_H) / 2;
      const PX_PER_CNT = 0.38;
      const barW = count * PX_PER_CNT;

      slide3.addShape(pptx.ShapeType.rect, {
        x: BAR_X, y: BAR_Y, w: barW, h: BAR_H,
        fill: { color: "0055D4" }, line: { pt: 0 },
      });

      slide3.addText(`${count}`, {
        x: BAR_X + barW + 0.07, y: rowY + 0.01, w: 0.4, h: ROW_H - 0.02,
        fontFace: "Segoe UI", fontSize: 7.8, color: "333333",
        margin: 0, valign: "middle", align: "left",
      });
    }
  });

  // ── NOTE BOX ──────────────────────────────────────────────────────────────
  // Fixed position on right side; vertically centred in the lower half.
  // NOTE_Y = 3.48 keeps it away from the header/logo area and matches design.
  const NOTE_X = 6.95;
  const NOTE_Y = 3.48;
  const NOTE_W = 3.55;
  const NOTE_HDR_H = 0.26;
  const NOTE_BODY_Y = NOTE_Y + NOTE_HDR_H;
  // Body stretches down to just above the footer (leaves 0.10" gap)
  const NOTE_BODY_H = FOOTER_TOP - NOTE_BODY_Y - 0.10;

  // Header bar
  slide3.addShape(pptx.ShapeType.rect, {
    x: NOTE_X, y: NOTE_Y, w: NOTE_W, h: NOTE_HDR_H,
    fill: { color: "A6A6A6" }, line: { color: "A6A6A6", pt: 0.5 },
  });

  slide3.addText("Note", {
    x: NOTE_X + 0.10, y: NOTE_Y + 0.01, w: 0.6, h: NOTE_HDR_H - 0.02,
    fontFace: "Segoe UI", fontSize: 8.5, bold: true, color: "FFFFFF",
    margin: 0, valign: "middle",
  });

  // Body background
  slide3.addShape(pptx.ShapeType.rect, {
    x: NOTE_X, y: NOTE_BODY_Y, w: NOTE_W, h: NOTE_BODY_H,
    fill: { color: "F2F2F2" }, line: { color: "D9D9D9", pt: 0.5 },
  });

  // Body text
  slide3.addText(
    [
      { text: "Insufficient Exposure", options: { bold: false } },
      {
        text: " is what a rater selects when they feel they have not seen sufficient evidence from you on this behavior to provide a rating. These answers are set ",
        options: {},
      },
      { text: "aside", options: { underline: true as any, color: "0563C1" } },
      {
        text: " and they do not lower your score.\n\nA high count can itself be a signal as it can mean a behavior is not be visible to the people around you, even if you feel you are demonstrating it.\n\nIf you have areas with high insufficient exposure, we recommend you consider why you are receiving this feedback and what you could do to address this.",
        options: {},
      },
    ],
    {
      x: NOTE_X + 0.10,
      y: NOTE_BODY_Y + 0.08,
      w: NOTE_W - 0.18,
      h: NOTE_BODY_H - 0.14,
      fontFace: "Segoe UI",
      fontSize: 8.3,
      color: "333333",
      margin: 0,
      breakLine: true,
      valign: "top",
      fit: "shrink",
    }
  );

  addFooterToSlide(slide3, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideSelfOther(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slide5 = pptx.addSlide();
  slide5.background = { color: "FFFFFF" };
  addLogoToSlide(slide5, false);

  // ── TITLE & SUBTITLE ───────────────────────────────────────────────────────
  slide5.addText("Self / other", {
    x: 0.38, y: 0.20, w: 8, h: 0.5,
    fontFace: "Segoe UI Semibold", fontSize: 36, color: "111111", margin: 0,
  });

  slide5.addText("This page shows where your self-view differs from how colleagues experience your micro-behaviors.", {
    x: 0.38, y: 0.75, w: 10, h: 0.3,
    fontFace: "Segoe UI", fontSize: 11.5, color: "1B365D", margin: 0,
  });

  // ── CARD CONFIGURATION ─────────────────────────────────────────────────────
  const cardW = 3.98;
  const cardStartY = 1.40;
  const cardGap = 0.30;
  const startX = 0.39;

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
    const HEADER_H = 0.90;
    const BODY_H = 3.75;
    const RADIUS = 0.08;

    // Header Background: rounded rectangle
    slide5.addShape(pptx.ShapeType.roundRect, {
      x: cardX, y: cardStartY, w: cardW, h: HEADER_H,
      rectRadius: RADIUS,
      fill: { color: "1D6FD8" }, line: { color: "1D6FD8", pt: 0 }
    });

    // Flatten bottom corners of the header by overlaying a small flat rect
    slide5.addShape(pptx.ShapeType.rect, {
      x: cardX, y: cardStartY + HEADER_H - 0.12, w: cardW, h: 0.12,
      fill: { color: "1D6FD8" }, line: { color: "1D6FD8", pt: 0 }
    });

    // Header Text: Multi-weight
    slide5.addText(
      [
        { text: card.header + "\n", options: { bold: true, fontSize: 13, color: "FFFFFF" } },
        { text: card.subheader, options: { bold: false, fontSize: 10.5, color: "FFFFFF" } }
      ],
      {
        x: cardX + 0.20, y: cardStartY + 0.08, w: cardW - 0.40, h: HEADER_H - 0.16,
        fontFace: "Segoe UI", valign: "middle", align: "left", margin: 0,
      }
    );

    // Body Background: rounded rectangle
    slide5.addShape(pptx.ShapeType.roundRect, {
      x: cardX, y: cardStartY + HEADER_H, w: cardW, h: BODY_H,
      rectRadius: RADIUS,
      fill: { color: "F4F4F6" }, line: { color: "F4F4F6", pt: 0 }
    });

    // Flatten top corners of the body by overlaying a small flat rect
    slide5.addShape(pptx.ShapeType.rect, {
      x: cardX, y: cardStartY + HEADER_H, w: cardW, h: 0.12,
      fill: { color: "F4F4F6" }, line: { color: "F4F4F6", pt: 0 }
    });

    // Bullets: Custom vibrant teal bullets with navy text
    let bulletY = cardStartY + HEADER_H + 0.25;
    const bulletSpacing = 0.90;

    card.bullets.forEach((bullet) => {
      slide5.addText(
        [
          { text: "•  ", options: { color: "1B7F83", bold: true } },
          { text: bullet, options: { color: "20354B" } }
        ],
        {
          x: cardX + 0.20, y: bulletY,
          w: cardW - 0.40, h: 0.70,
          fontFace: "Segoe UI", fontSize: 10.5,
          margin: 0, valign: "top",
        }
      );
      bulletY += bulletSpacing;
    });
  });

  // Footer text
  slide5.addText("Additional detail in appendix", {
    x: 0.38,
    y: 6.25,
    w: 3.5,
    h: 0.20,
    fontFace: "Segoe UI Semibold",
    fontSize: 12,
    color: "222222",
    margin: 0,
    bold: true,
  });

  addFooterToSlide(slide5, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideQualitativeFeedback(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slide6 = pptx.addSlide();
  slide6.background = { color: "F3F3F3" };
  addLogoToSlide(slide6, false);

  // TITLE
  slide6.addText("Qualitative feedback", {
    x: 0.28,
    y: 0.18,
    w: 5.5,
    h: 0.45,
    fontFace: "Segoe UI Light",
    bold:true,
    fontSize: 29,
    color: "111111",
    margin: 0,
  
  });

  // SUBTITLE
  slide6.addText(
    "This page summarises key themes from the qualitative feedback your respondents provided.",
    {
      x: 0.3,
      y: 0.82,
      w: 7.8,
      h: 0.2,
      fontFace: "Segoe UI",
      fontSize: 10.5,
      color: "20354B",
      margin: 0,
    }
  );

  // LAYOUT
  const qualCardW = 5.45;
  const qualCardGap = 0.12;
  const qualStartX = 0.22;
  const qualCardStartY = 1.25;

  const qualCards = [
    {
      header: "Perceived areas of strength",
      description:
        "You are perceived as operating across boundaries with credibility - colleagues describe this as your most distinctive contribution.",
      bullets: [
        "Seen as bringing the right stakeholders in early, which is felt to reduce rework and accelerate decisions.",
        "Perceived as trusted across functions, including on difficult topics.",
        "Experienced as setting direction with clarity - colleagues note they leave meetings clear on what was agreed and who owns what.",
        "Described as thinking at the system level, not only within your own area.",
      ],
    },
    {
      header: "Perceived areas of improvement",
      description:
        "Your support is perceived as genuine but reactive. Development of others appears to be the clearest opportunity for improvement.",
      bullets: [
        "Feedback is felt to follow a request rather than being offered proactively.",
        "Recognised potential is not always seen as translated into a clear stretch or development plan.",
        "Priorities are perceived as set clearly, but follow-through on standards is reported as less visible.",
        "Low-value work is seen as continuing past its useful life.",
      ],
    },
  ];

  // CARDS
  qualCards.forEach((card, idx) => {
    const cardX = qualStartX + idx * (qualCardW + qualCardGap);
    const HEADER_H = 0.42;
    const BODY_H = 4.65;

    // MAIN CARD
    slide6.addShape(pptx.ShapeType.roundRect, {
      x: cardX,
      y: qualCardStartY,
      w: qualCardW,
      h: BODY_H,
      rectRadius: 0.12,
      fill: { color: "E7E7ED" },
      line: { color: "E7E7ED", pt: 0 },
    });

    // HEADER
    slide6.addShape(pptx.ShapeType.rect, {
      x: cardX,
      y: qualCardStartY,
      w: qualCardW,
      h: HEADER_H,
      fill: { color: "1F73E8" },
      line: { color: "1F73E8", pt: 0 },
    });

    // HEADER TEXT
    slide6.addText(card.header, {
      x: cardX + 0.12,
      y: qualCardStartY + 0.07,
      w: qualCardW - 0.25,
      h: 0.18,
      fontFace: "Segoe UI",
      fontSize: 13,
      color: "FFFFFF",
      margin: 0,
    });

    // DESCRIPTION
    slide6.addText(card.description, {
      x: cardX + 0.16,
      y: qualCardStartY + 0.58,
      w: qualCardW - 0.35,
      h: 0.55,
      fontFace: "Segoe UI",
      fontSize: 10,
      color: "222222",
      breakLine: true,
      margin: 0,
      valign: "top",
    });

    // BULLETS
    let bulletY = qualCardStartY + 1.55;

    card.bullets.forEach((bullet) => {
      // Bullet dot
      slide6.addText("•", {
        x: cardX + 0.15,
        y: bulletY + 0.01,
        w: 0.1,
        h: 0.1,
        fontFace: "Segoe UI",
        fontSize: 10,
        color: "111111",
        margin: 0,
      });

      // Bullet text
      slide6.addText(bullet, {
        x: cardX + 0.38,
        y: bulletY,
        w: qualCardW - 0.65,
        h: 0.48,
        fontFace: "Segoe UI",
        fontSize: 10,
        color: "222222",
        breakLine: true,
        margin: 0,
        valign: "top",
      });

      bulletY += 0.72;
    });
  });

  // ADDITIONAL DETAIL
  slide6.addText("Additional detail in appendix", {
    x: 0.2,
    y: 5.95,
    w: 2.5,
    h: 0.16,
    fontFace: "Segoe UI",
    fontSize: 10,
    color: "222222",
    margin: 0,
    bold: true,
  });

  addFooterToSlide(slide6, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideScoreCalc(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slideScoreCalc = pptx.addSlide();
  slideScoreCalc.background = { color: "FFFFFF" };
  addLogoToSlide(slideScoreCalc, false);

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

  addFooterToSlide(slideScoreCalc, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideFAQ(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slideFAQ = pptx.addSlide();
  slideFAQ.background = { color: "FFFFFF" };
  addLogoToSlide(slideFAQ, false);

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

  addFooterToSlide(slideFAQ, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideVerbatimQuotes(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slide11 = pptx.addSlide();
  slide11.background = { color: "F2F2F2" };
  addLogoToSlide(slide11, false);

  // TITLE
  slide11.addText(`“${verbatimQuotesData.title}”`, {
    x: 0.35,
    y: 0.25,
    w: 7.8,
    h: 0.9,
    fontFace: "Segoe UI",
    fontSize: 30,
    color: "222222",
    margin: 0,
  });

  // SECTION TAG
  slide11.addShape(pptx.ShapeType.roundRect, {
    x: 8.85,
    y: 0.25,
    w: 2.15,
    h: 0.38,
    rectRadius: 0.04,
    fill: { color: "7A7A7A" },
    line: { color: "7A7A7A" },
  });

  slide11.addText(verbatimQuotesData.section, {
    x: 8.9,
    y: 0.32,
    w: 2.05,
    h: 0.15,
    fontFace: "Segoe UI",
    fontSize: 10,
    bold: true,
    color: "FFFFFF",
    align: "center",
    margin: 0,
  });

  // QUOTES
  let yPos = 1.4;

  verbatimQuotesData.quotes.forEach((quote) => {
    slide11.addText(`• ${quote}`, {
      x: 0.5,
      y: yPos,
      w: 11,
      h: 0.5,
      fontFace: "Segoe UI",
      fontSize: 15,
      color: "333333",
      margin: 0,
      breakLine: true,
      valign: "top",
    });

    yPos += 0.72;
  });

  addFooterToSlide(slide11, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlidePositiveImpact(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slide12 = pptx.addSlide();
  slide12.background = { color: "F2F2F2" };
  addLogoToSlide(slide12, false);

  // TITLE
  slide12.addText(`“${positiveImpactData.title}”`, {
    x: 0.4,
    y: 0.28,
    w: 7.2,
    h: 1,
    fontFace: "Segoe UI",
    fontSize: 29,
    color: "222222",
    margin: 0,
    breakLine: true,
  });

  // SECTION TAG
  slide12.addShape(pptx.ShapeType.roundRect, {
    x: 8.85,
    y: 0.25,
    w: 2.15,
    h: 0.38,
    rectRadius: 0.04,
    fill: { color: "7A7A7A" },
    line: { color: "7A7A7A" },
  });

  slide12.addText(positiveImpactData.section, {
    x: 8.9,
    y: 0.32,
    w: 2.05,
    h: 0.15,
    fontFace: "Segoe UI",
    fontSize: 10,
    bold: true,
    color: "FFFFFF",
    align: "center",
    margin: 0,
  });

  // QUOTES
  let yPoss = 1.7;

  positiveImpactData.quotes.forEach((quote) => {
    slide12.addText(`• ${quote}`, {
      x: 0.52,
      y: yPoss,
      w: 11.1,
      h: 0.32,
      fontFace: "Segoe UI",
      fontSize: 14,
      color: "333333",
      margin: 0,
      breakLine: true,
      valign: "top",
    });

    yPoss += 0.5;
  });

  addFooterToSlide(slide12, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideLeadershipImpact(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slide13 = pptx.addSlide();
  slide13.background = { color: "F2F2F2" };
  addLogoToSlide(slide13, false);

  // TITLE
  slide13.addText(`“${leadershipImpactData.title}”`, {
    x: 0.65,
    y: 0.22,
    w: 7.8,
    h: 1,
    fontFace: "Segoe UI",
    fontSize: 29,
    color: "222222",
    margin: 0,
    breakLine: true,
  });

  // SECTION TAG
  slide13.addShape(pptx.ShapeType.roundRect, {
    x: 8.85,
    y: 0.25,
    w: 2.15,
    h: 0.38,
    rectRadius: 0.04,
    fill: { color: "7A7A7A" },
    line: { color: "7A7A7A" },
  });

  slide13.addText(leadershipImpactData.section, {
    x: 8.9,
    y: 0.32,
    w: 2.05,
    h: 0.15,
    fontFace: "Segoe UI",
    fontSize: 10,
    bold: true,
    color: "FFFFFF",
    align: "center",
    margin: 0,
  });

  // QUOTES
  let yPo = 1.55;

  leadershipImpactData.quotes.forEach((quote) => {
    slide13.addText(`• ${quote}`, {
      x: 0.45,
      y: yPo,
      w: 11.1,
      h: 0.32,
      fontFace: "Segoe UI",
      fontSize: 14,
      color: "333333",
      margin: 0,
      breakLine: true,
      valign: "top",
    });

    yPo += 0.52;
  });

  addFooterToSlide(slide13, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlidePDP(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slide14 = pptx.addSlide();
  slide14.background = { color: "F2F2F2" };
  addLogoToSlide(slide14, false);

  // TITLE
  slide14.addText(`“${personalDevelopmentPlanData.title}”`, {
    x: 0.65,
    y: 0.22,
    w: 7.8,
    h: 1,
    fontFace: "Segoe UI",
    fontSize: 29,
    color: "222222",
    margin: 0,
    breakLine: true,
  });

  // SECTION TAG
  slide14.addShape(pptx.ShapeType.roundRect, {
    x: 9.15,
    y: 0.25,
    w: 2.15,
    h: 0.38,
    rectRadius: 0.04,
    fill: { color: "C40000" },
    line: { color: "C40000" },
  });

  slide14.addText(personalDevelopmentPlanData.section, {
    x: 9.2,
    y: 0.32,
    w: 2.05,
    h: 0.15,
    fontFace: "Segoe UI",
    fontSize: 10,
    bold: true,
    color: "FFFFFF",
    align: "center",
    margin: 0,
  });

  // PROFESSIONAL GOALS HEADER
  slide14.addShape(pptx.ShapeType.rect, {
    x: 0.42,
    y: 1.25,
    w: 11.2,
    h: 0.3,
    fill: { color: "102B46" },
    line: { color: "102B46" },
  });

  slide14.addText(
    [
      {
        text: "Professional Goals ",
        options: { bold: true }
      },
      {
        text:
          "What outcomes do you want to deliver over the next 12 months?",
        options: { italic: true }
      }
    ],
    {
      x: 0.5,
      y: 1.3,
      w: 7,
      h: 0.15,
      fontFace: "Segoe UI",
      fontSize: 10,
      color: "FFFFFF",
      margin: 0,
    }
  );

  // PROFESSIONAL GOALS TABLE
  slide14.addTable(
    [
      [
        {
          text: "Within your BU or Function",
          options: { bold: true }
        },
        {
          text: "Across the Organization",
          options: { bold: true }
        }
      ],
      [
        {
          text: `1. ${personalDevelopmentPlanData.professionalGoals.withinBU[0].text}`
        },
        {
          text: `1. ${personalDevelopmentPlanData.professionalGoals.acrossOrganization[0].text}`
        }
      ],
      [
        {
          text: `2. ${personalDevelopmentPlanData.professionalGoals.withinBU[1].text}`
        },
        {
          text: `2. ${personalDevelopmentPlanData.professionalGoals.acrossOrganization[1].text}`
        }
      ],
      [
        {
          text: `3. ${personalDevelopmentPlanData.professionalGoals.withinBU[2].text}`
        },
        {
          text: `3. ${personalDevelopmentPlanData.professionalGoals.acrossOrganization[2].text}`
        }
      ]
    ],
    {
      x: 0.42,
      y: 1.55,
      w: 11.2,
      h: 1,
      border: {
        type: "solid",
        pt: 1,
        color: "8A8A8A"
      },
      fontFace: "Segoe UI",
      fontSize: 10,
      color: "333333",
      fill: {
        color: "F2F2F2"
      },
      margin: 0.08,
      colW: [5.5, 5.7]
    }
  );

  // BEHAVIOR DEVELOPMENT HEADER
  slide14.addShape(pptx.ShapeType.rect, {
    x: 0.42,
    y: 3.05,
    w: 11.2,
    h: 0.3,
    fill: { color: "102B46" },
    line: { color: "102B46" },
  });

  slide14.addText(
    [
      {
        text: "Behavior Development ",
        options: { bold: true }
      },
      {
        text:
          "Pick two or three behaviors to build. What does success look like, and how will you know?",
        options: { italic: true }
      }
    ],
    {
      x: 0.5,
      y: 3.1,
      w: 9,
      h: 0.15,
      fontFace: "Segoe UI",
      fontSize: 10,
      color: "FFFFFF",
      margin: 0,
    }
  );

  // BEHAVIOR TABLE
  slide14.addTable(
    [
      personalDevelopmentPlanData.behaviorDevelopment.headers,
      ...personalDevelopmentPlanData.behaviorDevelopment.rows
    ],
    {
      x: 0.42,
      y: 3.35,
      w: 11.2,
      h: 1.3,
      border: {
        type: "solid",
        pt: 1,
        color: "8A8A8A"
      },
      fontFace: "Segoe UI",
      fontSize: 10,
      color: "333333",
      fill: {
        color: "F2F2F2"
      },
      margin: 0.08,
      colW: [1.65, 3.6, 5.95]
    }
  );

  // ACTION PLAN HEADER
  slide14.addShape(pptx.ShapeType.rect, {
    x: 0.42,
    y: 5.05,
    w: 5.6,
    h: 0.3,
    fill: { color: "102B46" },
    line: { color: "102B46" },
  });

  slide14.addText(
    [
      {
        text: "Action Plan ",
        options: { bold: true }
      },
      {
        text:
          "What will you do in the next 30 days to start building these behaviors?",
        options: { italic: true }
      }
    ],
    {
      x: 0.5,
      y: 5.1,
      w: 5.2,
      h: 0.15,
      fontFace: "Segoe UI",
      fontSize: 9,
      color: "FFFFFF",
      margin: 0,
    }
  );

  // STAKEHOLDER HEADER
  slide14.addShape(pptx.ShapeType.rect, {
    x: 6.02,
    y: 5.05,
    w: 5.6,
    h: 0.3,
    fill: { color: "102B46" },
    line: { color: "102B46" },
  });

  slide14.addText(
    [
      {
        text: "Stakeholders & Support ",
        options: { bold: true }
      },
      {
        text:
          "Who do you need to engage, and what help will you ask for?",
        options: { italic: true }
      }
    ],
    {
      x: 6.12,
      y: 5.1,
      w: 5.1,
      h: 0.15,
      fontFace: "Segoe UI",
      fontSize: 9,
      color: "FFFFFF",
      margin: 0,
    }
  );

  // ACTION / STAKEHOLDER TABLE
  slide14.addTable(
    [
      [
        personalDevelopmentPlanData.actionPlan.items[0],
        personalDevelopmentPlanData.stakeholders.items[0]
      ]
    ],
    {
      x: 0.42,
      y: 5.35,
      w: 11.2,
      h: 0.5,
      border: {
        type: "solid",
        pt: 1,
        color: "8A8A8A"
      },
      fontFace: "Segoe UI",
      fontSize: 10,
      color: "333333",
      fill: {
        color: "F2F2F2"
      },
      margin: 0.08,
      colW: [5.6, 5.6]
    }
  );

  addFooterToSlide(slide14, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

// THEME CARD FUNCTION
const addThemeCard = (
  pptx: any,
  slide: any,
  theme: any,
  x: number,
  y: number
) => {
  // Card background
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w: 5.35,
    h: 2.35,
    rectRadius: 0.06,
    fill: { color: "D9D9DF" },
    line: { color: "D9D9DF" },
  });

  // Header
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w: 5.35,
    h: 0.55,
    fill: { color: theme.headerColor },
    line: { color: theme.headerColor },
  });

  // Theme title
  slide.addText(theme.title, {
    x: x + 0.2,
    y: y + 0.16,
    w: 2.7,
    h: 0.2,
    fontFace: "Segoe UI",
    fontSize: 13,
    bold: true,
    color: "FFFFFF",
    margin: 0,
  });

  // Rating pill
  slide.addShape(pptx.ShapeType.roundRect, {
    x: x + 3.45,
    y: y + 0.12,
    w: 1.75,
    h: 0.28,
    rectRadius: 0.08,
    fill: { color: "FFFFFF" },
    line: { color: "FFFFFF" },
  });

  slide.addText(theme.rating, {
    x: x + 3.55,
    y: y + 0.18,
    w: 1.5,
    h: 0.1,
    fontFace: "Segoe UI",
    fontSize: 9,
    bold: true,
    color: theme.headerColor,
    align: "center",
    margin: 0,
  });

  // Linked behaviors
  slide.addText(theme.linkedBehaviors, {
    x: x + 0.22,
    y: y + 0.68,
    w: 4.9,
    h: 0.35,
    fontFace: "Segoe UI",
    fontSize: 7,
    italic: true,
    color: "6B7280",
    margin: 0,
  });

  // Quote 1 bar
  slide.addShape(pptx.ShapeType.rect, {
    x: x + 0.18,
    y: y + 1.02,
    w: 0.04,
    h: 0.42,
    fill: { color: theme.headerColor },
    line: { color: theme.headerColor },
  });

  // Quote 1
  slide.addText(theme.quotes[0], {
    x: x + 0.32,
    y: y + 1.05,
    w: 4.7,
    h: 0.45,
    fontFace: "Segoe UI",
    fontSize: 11,
    italic: true,
    color: "333333",
    margin: 0,
    breakLine: true,
  });

  // Quote 2 bar
  slide.addShape(pptx.ShapeType.rect, {
    x: x + 0.18,
    y: y + 1.65,
    w: 0.04,
    h: 0.42,
    fill: { color: theme.headerColor },
    line: { color: theme.headerColor },
  });

  // Quote 2
  slide.addText(theme.quotes[1], {
    x: x + 0.32,
    y: y + 1.68,
    w: 4.7,
    h: 0.45,
    fontFace: "Segoe UI",
    fontSize: 11,
    italic: true,
    color: "333333",
    margin: 0,
    breakLine: true,
  });
};

export function createSlideVerbatimThemes(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slide15 = pptx.addSlide();
  slide15.background = { color: "F2F2F2" };
  addLogoToSlide(slide15, false);

  // TITLE
  slide15.addText(verbatimThemesData.title, {
    x: 0.25,
    y: 0.18,
    w: 4,
    h: 0.5,
    fontFace: "Segoe UI",
    fontSize: 28,
    color: "222222",
    margin: 0,
  });

  // SECTION TAG
  slide15.addShape(pptx.ShapeType.roundRect, {
    x: 8.85,
    y: 0.25,
    w: 2.15,
    h: 0.38,
    rectRadius: 0.04,
    fill: { color: "C40000" },
    line: { color: "C40000" },
  });

  slide15.addText(verbatimThemesData.section, {
    x: 8.9,
    y: 0.32,
    w: 2.05,
    h: 0.15,
    fontFace: "Segoe UI",
    fontSize: 10,
    bold: true,
    color: "FFFFFF",
    align: "center",
    margin: 0,
  });

  // CARDS
addThemeCard(pptx, slide15, verbatimThemesData.themes[0], 0.35, 1.2);
addThemeCard(pptx, slide15, verbatimThemesData.themes[1], 6.18, 1.2);

addThemeCard(pptx, slide15, verbatimThemesData.themes[2], 0.35, 3.65);
addThemeCard(pptx, slide15, verbatimThemesData.themes[3], 6.18, 3.65);
  // FOOTNOTE
  slide15.addText(verbatimThemesData.footerNote, {
    x: 2.25,
    y: 6.1,
    w: 7,
    h: 0.2,
    fontFace: "Segoe UI",
    fontSize: 7,
    italic: true,
    color: "6B7280",
    align: "center",
    margin: 0,
  });

  addFooterToSlide(slide15, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

// QUADRANT FUNCTION
const addQuadrant = (
  pptx: any,
  slide: any,
  quadrant: any,
  x: number,
  y: number,
  w: number,
  h: number
) => {
  // Background
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w,
    h,
    fill: { color: quadrant.bgColor },
    line: { color: "BFBFBF", pt: 1 },
  });

  // Title
  slide.addText(quadrant.title, {
    x: x + 0.08,
    y: y + 0.12,
    w: w - 0.2,
    h: 0.2,
    fontFace: "Segoe UI",
    fontSize: 16,
    bold: false,
    color: "222222",
    margin: 0,
  });

  // Subtitle
  slide.addText(quadrant.subtitle, {
    x: x + 0.08,
    y: y + 0.48,
    w: w - 0.2,
    h: 0.35,
    fontFace: "Segoe UI",
    fontSize: 11,
    bold: true,
    color: "0066E6",
    margin: 0,
    breakLine: true,
  });

  // Bullet Items
  let itemY = y + 1.0;

  quadrant.items.forEach((item: string) => {
    slide.addText(`• ${item}`, {
      x: x + 0.08,
      y: itemY,
      w: w - 0.3,
      h: 0.28,
      fontFace: "Segoe UI",
      fontSize: 11,
      color: "222222",
      margin: 0,
      breakLine: true,
    });

    itemY += 0.42;
  });
};

export function createSlideCoachingQuadrant(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slide16 = pptx.addSlide();
  slide16.background = { color: "F2F2F2" };
  addLogoToSlide(slide16, false);

  // TITLE
  slide16.addText(coachingQuadrantData.title, {
    x: 0.38,
    y: 0.15,
    w: 8.0,
    h: 0.6,
    fontFace: "Segoe UI Semibold",
    fontSize: 32,
    color: "222222",
    margin: 0,
  });

  // SECTION TAG
  slide16.addShape(pptx.ShapeType.roundRect, {
    x: 9.35,
    y: 0.25,
    w: 2.15,
    h: 0.38,
    rectRadius: 0.04,
    fill: { color: "C40000" },
    line: { color: "C40000" },
  });

  slide16.addText(coachingQuadrantData.section, {
    x: 9.4,
    y: 0.32,
    w: 2.05,
    h: 0.15,
    fontFace: "Segoe UI",
    fontSize: 10,
    bold: true,
    color: "FFFFFF",
    align: "center",
    margin: 0,
  });

  // QUADRANTS (Widened and perfectly centered to form a continuous grid)
  const GRID_X = 0.76;
  const GRID_Y = 1.35;
  const QUAD_W = 4.68;
  const QUAD_H = 2.3;

  addQuadrant(pptx, slide16, coachingQuadrantData.quadrants[0], GRID_X, GRID_Y, QUAD_W, QUAD_H);
  addQuadrant(pptx, slide16, coachingQuadrantData.quadrants[1], GRID_X + QUAD_W, GRID_Y, QUAD_W, QUAD_H);
  addQuadrant(pptx, slide16, coachingQuadrantData.quadrants[2], GRID_X, GRID_Y + QUAD_H, QUAD_W, QUAD_H);
  addQuadrant(pptx, slide16, coachingQuadrantData.quadrants[3], GRID_X + QUAD_W, GRID_Y + QUAD_H, QUAD_W, QUAD_H);

  // SIDE LABELS (Perfectly centered vertically within each row, positioned horizontally on the left margin)
  slide16.addText("You don’t consider consistent", {
    x: -0.55,
    y: 2.2,
    w: 2.0,
    h: 0.3,
    rotate: 270,
    fontFace: "Segoe UI",
    fontSize: 9,
    color: "222222",
    align: "center",
    margin: 0,
  });

  slide16.addText("You consider consistent", {
    x: -0.55,
    y: 4.5,
    w: 2.0,
    h: 0.3,
    rotate: 270,
    fontFace: "Segoe UI",
    fontSize: 9,
    color: "222222",
    align: "center",
    margin: 0,
  });

  // BOTTOM LABELS (Perfectly centered under the two column boxes)
  slide16.addText("Colleagues see it consistently", {
    x: GRID_X,
    y: 6.05,
    w: QUAD_W,
    h: 0.3,
    fontFace: "Segoe UI",
    fontSize: 11,
    color: "222222",
    align: "center",
    margin: 0,
  });

  slide16.addText("Colleagues don’t see it consistently", {
    x: GRID_X + QUAD_W,
    y: 6.05,
    w: QUAD_W,
    h: 0.3,
    fontFace: "Segoe UI",
    fontSize: 11,
    color: "222222",
    align: "center",
    margin: 0,
  });

  addFooterToSlide(slide16, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}export function createSlideResultsAtGlance(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slide17 = pptx.addSlide();
  slide17.background = { color: "FFFFFF" };
  addLogoToSlide(slide17, false);

  // ── Layout constants ───────────────────────────────────────────────────────
  const TABLE_X     = 0.38;
  const TABLE_Y     = 1.45;
  const HEADER_H    = 0.35;
  const TEAL        = "1B7F83"; // Premium corporate teal
  const BORDER      = "BFDCDC";
  const BORDER_PT   = 0.75;

  const PART_W      = 2.60;
  const SEC_W       = 2.30;
  const BULL_W      = 4.40;
  const ACT_W       = 3.10;

  const X_PART      = TABLE_X;           // 0.38
  const X_FEED      = X_PART + PART_W + 0.18; // 3.16 (spacing gap of 0.18)
  const X_BULL      = X_FEED + SEC_W;    // 5.46
  const X_ACT       = X_BULL + BULL_W;   // 9.86
  const PART_BOX_Y  = TABLE_Y + HEADER_H; // 1.80

  // Row heights: sized to make it extremely spacious and elegant
  const rowHeights  = [1.15, 1.15, 1.20, 1.30];
  const TOTAL_ROW_H = rowHeights.reduce((a, b) => a + b, 0); // 4.80
  const PART_BOX_H  = TOTAL_ROW_H;

  // ── TITLE & SUBTITLE ───────────────────────────────────────────────────────
  // Title part 1: "Your results "
  slide17.addText("Your results ", {
    x: 0.38, y: 0.20, w: 2.6, h: 0.6,
    fontFace: "Segoe UI", fontSize: 36,
    color: "222222", bold: false, margin: 0,
  });

  // Title part 2: "at a glance"
  slide17.addText("at a glance", {
    x: 2.98, y: 0.20, w: 2.5, h: 0.6,
    fontFace: "Segoe UI", fontSize: 36,
    color: "222222", bold: false, margin: 0,
  });

  // // Custom dotted underline under "at a glance"
  // slide17.addShape(pptx.ShapeType.line, {
  //   x: 2.98, y: 0.72, w: 2.05, h: 0,
  //   line: { color: "B8AA97", pt: 1.5, dashType: "dash" }
  // });

  // Subtitle
  slide17.addText(resultsAtGlanceData.subtitle, {
    x: 0.38, y: 0.95, w: 7.0, h: 0.25,
    fontFace: "Segoe UI", fontSize: 13,
    color: "1B365D", margin: 0,
  });

  // ── BLUE BADGE (Top Right) ─────────────────────────────────────────────────
  slide17.addShape(pptx.ShapeType.roundRect, {
    x: 6.20, y: 0.30, w: 2.20, h: 0.50,
    rectRadius: 0.10,
    fill: { color: "0F62FE" }, line: { color: "0F62FE" },
  });
  slide17.addText(resultsAtGlanceData.sectionLeft, {
    x: 6.20, y: 0.30, w: 2.20, h: 0.50,
    fontFace: "Segoe UI", fontSize: 9, bold: true,
    color: "FFFFFF", align: "center", valign: "middle",
    margin: 0, breakLine: true,
  });

  // ── GREY BADGE (Top Right) ─────────────────────────────────────────────────
  slide17.addShape(pptx.ShapeType.roundRect, {
    x: 8.60, y: 0.30, w: 2.20, h: 0.50,
    rectRadius: 0.10,
    fill: { color: "5E5E5E" }, line: { color: "5E5E5E" },
  });
  slide17.addText(resultsAtGlanceData.sectionRight, {
    x: 8.60, y: 0.30, w: 2.20, h: 0.50,
    fontFace: "Segoe UI", fontSize: 9, bold: true,
    color: "FFFFFF", align: "center", valign: "middle",
    margin: 0, breakLine: true,
  });

  // ── HEADER ROW ─────────────────────────────────────────────────────────────
  // Left Column Header
  slide17.addShape(pptx.ShapeType.rect, {
    x: X_PART, y: TABLE_Y, w: PART_W, h: HEADER_H,
    fill: { color: TEAL }, line: { color: TEAL },
  });
  slide17.addText("Participant information", {
    x: X_PART + 0.12, y: TABLE_Y, w: PART_W - 0.24, h: HEADER_H,
    fontFace: "Segoe UI", fontSize: 11, bold: true,
    color: "FFFFFF", valign: "middle", margin: 0,
  });

  // Right Column Header 1: What the feedback suggests (spans both title and bullets)
  slide17.addShape(pptx.ShapeType.rect, {
    x: X_FEED, y: TABLE_Y, w: SEC_W + BULL_W, h: HEADER_H,
    fill: { color: TEAL }, line: { color: TEAL },
  });
  slide17.addText("What the feedback suggests", {
    x: X_FEED + 0.12, y: TABLE_Y, w: SEC_W + BULL_W - 0.24, h: HEADER_H,
    fontFace: "Segoe UI", fontSize: 11, bold: true,
    color: "FFFFFF", valign: "middle", margin: 0,
  });

  // Right Column Header 2: What to do
  slide17.addShape(pptx.ShapeType.rect, {
    x: X_ACT, y: TABLE_Y, w: ACT_W, h: HEADER_H,
    fill: { color: TEAL }, line: { color: TEAL },
  });
  slide17.addText("What to do", {
    x: X_ACT + 0.12, y: TABLE_Y, w: ACT_W - 0.24, h: HEADER_H,
    fontFace: "Segoe UI", fontSize: 11, bold: true,
    color: "FFFFFF", valign: "middle", margin: 0,
  }); 

  // ── PARTICIPANT INFO CARD ──────────────────────────────────────────────────
  // Outer Border Box
  slide17.addShape(pptx.ShapeType.rect, {
    x: X_PART, y: PART_BOX_Y, w: PART_W, h: PART_BOX_H,
    fill: { color: "FFFFFF" }, line: { color: BORDER, pt: BORDER_PT },
  });

  // Label background (grey block on the left side of the participant info box)
  slide17.addShape(pptx.ShapeType.rect, {
    x: X_PART, y: PART_BOX_Y, w: 1.00, h: 1.60,
    fill: { color: "F4F4F6" }, line: { color: "F4F4F6" },
  });

  const INFO_LABEL_X  = X_PART + 0.10;
  const INFO_VALUE_X  = X_PART + 1.12;
  const INFO_LABEL_W  = 0.85;
  const INFO_VALUE_W  = 1.35;
  const INFO_DIV_X    = X_PART + 1.00;

  const infoRows = [
    { label: "Name",               value: resultsAtGlanceData.participantInfo.name,            rowH: 0.45 },
    { label: "Reporting\nperiod",  value: resultsAtGlanceData.participantInfo.reportingPeriod,  rowH: 0.55 },
    { label: "Response\nrate",     value: resultsAtGlanceData.participantInfo.responseRate,     rowH: 0.50 },
  ];

  let infoY = PART_BOX_Y + 0.05;
  infoRows.forEach(({ label, value, rowH }, i) => {
    // Subtle horizontal divider between info rows
    if (i > 0) {
      slide17.addShape(pptx.ShapeType.line, {
        x: X_PART, y: infoY - 0.03, w: PART_W, h: 0,
        line: { color: BORDER, pt: BORDER_PT },
      });
    }

    slide17.addText(label, {
      x: INFO_LABEL_X, y: infoY, w: INFO_LABEL_W, h: rowH,
      fontFace: "Segoe UI", fontSize: 9.5,
      color: "333333", breakLine: true, valign: "middle", margin: 0,
    });
    slide17.addText(value, {
      x: INFO_VALUE_X, y: infoY, w: INFO_VALUE_W, h: rowH,
      fontFace: "Segoe UI", fontSize: 10,
      bold: true, color: "111111", valign: "middle", margin: 0,
    });

    infoY += rowH + 0.05;
  });

  // Vertical line separator between label column and value column
  slide17.addShape(pptx.ShapeType.line, {
    x: INFO_DIV_X, y: PART_BOX_Y, w: 0, h: 1.60,
    line: { color: BORDER, pt: BORDER_PT },
  });

  // Horizontal line separating info rows from confidentiality text
  const DIVIDER_Y = PART_BOX_Y + 1.60;
  slide17.addShape(pptx.ShapeType.line, {
    x: X_PART, y: DIVIDER_Y, w: PART_W, h: 0,
    line: { color: BORDER, pt: BORDER_PT },
  });

  // Confidential & Anonymity texts below info rows
  slide17.addText(resultsAtGlanceData.participantInfo.confidentialityText, {
    x: INFO_LABEL_X, y: DIVIDER_Y + 0.18, w: PART_W - 0.20, h: 1.35,
    fontFace: "Segoe UI", fontSize: 7.8, color: "666666",
    breakLine: true, margin: 0, valign: "top",
  });

  slide17.addText(resultsAtGlanceData.participantInfo.anonymityText, {
    x: INFO_LABEL_X, y: DIVIDER_Y + 1.65, w: PART_W - 0.20, h: 1.35,
    fontFace: "Segoe UI", fontSize: 7.8, color: "666666",
    breakLine: true, margin: 0, valign: "top",
  });

  // ── FEEDBACK & ACTION ROWS ─────────────────────────────────────────────────
  let currentY = PART_BOX_Y;

  resultsAtGlanceData.feedbackSections.forEach((section, index) => {
    const rowH = rowHeights[index];

    // Cell 1: Section Title
    slide17.addShape(pptx.ShapeType.rect, {
      x: X_FEED, y: currentY, w: SEC_W, h: rowH,
      fill: { color: "FFFFFF" }, line: { color: BORDER, pt: BORDER_PT },
    });

    // Cell 2: Bullets
    slide17.addShape(pptx.ShapeType.rect, {
      x: X_BULL, y: currentY, w: BULL_W, h: rowH,
      fill: { color: "FFFFFF" }, line: { color: BORDER, pt: BORDER_PT },
    });

    // Cell 3: Action (What to do)
    // Row 3 and Row 4 are merged for this column!
    if (index === 2) {
      // Draw a merged cell spanning Row 3 + Row 4
      const mergedH = rowHeights[2] + rowHeights[3];
      slide17.addShape(pptx.ShapeType.rect, {
        x: X_ACT, y: currentY, w: ACT_W, h: mergedH,
        fill: { color: "FFFFFF" }, line: { color: BORDER, pt: BORDER_PT },
      });

      // Combine Row 3 and Row 4 action texts
      const combinedAction = section.action + "\n\n" + resultsAtGlanceData.feedbackSections[3].action;
      slide17.addText(combinedAction, {
        x: X_ACT + 0.12, y: currentY + 0.12,
        w: ACT_W - 0.24, h: mergedH - 0.24,
        fontFace: "Segoe UI", fontSize: 9.5,
        color: "333333", breakLine: true, valign: "top", margin: 0,
      });
    } else if (index < 2) {
      // Row 1 and Row 2 get regular action cells
      slide17.addShape(pptx.ShapeType.rect, {
        x: X_ACT, y: currentY, w: ACT_W, h: rowH,
        fill: { color: "FFFFFF" }, line: { color: BORDER, pt: BORDER_PT },
      });

      slide17.addText(section.action, {
        x: X_ACT + 0.12, y: currentY + 0.12,
        w: ACT_W - 0.24, h: rowH - 0.24,
        fontFace: "Segoe UI", fontSize: 9.5,
        color: "333333", breakLine: true, valign: "top", margin: 0,
      });
    }

    // Section Title Text
    slide17.addText(section.title, {
      x: X_FEED + 0.12, y: currentY + 0.12,
      w: SEC_W - 0.24, h: rowH - 0.24,
      fontFace: "Segoe UI", fontSize: 10.5, bold: true,
      color: "20354B", breakLine: true, valign: "top", margin: 0,
    });

    // Custom Styled Bullets (Teal bullet point + Navy body text)
    let bulletY = currentY + 0.12;
    const bulletSpacing = 0.32;
    section.points.forEach((point) => {
      slide17.addText(
        [
          { text: "•  ", options: { color: "1B7F83", bold: true } },
          { text: point, options: { color: "20354B" } }
        ],
        {
          x: X_BULL + 0.12, y: bulletY,
          w: BULL_W - 0.24, h: 0.28,
          fontFace: "Segoe UI", fontSize: 10,
          margin: 0, valign: "top",
        }
      );
      bulletY += bulletSpacing;
    });

    currentY += rowH;
  });

  addFooterToSlide(slide17, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}
export function createSlideAppendixDivider(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slide18 = pptx.addSlide();
  slide18.background = { color: "111111" };
  addLogoToSlide(slide18, true);

  // SECTION TAG
  slide18.addShape(pptx.ShapeType.roundRect, {
    x: 8.62,
    y: 0.28,
    w: 2.2,
    h: 0.38,
    rectRadius: 0.04,
    fill: { color: "D40000" },
    line: { color: "D40000" },
  });

  slide18.addText("DRAFT FOR INPUT", {
    x: 8.72,
    y: 0.36,
    w: 2.0,
    h: 0.12,
    fontFace: "Segoe UI",
    fontSize: 10,
    bold: true,
    color: "FFFFFF",
    align: "center",
    margin: 0,
  });

  // TITLE
  slide18.addText("APPENDIX", {
    x: 0.55,
    y: 3.65,
    w: 3,
    h: 0.45,
    fontFace: "Segoe UI Light",
    fontSize: 32,
    color: "FFFFFF",
    margin: 0,
    bold: true,
  });

  addFooterToSlide(slide18, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideSelectiveOutputsDivider(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slide19 = pptx.addSlide();
  slide19.background = { color: "111111" };
  addLogoToSlide(slide19, true);

  // SECTION TAG
  slide19.addShape(pptx.ShapeType.roundRect, {
    x: 8.62,
    y: 0.32,
    w: 2.2,
    h: 0.38,
    rectRadius: 0.04,
    fill: { color: "D40000" },
    line: { color: "D40000" },
  });

  slide19.addText("DRAFT FOR INPUT", {
    x: 8.72,
    y: 0.40,
    w: 2.0,
    h: 0.12,
    fontFace: "Segoe UI",
    fontSize: 10,
    bold: true,
    color: "FFFFFF",
    align: "center",
    margin: 0,
  });

  // TITLE
  slide19.addText("SELECTIVE ADDITIONAL OUTPUTS", {
    x: 0.55,
    y: 3.78,
    w: 7.5,
    h: 0.5,
    fontFace: "Segoe UI Light",
    fontSize: 30,
    color: "FFFFFF",
    margin: 0,
  });

  addFooterToSlide(slide19, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
}

export function createSlideEnd(params: SlideParams) {
  const { pptx, participantName, footerDate, pageNumber } = params;
  const slide20 = pptx.addSlide();
  slide20.background = { color: "111111" };
  addLogoToSlide(slide20, true);

  // TITLE
  slide20.addText("END", {
    x: 0.55,
    y: 5.05,
    w: 2,
    h: 0.45,
    fontFace: "Segoe UI Light",
    fontSize: 30,
    color: "FFFFFF",
    margin: 0,
  });

  addFooterToSlide(slide20, pptx, participantName, footerDate, pageNumber, 10.80, 7.10, 2.30, 0.28);
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
          participant: true,
          answers: { include: { option: true } },
        },
      },
    },
  });

  if (!form) {
    return Response.json({ error: "Form not found" }, { status: 404 });
  }

  //////////////////////////////////////////////////////
  // GLOBAL FOOTER VARIABLES
  //////////////////////////////////////////////////////

  const participantName = form.responses[0]?.participant?.fullName || "PARTICIPANT";
  const footerDate = new Date().toLocaleDateString();

  //////////////////////////////////////////////////////
  // RESPONSE POOLS
  //////////////////////////////////////////////////////

  const filteredResponses = participantId
    ? form.responses.filter(
      (r: any) => r.participantId === participantId
    )
    : form.responses;

  const allResponses = form.responses;

  const selfResponses = filteredResponses.filter(
    (r: any) => !!r.participant?.email && r.email.toLowerCase() === r.participant.email.toLowerCase()
  );

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

  // Diagram Base64
  const diagramPath = path.join(process.cwd(), "public", "360_diagram.png");
  let diagramBase64 = "";
  try {
    diagramBase64 = fs.readFileSync(diagramPath).toString("base64");
  } catch (err) {
    console.error("Error reading 360 diagram image:", err);
  }

  const baseParams = {
    pptx,
    participantName,
    footerDate,
    analytics,
    selfAnalytics,
    othersAnalytics,
    enterpriseAnalytics,
    diagramBase64,
  };

  let pageNumber = 1;

  createSlideIntro({ ...baseParams, pageNumber: pageNumber++ });
  createSlideHowCreated({ ...baseParams, pageNumber: pageNumber++ });
  createSlideHowToUse({ ...baseParams, pageNumber: pageNumber++ });
  createSlideResultsAtGlance({ ...baseParams, pageNumber: pageNumber++ });
  createSlideOverallResults({ ...baseParams, pageNumber: pageNumber++ });
  createSlideSelfOther({ ...baseParams, pageNumber: pageNumber++ });
  createSlideQualitativeFeedback({ ...baseParams, pageNumber: pageNumber++ });
  createSlideAppendixDivider({ ...baseParams, pageNumber: pageNumber++ });
  createSlideResultsDistribution({ ...baseParams, pageNumber: pageNumber++ });
  createSlideInsufficientExposures({ ...baseParams, pageNumber: pageNumber++ });
  createSlideVerbatimQuotes({ ...baseParams, pageNumber: pageNumber++ });
  createSlidePositiveImpact({ ...baseParams, pageNumber: pageNumber++ });
  createSlideLeadershipImpact({ ...baseParams, pageNumber: pageNumber++ });
  createSlidePDP({ ...baseParams, pageNumber: pageNumber++ });
  createSlideScoreCalc({ ...baseParams, pageNumber: pageNumber++ });
  createSlideFAQ({ ...baseParams, pageNumber: pageNumber++ });
  createSlideSelectiveOutputsDivider({ ...baseParams, pageNumber: pageNumber++ });

  createSlideVerbatimThemes({ ...baseParams, pageNumber: pageNumber++ });
  createSlideCoachingQuadrant({ ...baseParams, pageNumber: pageNumber++ });



  createSlideEnd({ ...baseParams, pageNumber: pageNumber++ });

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

