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
      x: barX, y: currentY - 0.02, w: 0, h: 0.38,
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

    slide3.addShape(pptx.ShapeType.line, {
      x: 0.15, y: tableY + 0.27, w: 7.1, h: 0,
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
          radarStyle: "Marker",
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
          radarStyle: "Marker",
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
  // SLIDE 5 — SELF vs OTHERS
  // Uses relationshipType === "SELF" for self responses
  // Uses all other types for "Others"
  // Only renders when there are both self and other responses
  //////////////////////////////////////////////////////

  const hasSelfData = selfAnalytics.some((q) => q.validResponses > 0);
  const hasOthersData = othersAnalytics.some((q) => q.validResponses > 0);

  if (hasSelfData || hasOthersData) {

    const slide5 = pptx.addSlide();
    slide5.background = { color: "FFFFFF" };

    slide5.addText("Self / other detail", {
      x: 0.38, y: 0.15, w: 5, h: 0.5,
      fontFace: "Segoe UI Semibold", fontSize: 32, color: "111111", margin: 0,
    });

    const LEGEND_Y = 0.25;
    const LEGEND_START_X = 5.2;
    const legendItems = [
      { text: "Colleagues see more often (> +10%)", color: "1D2951" },
      { text: "Broadly aligned (-10-10%)", color: "D9D9D9" },
      { text: "Colleagues see less often (< -10%)", color: "1D6FD8" },
    ];

    legendItems.forEach((item, idx) => {
      const x = LEGEND_START_X + idx * 2.3;
      slide5.addShape(pptx.ShapeType.rect, {
        x, y: LEGEND_Y, w: 0.15, h: 0.15,
        fill: { color: item.color }, line: { color: item.color },
      });
      slide5.addText(item.text, {
        x: x + 0.2, y: LEGEND_Y - 0.05, w: 2, h: 0.25,
        fontFace: "Segoe UI", fontSize: 7.5, color: "444444", margin: 0,
      });
    });

    slide5.addText("Shown as the difference from your self-rating", {
      x: 5.5, y: 0.65, w: 4, h: 0.2,
      fontFace: "Segoe UI", fontSize: 9, bold: true, color: "555555", align: "center",
    });

    const HEADER_Y = 0.9;
    const COL_RESULTS = 0.38;
    const COL_SELF = 3.2;
    const COL_MGR = 4.8;
    const COL_PEER = 5.9;
    const COL_DR = 7.0;
    const COL_OTH = 8.1;

    const rowStyle = { fontFace: "Segoe UI", fontSize: 8, bold: true, color: "000000", margin: 0 };

    slide5.addText("Results", { x: COL_RESULTS, y: HEADER_Y, w: 1, h: 0.2, ...rowStyle });
    slide5.addText("Self Rating", { x: COL_SELF + 0.4, y: HEADER_Y, w: 1, h: 0.2, ...rowStyle, align: "center" });
    slide5.addText("Manager", { x: COL_MGR, y: HEADER_Y, w: 1, h: 0.2, ...rowStyle, align: "center" });
    slide5.addText("Peers", { x: COL_PEER, y: HEADER_Y, w: 1, h: 0.2, ...rowStyle, align: "center" });
    slide5.addText("Direct Reports", { x: COL_DR, y: HEADER_Y, w: 1, h: 0.2, ...rowStyle, align: "center" });
    slide5.addText("Others", { x: COL_OTH, y: HEADER_Y, w: 1, h: 0.2, ...rowStyle, align: "center" });

    slide5.addShape(pptx.ShapeType.rect, {
      x: 0, y: HEADER_Y + 0.22, w: "100%", h: 0.25,
      fill: { color: "EFEFEF" }, line: { color: "EFEFEF" },
    });

    let currentY = 1.45;
    const ROW_OFFSET = 0.42; // Increased for better readability

    analytics.forEach((item, idx) => {
      // Question Text
      slide5.addText(item.question, {
        x: COL_RESULTS, y: currentY, w: 2.7, h: 0.35,
        fontFace: "Segoe UI", fontSize: 8, color: "111111", margin: 0, fit: "shrink",
        valign: "middle",
      });

      // Self Bar + Value
      const selfVal = selfAnalytics[idx].rawAverage || 0;
      const selfBarMaxW = 0.7;
      const selfBarW = (selfVal / 100) * selfBarMaxW;

      slide5.addShape(pptx.ShapeType.rect, {
        x: COL_SELF, y: currentY + 0.12, w: selfBarMaxW, h: 0.1,
        fill: { color: "EAEAEA" },
      });
      slide5.addShape(pptx.ShapeType.rect, {
        x: COL_SELF, y: currentY + 0.12, w: selfBarW, h: 0.1,
        fill: { color: "999999" },
      });
      slide5.addText(`${Math.round(selfVal)}%`, {
        x: COL_SELF + selfBarMaxW + 0.05, y: currentY + 0.09, w: 0.35, h: 0.15,
        fontFace: "Segoe UI", fontSize: 7.5, color: "333333", align: "left",
      });

      // Roles diff
      const roles = [
        { data: managerAnalytics[idx], x: COL_MGR },
        { data: peerAnalytics[idx], x: COL_PEER },
        { data: directReportAnalytics[idx], x: COL_DR },
        { data: otherRoleAnalytics[idx], x: COL_OTH },
      ];

      roles.forEach(role => {
        const centerPos = role.x + 0.45;

        if (role.data.validResponses > 0 && selfAnalytics[idx].validResponses > 0) {
          const diff = role.data.rawAverage - selfAnalytics[idx].rawAverage;
          const diffText = diff > 0 ? `+${Math.round(diff)}%` : `${Math.round(diff)}%`;

          let color = "D9D9D9";
          if (diff > 10) color = "1D2951";
          if (diff < -10) color = "1D6FD8";

          // Bar
          const barW = Math.abs(diff) / 100 * 0.4;
          slide5.addShape(pptx.ShapeType.rect, {
            x: diff >= 0 ? centerPos : centerPos - barW, y: currentY + 0.14, w: barW, h: 0.07,
            fill: { color: color }, line: { color: color },
          });

          // Text
          slide5.addText(diffText, {
            x: role.x, y: currentY, w: 0.9, h: 0.15,
            fontFace: "Segoe UI", fontSize: 7, color: "111111", align: "center",
          });
        }
      });

      // Grid line
      slide5.addShape(pptx.ShapeType.line, {
        x: COL_RESULTS, y: currentY + ROW_OFFSET - 0.01, w: 9.15, h: 0,
        line: { color: "F0F0F0", pt: 0.5 },
      });

      currentY += ROW_OFFSET;
    });
  }

  //////////////////////////////////////////////////////
  // SLIDE 6 — COACHING QUADRANT
  //////////////////////////////////////////////////////

  // const slide6 = pptx.addSlide();
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