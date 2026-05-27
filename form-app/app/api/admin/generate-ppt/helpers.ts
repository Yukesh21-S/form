// src/lib/helpers.ts

// Helper functions for PPTX generation. No imports from route.tsx to avoid circular dependencies.

function getEricssonLogoSvg(color: string) {
  const hexColor = color.startsWith("#") ? color : `#${color}`;
  const dPath = "M20.76 1.593A2.36 2.36 0 0 0 19.572.225c-.54-.27-1.188-.336-2.256.02L5.187 4.29c-1.068.357-1.548.795-1.818 1.338a2.36 2.36 0 0 0 1.059 3.174c.54.27 1.188.336 2.256-.021l12.129-4.044c1.068-.354 1.548-.795 1.818-1.338a2.35 2.35 0 0 0 .13-1.806zm0 7.485a2.36 2.36 0 0 0-1.188-1.368c-.54-.27-1.188-.336-2.256.021L5.187 11.775c-1.068.357-1.548.795-1.818 1.338a2.36 2.36 0 0 0 1.059 3.174c.54.27 1.188.336 2.256-.021l12.129-4.041c1.068-.357 1.548-.795 1.818-1.341a2.35 2.35 0 0 0 .13-1.806zm0 7.488a2.36 2.36 0 0 0-1.188-1.368c-.54-.27-1.188-.336-2.256.021L5.187 19.263c-1.068.357-1.548.795-1.818 1.338a2.36 2.36 0 0 0 1.059 3.174c.54.27 1.188.336 2.256-.02l12.129-4.045c1.068-.354 1.548-.795 1.818-1.338a2.35 2.35 0 0 0 .13-1.806z";
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path fill="${hexColor}" d="${dPath}"/></svg>`;
}

function getEricssonLogoBase64(color: string) {
  const svg = getEricssonLogoSvg(color);
  return "data:image/svg+xml;base64," + Buffer.from(svg).toString("base64");
}

export function addLogoToSlide(slide: any, isDarkBg: boolean) {
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

export function addFooterToSlide(
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
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w,
    h,
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
