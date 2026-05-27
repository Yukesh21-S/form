// src/lib/tableUtility.ts
import pptxgen from "pptxgenjs";
import { COLORS } from "./constants";

/**
 * Adds a data table to a slide.
 * Used for result distribution and insufficient exposure tables.
 */
export function addDataTable(
  slide: any,
  data: any[],
  config: {
    startX: number;
    startY: number;
    colWidths: number[];
    headerHeight: number;
    rowHeight: number;
    headerBg?: string;
    stripeColors?: string[];
  }
) {
  const { startX, startY, colWidths, headerHeight, rowHeight, headerBg, stripeColors } = config;
  const totalWidth = colWidths.reduce((a, b) => a + b, 0);
  const headerY = startY;

  // Header background
  if (headerBg) {
    slide.addShape(pptxgen.ShapeType.rect, {
      x: startX,
      y: headerY,
      w: totalWidth,
      h: headerHeight,
      fill: { color: headerBg },
    });
  }

  // Header texts (assumes data[0] contains header labels) – caller should add separately if needed.

  // Row stripes
  data.forEach((row, idx) => {
    const y = startY + headerHeight + idx * rowHeight;
    if (stripeColors && stripeColors[idx % stripeColors.length]) {
      slide.addShape(pptxgen.ShapeType.rect, {
        x: startX + colWidths[0] + colWidths[1], // start after first two columns (label + distribution)
        y,
        w: colWidths.slice(2).reduce((a, b) => a + b, 0),
        h: rowHeight,
        fill: { color: stripeColors[idx % stripeColors.length] },
        line: { pt: 0 },
      });
    }
    // Render cells – caller should add text/bars per column as needed.
  });
}
