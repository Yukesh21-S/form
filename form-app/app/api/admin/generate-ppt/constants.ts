// src/lib/constants.ts
export const COLORS = {
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

export const LAYOUT = {
  ROW_HEIGHT: 0.31,
  START_Y: 1.38,
  QUESTION_X: 0.52,
  QUESTION_W: 3.55,
  BAR_X: 4.15,
  BAR_W: 2.55,
  COUNT_START_X: 7.05,
  COUNT_W: 0.92,
};

export function getScoreColor(score: number) {
  if (score > 85) return COLORS.green;
  if (score >= 70) return COLORS.orange;
  return COLORS.red;
}
