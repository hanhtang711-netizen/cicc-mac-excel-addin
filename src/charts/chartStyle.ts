import type { ChartStyleInput, ChartStylePlan } from "../core/types";

export const CICC_SERIES_COLORS = [
  "#640000",
  "#8A2626",
  "#3D889A",
  "#646C86",
  "#BE995D",
  "#DD965D",
] as const;

export const SIZE_PRESETS_CM = {
  small: { width: 9, height: 5.5 },
  medium: { width: 11.5, height: 6.7 },
  large: { width: 15, height: 8.7 },
} as const;

export const cmToPoints = (centimeters: number): number => centimeters * 28.3464567;

const WHITE = "#FFFFFF";
const LIGHT_GRAY = "#D9D9D9";
const TEXT_SIZE_POINTS = 8;
const PLACEMENT_GUTTER_POINTS = 18;
const EXCEL_FINAL_COLUMN = 16384;

export function buildChartStylePlan(input: ChartStyleInput): ChartStylePlan {
  const { widthCm, heightCm } = resolveDimensions(input);
  const placement = resolvePlacement(input);

  return {
    seriesColors: resolveSeriesColors(input.seriesCount),
    widthPoints: cmToPoints(widthCm),
    heightPoints: cmToPoints(heightCm),
    legendPosition: input.options.legendPosition ?? "bottom",
    chartAreaFill: WHITE,
    plotAreaFill: WHITE,
    showOuterBorder: false,
    textSizePoints: TEXT_SIZE_POINTS,
    majorGridlineColor: LIGHT_GRAY,
    valueAxisNumberFormat: resolveValueAxisNumberFormat(input.sourceFormat),
    categoryAxisNumberFormat: resolveCategoryAxisNumberFormat(input.categoryFormat),
    warnings: input.seriesCount > CICC_SERIES_COLORS.length ? ["series_palette_reused"] : [],
    placement,
  };
}

function resolveDimensions(input: ChartStyleInput): { widthCm: number; heightCm: number } {
  if (input.options.sizePreset === "custom") {
    const { widthCm, heightCm } = input.options;
    if (!isPositiveFinite(widthCm) || !isPositiveFinite(heightCm)) {
      throw new Error("invalid_chart_dimensions");
    }
    return { widthCm, heightCm };
  }

  const sizePreset = input.options.sizePreset ?? "medium";
  const size = sizePreset === "small" || sizePreset === "large"
    ? SIZE_PRESETS_CM[sizePreset]
    : SIZE_PRESETS_CM.medium;
  return { widthCm: size.width, heightCm: size.height };
}

function resolveSeriesColors(seriesCount: number): string[] {
  return Array.from(
    { length: Math.max(0, seriesCount) },
    (_, index) => CICC_SERIES_COLORS[index % CICC_SERIES_COLORS.length],
  );
}

function resolveValueAxisNumberFormat(sourceFormat: string): string | undefined {
  return /^0(?:\.0{1,2})?%$/.test(sourceFormat) ? sourceFormat : undefined;
}

function resolveCategoryAxisNumberFormat(categoryFormat: string | undefined): string | undefined {
  return categoryFormat !== undefined && isDateLikeFormat(categoryFormat) ? "yyyy-mm-dd" : undefined;
}

function isDateLikeFormat(numberFormat: string): boolean {
  const formatWithoutLiterals = numberFormat
    .replace(/"(?:[^"]|"")*"/g, "")
    .replace(/\\./g, "");
  return /(?:^|[^a-z])(?:d+|m+|y+|h+|s+)(?:[^a-z]|$)/i.test(formatWithoutLiterals);
}

function resolvePlacement(input: ChartStyleInput): ChartStylePlan["placement"] {
  const selectionColumn = input.selectionColumn ?? 0;
  const selectionColumnCount = input.selectionColumnCount ?? 0;
  const estimatedChartColumns = input.estimatedChartColumns ?? 1;
  const placeBelow = selectionColumn + selectionColumnCount + estimatedChartColumns > EXCEL_FINAL_COLUMN;

  return {
    side: placeBelow ? "below" : "right",
    gutterPoints: PLACEMENT_GUTTER_POINTS,
  };
}

function isPositiveFinite(value: number | undefined): value is number {
  return value !== undefined && Number.isFinite(value) && value > 0;
}
