import type { ChartKind, ChartSeriesStyle, ChartStyleInput, ChartStylePlan } from "../core/types";

export const CICC_SERIES_COLORS = [
  "#640000",
  "#B9B8A6",
  "#3D889A",
  "#BE995D",
  "#646C86",
  "#DD965D",
  "#44546A",
] as const;

export const cmToPoints = (centimeters: number): number => centimeters * 28.3464567;

const WHITE = "#FFFFFF";
const CHART_WIDTH_CM = 16;
const CHART_HEIGHT_CM = 9;
const AXIS_SIZE_POINTS = 8;
const LEGEND_SIZE_POINTS = 9;
const LINE_WIDTH_POINTS = 1.5;
const PLACEMENT_GUTTER_POINTS = 18;
const EXCEL_FINAL_COLUMN = 16384;

export function buildChartStylePlan(input: ChartStyleInput): ChartStylePlan {
  const placement = resolvePlacement(input);

  return {
    seriesColors: resolveSeriesColors(input.seriesCount),
    seriesStyle: resolveSeriesStyle(input.kind),
    widthPoints: cmToPoints(CHART_WIDTH_CM),
    heightPoints: cmToPoints(CHART_HEIGHT_CM),
    legendPosition: "bottom",
    legendOverlay: false,
    legendFontSizePoints: LEGEND_SIZE_POINTS,
    chartAreaFill: WHITE,
    plotAreaFill: WHITE,
    showOuterBorder: false,
    showTitle: false,
    showDataLabels: false,
    showGridlines: false,
    lineWidthPoints: LINE_WIDTH_POINTS,
    smoothLines: input.kind === "line" || input.kind === "lineMarkers" || input.kind === "lineStacked",
    textSizePoints: AXIS_SIZE_POINTS,
    valueAxisNumberFormat: resolveValueAxisNumberFormat(input.sourceFormat),
    categoryAxisNumberFormat: resolveCategoryAxisNumberFormat(input.categoryFormat),
    warnings: input.seriesCount > CICC_SERIES_COLORS.length ? ["series_palette_reused"] : [],
    placement,
  };
}

function resolveSeriesStyle(kind: ChartKind): ChartSeriesStyle {
  switch (kind) {
    case "column":
    case "columnStacked":
    case "columnStacked100":
    case "bar":
      return "fill-no-border";
    case "line":
    case "lineMarkers":
    case "lineStacked":
      return "line";
    case "scatterTrend":
      return "scatter";
    case "pie":
    case "pieExploded":
      return "pie-points";
  }
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
