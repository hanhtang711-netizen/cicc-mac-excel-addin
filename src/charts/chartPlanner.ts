import { AddinError } from "../core/errors";
import type { ChartKind, ChartOptions, ChartPlan, ParsedSelection } from "../core/types";
import { CHART_CATALOG } from "./chartCatalog";

export function buildChartPlan(
  parsed: ParsedSelection,
  kind: ChartKind,
  options: ChartOptions = {},
): ChartPlan {
  if ((kind === "pie" || kind === "pieExploded") && parsed.series.length !== 1) {
    throw new AddinError("pie_requires_one_series");
  }

  if (kind === "scatterTrend") {
    assertScatterData(parsed);
  }

  return {
    kind,
    excelType: CHART_CATALOG[kind].excelType,
    sourceAddress: parsed.sourceAddress,
    orientation: options.orientation ?? parsed.orientation,
    series: parsed.series.map((series) => ({
      name: series.name,
      categoryAddress: parsed.categoryAddress,
      valuesAddress: series.valuesAddress,
    })),
    ...(options.title ?? parsed.title ? { title: options.title ?? parsed.title } : {}),
    showDataLabels: options.showDataLabels ?? false,
    addLinearTrendline: kind === "scatterTrend" && options.addTrendline !== false,
  };
}

function assertScatterData(parsed: ParsedSelection): void {
  if (!isNumericOrDateFormat(categoryFormat(parsed))) {
    throw new AddinError("scatter_requires_numeric_x");
  }

  const hasNumericY = parsed.series.some((series, seriesIndex) =>
    isNumericFormat(seriesFormat(parsed, series.valueColumnOffset, seriesIndex)),
  );
  if (!hasNumericY) {
    throw new AddinError("scatter_requires_numeric_x");
  }
}

function categoryFormat(parsed: ParsedSelection): string | undefined {
  if (parsed.orientation === "columns") {
    return parsed.numberFormats[parsed.headerRows]?.[parsed.categoryColumnOffset];
  }

  const categoryRow = parsed.headerRows === 0 ? 0 : parsed.headerRows - 1;
  return parsed.numberFormats[categoryRow]?.[1];
}

function seriesFormat(
  parsed: ParsedSelection,
  valueColumnOffset: number,
  seriesIndex: number,
): string | undefined {
  if (parsed.orientation === "columns") {
    return parsed.numberFormats[parsed.headerRows]?.[valueColumnOffset];
  }

  return parsed.numberFormats[parsed.headerRows + seriesIndex]?.[valueColumnOffset];
}

function isNumericOrDateFormat(format: string | undefined): boolean {
  return !isTextFormat(format);
}

function isNumericFormat(format: string | undefined): boolean {
  return !isTextFormat(format) && !isDateFormat(format);
}

function isTextFormat(format: string | undefined): boolean {
  return format === undefined || /(^|[^\\])@|text/i.test(format);
}

function isDateFormat(format: string | undefined): boolean {
  if (format === undefined) {
    return false;
  }
  const withoutQuotedText = format.replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "");
  return /y{1,4}|d{1,4}|h{1,2}|s{1,2}|m{1,4}/i.test(withoutQuotedText);
}
