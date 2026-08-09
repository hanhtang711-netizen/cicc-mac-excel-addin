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
    worksheetName: parsed.worksheetName,
    sourceAddress: parsed.sourceAddress,
    orientation: parsed.orientation,
    series: parsed.series.map((series) => ({
      name: series.name,
      categoryAddress: parsed.categoryAddress,
      valuesAddress: series.valuesAddress,
    })),
    addLinearTrendline: kind === "scatterTrend",
  };
}

function assertScatterData(parsed: ParsedSelection): void {
  const validX = parsed.categoryKinds.length > 0 &&
    parsed.categoryKinds.every((kind) => kind === "number" || kind === "date" || kind === "blank") &&
    parsed.categoryKinds.some((kind) => kind === "number" || kind === "date");
  if (!validX) {
    throw new AddinError("scatter_requires_numeric_x");
  }

  const validY = parsed.series.length > 0 &&
    parsed.series.every((series) =>
      series.valueKinds.every((kind) => kind === "number" || kind === "blank"),
    ) &&
    parsed.series.some((series) => series.valueKinds.some((kind) => kind === "number"));
  if (!validY) {
    throw new AddinError("scatter_requires_numeric_x");
  }
}
