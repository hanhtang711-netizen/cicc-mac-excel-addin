import { describe, expect, it } from "vitest";
import {
  buildChartStylePlan,
  CICC_SERIES_COLORS,
  cmToPoints,
} from "../../src/charts/chartStyle";

describe("CICC chart style", () => {
  it("locks every line chart to the approved Skill contract", () => {
    const style = buildChartStylePlan({
      kind: "line",
      seriesCount: 2,
      sourceFormat: "0.0%",
      categoryFormat: "yyyy-mm-dd",
      selectionColumn: 0,
      selectionColumnCount: 3,
    });

    expect(style).toMatchObject({
      seriesStyle: "line",
      widthPoints: cmToPoints(16),
      heightPoints: cmToPoints(9),
      legendPosition: "bottom",
      legendOverlay: false,
      legendFontSizePoints: 9,
      showTitle: false,
      showDataLabels: false,
      showGridlines: false,
      lineWidthPoints: 1.5,
      smoothLines: true,
    });
  });

  it("cycles the seven approved chart colors", () => {
    const style = buildChartStylePlan({ kind: "line", seriesCount: 9, sourceFormat: "0.0%" });

    expect(style.seriesColors).toEqual([
      ...CICC_SERIES_COLORS,
      "#640000",
      "#B9B8A6",
    ]);
    expect(style.warnings).toContain("series_palette_reused");
  });

  it("uses the fixed dimensions and percentage precision", () => {
    const style = buildChartStylePlan({ kind: "column", seriesCount: 2, sourceFormat: "0.0%" });

    expect(style.widthPoints).toBeCloseTo(cmToPoints(16), 4);
    expect(style.heightPoints).toBeCloseTo(cmToPoints(9), 4);
    expect(style.valueAxisNumberFormat).toBe("0.0%");
    expect(style.legendPosition).toBe("bottom");
  });

  it.each([
    ["column", "fill-no-border"],
    ["bar", "fill-no-border"],
    ["lineMarkers", "line"],
    ["scatterTrend", "scatter"],
    ["pie", "pie-points"],
  ] as const)("maps %s to its type-specific series style", (kind, seriesStyle) => {
    expect(buildChartStylePlan({ kind, seriesCount: 1, sourceFormat: "General" }).seriesStyle).toBe(seriesStyle);
  });

  it("preserves supported percentage formats and normalizes date-like categories", () => {
    const style = buildChartStylePlan({
      seriesCount: 1,
      kind: "line",
      sourceFormat: "0.00%",
      categoryFormat: "m/d/yyyy",
    });

    expect(style.valueAxisNumberFormat).toBe("0.00%");
    expect(style.categoryAxisNumberFormat).toBe("yyyy-mm-dd");
  });

  it("does not treat quoted currency text as a date-like category format", () => {
    const style = buildChartStylePlan({
      seriesCount: 1,
      kind: "column",
      sourceFormat: "General",
      categoryFormat: "\"RMB\" #,##0",
    });

    expect(style.categoryAxisNumberFormat).toBeUndefined();
  });

  it("applies the fixed CICC base appearance", () => {
    const style = buildChartStylePlan({
      kind: "column",
      seriesCount: 1,
      sourceFormat: "0.000%",
    });

    expect(style).toMatchObject({
      legendPosition: "bottom",
      legendOverlay: false,
      chartAreaFill: "#FFFFFF",
      plotAreaFill: "#FFFFFF",
      showOuterBorder: false,
      textSizePoints: 8,
      showGridlines: false,
      valueAxisNumberFormat: undefined,
    });
  });

  it("plans right placement except where the estimated chart would cross Excel's final column", () => {
    const right = buildChartStylePlan({
      seriesCount: 1,
      kind: "column",
      sourceFormat: "General",
      selectionColumn: 16380,
      selectionColumnCount: 2,
      estimatedChartColumns: 2,
    });
    const below = buildChartStylePlan({
      seriesCount: 1,
      kind: "column",
      sourceFormat: "General",
      selectionColumn: 16380,
      selectionColumnCount: 2,
      estimatedChartColumns: 3,
    });

    expect(right.placement).toEqual({ side: "right", gutterPoints: 18 });
    expect(below.placement).toEqual({ side: "below", gutterPoints: 18 });
  });
});
