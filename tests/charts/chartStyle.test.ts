import { describe, expect, it } from "vitest";
import {
  buildChartStylePlan,
  CICC_SERIES_COLORS,
  cmToPoints,
} from "../../src/charts/chartStyle";

describe("CICC chart style", () => {
  it("cycles the six exact series colors", () => {
    const style = buildChartStylePlan({ seriesCount: 8, options: {}, sourceFormat: "0.0%" });

    expect(style.seriesColors).toEqual([
      ...CICC_SERIES_COLORS,
      "#640000",
      "#8A2626",
    ]);
    expect(style.warnings).toContain("series_palette_reused");
  });

  it("uses the medium preset and percentage precision", () => {
    const style = buildChartStylePlan({ seriesCount: 2, options: {}, sourceFormat: "0.0%" });

    expect(style.widthPoints).toBeCloseTo(cmToPoints(11.5), 4);
    expect(style.heightPoints).toBeCloseTo(cmToPoints(6.7), 4);
    expect(style.valueAxisNumberFormat).toBe("0.0%");
    expect(style.legendPosition).toBe("bottom");
  });

  it("uses the requested custom dimensions only when both are finite and positive", () => {
    const style = buildChartStylePlan({
      seriesCount: 1,
      options: { sizePreset: "custom", widthCm: 14, heightCm: 8 },
      sourceFormat: "General",
    });

    expect(style.widthPoints).toBeCloseTo(396.8503938, 4);
    expect(style.heightPoints).toBeCloseTo(226.7716536, 4);
    expect(() => buildChartStylePlan({
      seriesCount: 1,
      options: { sizePreset: "custom", widthCm: 0, heightCm: 8 },
      sourceFormat: "General",
    })).toThrow("invalid_chart_dimensions");
  });

  it("preserves supported percentage formats and normalizes date-like categories", () => {
    const style = buildChartStylePlan({
      seriesCount: 1,
      options: {},
      sourceFormat: "0.00%",
      categoryFormat: "m/d/yyyy",
    });

    expect(style.valueAxisNumberFormat).toBe("0.00%");
    expect(style.categoryAxisNumberFormat).toBe("yyyy-mm-dd");
  });

  it("does not treat quoted currency text as a date-like category format", () => {
    const style = buildChartStylePlan({
      seriesCount: 1,
      options: {},
      sourceFormat: "General",
      categoryFormat: "\"RMB\" #,##0",
    });

    expect(style.categoryAxisNumberFormat).toBeUndefined();
  });

  it("applies the CICC base appearance and preserves a selected legend position", () => {
    const style = buildChartStylePlan({
      seriesCount: 1,
      options: { legendPosition: "right" },
      sourceFormat: "0.000%",
    });

    expect(style).toMatchObject({
      legendPosition: "right",
      chartAreaFill: "#FFFFFF",
      plotAreaFill: "#FFFFFF",
      showOuterBorder: false,
      textSizePoints: 8,
      majorGridlineColor: "#D9D9D9",
      valueAxisNumberFormat: undefined,
    });
  });

  it("plans right placement except where the estimated chart would cross Excel's final column", () => {
    const right = buildChartStylePlan({
      seriesCount: 1,
      options: {},
      sourceFormat: "General",
      selectionColumn: 16380,
      selectionColumnCount: 2,
      estimatedChartColumns: 2,
    });
    const below = buildChartStylePlan({
      seriesCount: 1,
      options: {},
      sourceFormat: "General",
      selectionColumn: 16380,
      selectionColumnCount: 2,
      estimatedChartColumns: 3,
    });

    expect(right.placement).toEqual({ side: "right", gutterPoints: 18 });
    expect(below.placement).toEqual({ side: "below", gutterPoints: 18 });
  });
});
