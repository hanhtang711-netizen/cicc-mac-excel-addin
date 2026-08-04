import { describe, expect, it } from "vitest";
import { CHART_CATALOG } from "../../src/charts/chartCatalog";
import { buildChartPlan } from "../../src/charts/chartPlanner";
import type { ParsedSelection } from "../../src/core/types";

const parsedSelection: ParsedSelection = {
  worksheetName: "Data",
  sourceAddress: "'Data'!$A$1:$C$4",
  title: "Weekly prices",
  headerRows: 1,
  categoryColumnOffset: 0,
  categoryAddress: "'Data'!$A$2:$A$4",
  orientation: "columns",
  series: [
    { name: "Pulp", valuesAddress: "'Data'!$B$2:$B$4", valueColumnOffset: 1 },
    { name: "Paper", valuesAddress: "'Data'!$C$2:$C$4", valueColumnOffset: 2 },
  ],
  numberFormats: [
    ["General", "General", "General"],
    ["yyyy-mm-dd", "0.0", "0.0"],
    ["yyyy-mm-dd", "0.0", "0.0"],
    ["yyyy-mm-dd", "0.0", "0.0"],
  ],
};

describe("buildChartPlan", () => {
  it("contains exactly the ten supported native chart actions", () => {
    expect(Object.keys(CHART_CATALOG)).toEqual([
      "column",
      "columnStacked",
      "line",
      "lineMarkers",
      "pie",
      "bar",
      "scatterTrend",
      "pieExploded",
      "columnStacked100",
      "lineStacked",
    ]);
  });

  it.each([
    ["column", "columnClustered"],
    ["columnStacked", "columnStacked"],
    ["line", "line"],
    ["lineMarkers", "lineMarkers"],
    ["pie", "pie"],
    ["bar", "barClustered"],
    ["scatterTrend", "xyscatter"],
    ["pieExploded", "pieExploded"],
    ["columnStacked100", "columnStacked100"],
    ["lineStacked", "lineStacked"],
  ] as const)("maps %s to %s", (kind, excelType) => {
    expect(CHART_CATALOG[kind].excelType).toBe(excelType);
    const selection = kind === "pie" || kind === "pieExploded"
      ? { ...parsedSelection, series: [parsedSelection.series[0]] }
      : parsedSelection;
    expect(buildChartPlan(selection, kind, {}).excelType).toBe(excelType);
  });

  it("rejects a pie selection with two value series", () => {
    expect(() => buildChartPlan(parsedSelection, "pie", {})).toThrow("pie_requires_one_series");
  });

  it("rejects a scatter selection whose X values are text", () => {
    const textCategories = {
      ...parsedSelection,
      numberFormats: parsedSelection.numberFormats.map((row, rowIndex) =>
        rowIndex === 0 ? row : ["@", ...row.slice(1)],
      ),
    };

    expect(() => buildChartPlan(textCategories, "scatterTrend", {})).toThrow(
      "scatter_requires_numeric_x",
    );
  });

  it("rejects a scatter selection without a numeric Y series", () => {
    const textValues = {
      ...parsedSelection,
      numberFormats: parsedSelection.numberFormats.map((row, rowIndex) =>
        rowIndex === 0 ? row : [row[0], "@", "@"],
      ),
    };

    expect(() => buildChartPlan(textValues, "scatterTrend", {})).toThrow(
      "scatter_requires_numeric_x",
    );
  });

  it("preserves series order and defaults a scatter trendline on", () => {
    const plan = buildChartPlan(parsedSelection, "scatterTrend", {});

    expect(plan.series).toEqual([
      { name: "Pulp", categoryAddress: "'Data'!$A$2:$A$4", valuesAddress: "'Data'!$B$2:$B$4" },
      { name: "Paper", categoryAddress: "'Data'!$A$2:$A$4", valuesAddress: "'Data'!$C$2:$C$4" },
    ]);
    expect(plan.addLinearTrendline).toBe(true);
  });

  it("allows the advanced pane to disable a scatter trendline", () => {
    expect(buildChartPlan(parsedSelection, "scatterTrend", { addTrendline: false }).addLinearTrendline).toBe(
      false,
    );
  });

  it("does not add a trendline to non-scatter charts", () => {
    expect(buildChartPlan(parsedSelection, "column", { addTrendline: true }).addLinearTrendline).toBe(false);
  });
});
