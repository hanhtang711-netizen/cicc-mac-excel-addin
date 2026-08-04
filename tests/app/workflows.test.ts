import { describe, expect, it, vi } from "vitest";
import { ChartService } from "../../src/app/chartService";
import { TableService } from "../../src/app/tableService";
import type { ChartPlan, ChartStylePlan, SelectionSnapshot, TableFormatPlan } from "../../src/core/types";
import type { AddinCapabilities } from "../../src/office/capability";
import {
  blankValuesFixture,
  dateMultiSeriesFixture,
  dateSingleSeriesFixture,
  eightSeriesFixture,
  horizontalOrientationFixture,
  invalidScatterXFixture,
  macLineRegressionFixture,
  negativeValuesFixture,
  percentageFixture,
  textCategoriesFixture,
  titleRowFixture,
} from "../fixtures/selectionCases";

const supportedCapabilities: AddinCapabilities = {
  base: true,
  explodedPie: true,
  trendlines: true,
};

const createFixtureGateway = (snapshot: SelectionSnapshot) => ({
  readSelection: vi.fn().mockResolvedValue(snapshot),
  createChart: vi.fn<(plan: ChartPlan, style: ChartStylePlan) => Promise<void>>().mockResolvedValue(undefined),
  applyTablePlan: vi.fn<(plan: TableFormatPlan) => Promise<void>>().mockResolvedValue(undefined),
});

describe("chart workflows", () => {
  it("builds the exact Mac line regression selection without explicit options", async () => {
    const gateway = createFixtureGateway(macLineRegressionFixture);

    await new ChartService(gateway, supportedCapabilities).create("line");

    expect(gateway.createChart.mock.calls[0][0]).toMatchObject({
      kind: "line",
      orientation: "columns",
      series: [
        { name: "A", categoryAddress: "'Fixture'!$A$2:$A$13", valuesAddress: "'Fixture'!$B$2:$B$13" },
        { name: "B", categoryAddress: "'Fixture'!$A$2:$A$13", valuesAddress: "'Fixture'!$C$2:$C$13" },
      ],
    });
  });

  it.each([
    ["column", "columnClustered"],
    ["columnStacked", "columnStacked"],
    ["line", "line"],
    ["lineMarkers", "lineMarkers"],
    ["bar", "barClustered"],
    ["scatterTrend", "xyscatter"],
    ["columnStacked100", "columnStacked100"],
    ["lineStacked", "lineStacked"],
  ] as const)("builds the complete %s plan from a multi-series selection", async (kind, excelType) => {
    const gateway = createFixtureGateway(dateMultiSeriesFixture);

    await new ChartService(gateway, supportedCapabilities).create(kind);

    const [plan, style] = gateway.createChart.mock.calls[0];
    expect(plan).toMatchObject({
      kind,
      excelType,
      orientation: "columns",
      sourceAddress: "Fixture!A1:C4",
      series: [
        { name: "Alpha", categoryAddress: "'Fixture'!$A$2:$A$4", valuesAddress: "'Fixture'!$B$2:$B$4" },
        { name: "Beta", categoryAddress: "'Fixture'!$A$2:$A$4", valuesAddress: "'Fixture'!$C$2:$C$4" },
      ],
    });
    expect(style).toMatchObject({
      seriesColors: ["#640000", "#8A2626"],
      categoryAxisNumberFormat: "yyyy-mm-dd",
      warnings: [],
    });
  });

  it.each([
    ["pie", "pie"],
    ["pieExploded", "pieExploded"],
  ] as const)("builds a one-series %s plan", async (kind, excelType) => {
    const gateway = createFixtureGateway(dateSingleSeriesFixture);

    await new ChartService(gateway, supportedCapabilities).create(kind);

    const [plan, style] = gateway.createChart.mock.calls[0];
    expect(plan).toMatchObject({ kind, excelType, series: [{ name: "Price" }] });
    expect(style.seriesColors).toEqual(["#640000"]);
  });

  it("ignores legacy visual overrides and applies the fixed Skill style", async () => {
    const gateway = createFixtureGateway(titleRowFixture);

    await new ChartService(gateway, supportedCapabilities).create("lineMarkers", {
      title: "Override",
      sizePreset: "custom",
      widthCm: 14,
      heightCm: 8,
      legendPosition: "top",
      showDataLabels: true,
    });

    const [plan, style] = gateway.createChart.mock.calls[0];
    expect(plan).toMatchObject({
      series: [{ valuesAddress: "'Fixture'!$B$3:$B$5" }, { valuesAddress: "'Fixture'!$C$3:$C$5" }],
    });
    expect(plan).not.toHaveProperty("title");
    expect(plan).not.toHaveProperty("showDataLabels");
    expect(plan).not.toHaveProperty("sizePreset");
    expect(style).toMatchObject({
      legendPosition: "bottom",
      showTitle: false,
      showDataLabels: false,
      categoryAxisNumberFormat: "yyyy-mm-dd",
    });
    expect(style.widthPoints).toBeCloseTo(453.5433072);
    expect(style.heightPoints).toBeCloseTo(255.1181103);
  });

  it("does not turn a standalone data title into a chart title", async () => {
    const gateway = createFixtureGateway(titleRowFixture);

    await new ChartService(gateway, supportedCapabilities).create("column");

    expect(gateway.createChart.mock.calls[0][0]).not.toHaveProperty("title");
  });

  it("maps row-oriented categories and series without transposing the fixture", async () => {
    const gateway = createFixtureGateway(horizontalOrientationFixture);

    await new ChartService(gateway, supportedCapabilities).create("line", { orientation: "rows" });

    expect(gateway.createChart.mock.calls[0][0]).toMatchObject({
      orientation: "rows",
      series: [
        { name: "Alpha", categoryAddress: "'Fixture'!$B$1:$D$1", valuesAddress: "'Fixture'!$B$2:$D$2" },
        { name: "Beta", categoryAddress: "'Fixture'!$B$1:$D$1", valuesAddress: "'Fixture'!$B$3:$D$3" },
      ],
    });
  });

  it("preserves text category addresses in a standard chart plan", async () => {
    const gateway = createFixtureGateway(textCategoriesFixture);

    await new ChartService(gateway, supportedCapabilities).create("column");

    expect(gateway.createChart.mock.calls[0][0].series).toEqual([
      { name: "Plan", categoryAddress: "'Fixture'!$A$2:$A$4", valuesAddress: "'Fixture'!$B$2:$B$4" },
      { name: "Actual", categoryAddress: "'Fixture'!$A$2:$A$4", valuesAddress: "'Fixture'!$C$2:$C$4" },
    ]);
  });

  it("uses a percentage value-axis format", async () => {
    const gateway = createFixtureGateway(percentageFixture);

    await new ChartService(gateway, supportedCapabilities).create("columnStacked100");

    expect(gateway.createChart.mock.calls[0][1].valueAxisNumberFormat).toBe("0.0%");
  });

  it.each([
    ["negative", negativeValuesFixture],
    ["blank", blankValuesFixture],
  ] as const)("accepts %s numeric values without rewriting the source range", async (_label, snapshot) => {
    const gateway = createFixtureGateway(snapshot);

    await new ChartService(gateway, supportedCapabilities).create("line");

    expect(gateway.createChart.mock.calls[0][0].sourceAddress).toBe(snapshot.address);
    expect(gateway.createChart.mock.calls[0][0].series).toHaveLength(snapshot.columnCount - 1);
  });

  it("returns a palette-reuse warning and repeats the six colors deterministically", async () => {
    const gateway = createFixtureGateway(eightSeriesFixture);

    const result = await new ChartService(gateway, supportedCapabilities).create("line");

    expect(result.warnings).toEqual(["series_palette_reused"]);
    expect(gateway.createChart.mock.calls[0][1].seriesColors).toEqual([
      "#640000", "#8A2626", "#3D889A", "#646C86", "#BE995D", "#DD965D", "#640000", "#8A2626",
    ]);
  });

  it("rejects invalid scatter X values before any workbook write", async () => {
    const gateway = createFixtureGateway(invalidScatterXFixture);

    await expect(new ChartService(gateway, supportedCapabilities).create("scatterTrend"))
      .rejects.toMatchObject({ code: "scatter_requires_numeric_x" });
    expect(gateway.createChart).not.toHaveBeenCalled();
  });
});

describe("table workflows", () => {
  it("keeps standard formatting and zebra fills as independent plans", async () => {
    const gateway = createFixtureGateway(dateMultiSeriesFixture);
    const service = new TableService(gateway, supportedCapabilities);

    await service.formatStandard();
    await service.applyZebra();

    const standard = gateway.applyTablePlan.mock.calls[0][0];
    const zebra = gateway.applyTablePlan.mock.calls[1][0];
    expect(standard).toMatchObject({
      kind: "standard",
      header: { fill: "#8A2626", fontColor: "#FFFFFF" },
      body: { fill: "#FFFFFF" },
      rowFills: [],
      preserve: ["values", "formulas", "numberFormats", "merges", "conditionalFormats"],
    });
    expect(zebra).toEqual({
      kind: "zebra",
      worksheetName: "Fixture",
      address: "Fixture!A1:C4",
      rowCount: 4,
      columnCount: 3,
      rowFills: [
        { rowOffset: 1, fill: "#FFFFFF" },
        { rowOffset: 2, fill: "#F5F5F5" },
        { rowOffset: 3, fill: "#FFFFFF" },
      ],
      preserve: ["values", "formulas", "numberFormats", "merges", "conditionalFormats", "fonts", "borders", "alignment"],
    });
  });
});
