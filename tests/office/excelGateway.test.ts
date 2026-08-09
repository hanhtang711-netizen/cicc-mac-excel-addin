import { afterEach, describe, expect, it, vi } from "vitest";
import { buildChartPlan } from "../../src/charts/chartPlanner";
import { parseSelection } from "../../src/core/selectionParser";
import { ExcelGateway } from "../../src/office/excelGateway";
import { AddinError } from "../../src/core/errors";
import type { ChartPlan, ChartStylePlan } from "../../src/core/types";

const plan: ChartPlan = {
  kind: "column",
  excelType: "columnClustered",
  worksheetName: "Data",
  sourceAddress: "'Data'!$A$1:$C$4",
  orientation: "columns",
  series: [
    { name: "Pulp", categoryAddress: "'Data'!$A$2:$A$4", valuesAddress: "'Data'!$B$2:$B$4" },
    { name: "Paper", categoryAddress: "'Data'!$A$2:$A$4", valuesAddress: "'Data'!$C$2:$C$4" },
  ],
  addLinearTrendline: false,
};

const style: ChartStylePlan = {
  seriesColors: ["#640000", "#8A2626"],
  seriesStyle: "fill-no-border",
  widthPoints: 453.5433072,
  heightPoints: 255.1181103,
  legendPosition: "bottom",
  legendOverlay: false,
  legendFontSizePoints: 9,
  chartAreaFill: "#FFFFFF",
  plotAreaFill: "#FFFFFF",
  showOuterBorder: false,
  showTitle: false,
  showDataLabels: false,
  showGridlines: false,
  lineWidthPoints: 1.5,
  smoothLines: false,
  textSizePoints: 8,
  valueAxisNumberFormat: "0.0%",
  categoryAxisNumberFormat: "yyyy-mm-dd",
  warnings: [],
  placement: { side: "right", gutterPoints: 18 },
};

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("ExcelGateway.readSelection", () => {
  it("checks a single area and returns a detached snapshot", async () => {
    const values = [["Date", "Value"], [1, 10]];
    const texts = [["Date", "Value"], ["2026-01-01", "10"]];
    const numberFormat = [["General", "General"], ["yyyy-mm-dd", "0"]];
    const areas = { areaCount: 1, load: vi.fn() };
    const range = {
      address: "Data!A1:B2",
      rowIndex: 0,
      columnIndex: 0,
      rowCount: 2,
      columnCount: 2,
      values,
      text: texts,
      numberFormat,
      load: vi.fn(),
    };
    const sheet = { name: "Data", load: vi.fn() };
    const sync = vi.fn().mockResolvedValue(undefined);
    stubExcel({
      workbook: {
        getSelectedRanges: vi.fn(() => areas),
        getSelectedRange: vi.fn(() => range),
        worksheets: { getActiveWorksheet: vi.fn(() => sheet) },
      },
      sync,
    });

    const snapshot = await new ExcelGateway().readSelection();

    expect(areas.load).toHaveBeenCalledWith("areaCount");
    expect(range.load).toHaveBeenCalledWith(
      "address,rowIndex,columnIndex,rowCount,columnCount,values,text,numberFormat",
    );
    expect(sheet.load).toHaveBeenCalledWith("name");
    expect(sync).toHaveBeenCalledTimes(2);
    expect(snapshot).toEqual({
      worksheetName: "Data",
      address: "Data!A1:B2",
      rowIndex: 0,
      columnIndex: 0,
      rowCount: 2,
      columnCount: 2,
      values,
      texts,
      numberFormats: numberFormat,
    });
    expect(snapshot.values).not.toBe(values);
    expect(snapshot.values[0]).not.toBe(values[0]);
  });

  it("rejects a multi-area selection before reading a range", async () => {
    const getSelectedRange = vi.fn();
    stubExcel({
      workbook: {
        getSelectedRanges: vi.fn(() => ({ areaCount: 2, load: vi.fn() })),
        getSelectedRange,
        worksheets: { getActiveWorksheet: vi.fn() },
      },
      sync: vi.fn().mockResolvedValue(undefined),
    });

    const error = await new ExcelGateway().readSelection().catch((caught: unknown) => caught);

    expect(error).toBeInstanceOf(AddinError);
    expect(error).toMatchObject({ code: "invalid_selection" });
    expect(getSelectedRange).not.toHaveBeenCalled();
  });
});

describe("ExcelGateway.createChart", () => {
  it("creates on the worksheet read earlier when the active sheet later changes", async () => {
    const data = makeChartHarness("Data");
    const other = makeChartHarness("Other");
    let activeWorksheet = data.worksheet;
    const getItem = vi.fn(() => data.worksheet);
    const selectionRange = {
      address: "Data!A1:C4",
      rowIndex: 0,
      columnIndex: 0,
      rowCount: 4,
      columnCount: 3,
      values: [["X", "Pulp", "Paper"], [1, 10, 20], [2, 11, 21], [3, 12, 22]],
      text: [["X", "Pulp", "Paper"], ["1", "10", "20"], ["2", "11", "21"], ["3", "12", "22"]],
      numberFormat: [
        ["General", "General", "General"],
        ["0", "0", "0"],
        ["0", "0", "0"],
        ["0", "0", "0"],
      ],
      load: vi.fn(),
    };
    stubExcel({
      workbook: {
        getSelectedRanges: vi.fn(() => ({ areaCount: 1, load: vi.fn() })),
        getSelectedRange: vi.fn(() => selectionRange),
        worksheets: {
          getActiveWorksheet: vi.fn(() => activeWorksheet),
          getItem,
        },
      },
      sync: data.sync,
    });
    const gateway = new ExcelGateway();
    const snapshot = await gateway.readSelection();
    activeWorksheet = other.worksheet;
    const switchedPlan = buildChartPlan(parseSelection(snapshot), "scatterTrend");

    await gateway.createChart(switchedPlan, { ...style, seriesStyle: "scatter" });

    expect(snapshot.worksheetName).toBe("Data");
    expect(getItem).toHaveBeenCalledWith("Data");
    expect(data.chartsAdd).toHaveBeenCalledWith("XYScatter", data.sourceRange, "Columns");
    expect(other.chartsAdd).not.toHaveBeenCalled();
    expect(data.worksheet.getRange).toHaveBeenCalledWith("A1:C4");
    expect(data.worksheet.getRange).toHaveBeenCalledWith("$A$2:$A$4");
    expect(data.worksheet.getRange).toHaveBeenCalledWith("$B$2:$B$4");
    expect(data.worksheet.getRange).toHaveBeenCalledWith("$C$2:$C$4");
  });

  it("creates and styles a native chart using live selection geometry", async () => {
    const fake = makeChartHarness();
    stubExcel(fake.context);

    await new ExcelGateway().createChart(plan, style);

    expect(fake.chartsAdd).toHaveBeenCalledWith("ColumnClustered", fake.sourceRange, "Columns");
    expect(fake.chart).toMatchObject({
      width: 453.5433072,
      height: 255.1181103,
      left: 338,
      top: 40,
    });
    expect(fake.chart.title).toMatchObject({ visible: false });
    expect(fake.chart.legend).toMatchObject({ visible: true, position: "Bottom" });
    expect(fake.series[0].format.fill.setSolidColor).toHaveBeenCalledWith("#640000");
    expect(fake.series[1].format.fill.setSolidColor).toHaveBeenCalledWith("#8A2626");
    expect(fake.chart.dataLabels.showValue).toBe(false);
    expect(fake.valueAxis.numberFormat).toBe("0.0%");
    expect(fake.categoryAxis.numberFormat).toBe("yyyy-mm-dd");
    expect(fake.valueAxis.majorGridlines.visible).toBe(false);
  });

  it("applies the strict common legend, axis, and gridline contract", async () => {
    const fake = makeChartHarness();
    stubExcel(fake.context);

    await new ExcelGateway().createChart(plan, style);

    expect(fake.chart.title.visible).toBe(false);
    expect(fake.chart.dataLabels.showValue).toBe(false);
    expect(fake.chart.legend).toMatchObject({
      visible: true,
      position: "Bottom",
      overlay: false,
    });
    expect(fake.chart.legend.format.font).toMatchObject({
      name: "Arial",
      size: 9,
      color: "#000000",
    });
    expect(fake.categoryAxis).toMatchObject({
      visible: true,
      position: "Minimum",
      majorTickMark: "Outside",
      tickLabelPosition: "NextToAxis",
    });
    expect(fake.categoryAxis.title.visible).toBe(false);
    expect(fake.valueAxis.title.visible).toBe(false);
    expect(fake.valueAxis.majorGridlines.visible).toBe(false);
    expect(fake.valueAxis.minorGridlines.visible).toBe(false);
    expect(fake.categoryAxis.format.font).toMatchObject({ name: "Arial", size: 8, color: "#000000" });
    expect(fake.valueAxis.format.font).toMatchObject({ name: "Arial", size: 8, color: "#000000" });
  });

  it("styles line series without touching the unsupported Mac series fill", async () => {
    const fake = makeChartHarness();
    stubExcel(fake.context);

    await new ExcelGateway().createChart(
      { ...plan, kind: "line", excelType: "line" },
      { ...style, seriesStyle: "line", smoothLines: true },
    );

    expect(fake.series[0].format.fill.setSolidColor).not.toHaveBeenCalled();
    expect(fake.series[0].format.line.color).toBe("#640000");
    expect(fake.series[0].format.line.weight).toBe(1.5);
    expect(fake.series[0].smooth).toBe(true);
    expect(fake.series[0].markerStyle).toBe("None");
  });

  it("fills column series and removes their outlines", async () => {
    const fake = makeChartHarness();
    stubExcel(fake.context);

    await new ExcelGateway().createChart(plan, { ...style, seriesStyle: "fill-no-border" });

    expect(fake.series[0].format.fill.setSolidColor).toHaveBeenCalledWith("#640000");
    expect(fake.series[0].format.line.clear).toHaveBeenCalledOnce();
  });

  it("places a chart below the source range when requested", async () => {
    const fake = makeChartHarness();
    stubExcel(fake.context);

    await new ExcelGateway().createChart(plan, {
      ...style,
      placement: { side: "below", gutterPoints: 18 },
    });

    expect(fake.chart.left).toBe(20);
    expect(fake.chart.top).toBe(118);
  });

  it("builds scatter series from explicit X and Y ranges and adds linear trendlines", async () => {
    const fake = makeChartHarness();
    stubExcel(fake.context);

    await new ExcelGateway().createChart({
      ...plan,
      kind: "scatterTrend",
      excelType: "xyscatter",
      addLinearTrendline: true,
    }, { ...style, seriesStyle: "scatter" });

    expect(fake.chartsAdd).toHaveBeenCalledWith("XYScatter", fake.sourceRange, "Columns");
    expect(fake.defaultSeries.delete).toHaveBeenCalledOnce();
    expect(fake.addedSeries).toHaveLength(2);
    expect(fake.addedSeries[0].name).toBe("Pulp");
    expect(fake.addedSeries[0].setXAxisValues).toHaveBeenCalledWith(fake.ranges.get("$A$2:$A$4"));
    expect(fake.addedSeries[0].setValues).toHaveBeenCalledWith(fake.ranges.get("$B$2:$B$4"));
    expect(fake.addedSeries[1].setValues).toHaveBeenCalledWith(fake.ranges.get("$C$2:$C$4"));
    expect(fake.addedSeries[0].trendlines.add).toHaveBeenCalledWith("Linear");
    expect(fake.addedSeries[1].trendlines.add).toHaveBeenCalledWith("Linear");
    expect(fake.addedSeries[0].format.fill.setSolidColor).not.toHaveBeenCalled();
    expect(fake.addedSeries[0].format.line.color).toBe("#640000");
    expect(fake.addedSeries[0].trendline.format.line.color).toBe("#640000");
    expect(fake.addedSeries[0].trendline.format.line.weight).toBe(1.5);
  });

  it("colors pie points individually in source order", async () => {
    const fake = makeChartHarness();
    stubExcel(fake.context);

    await new ExcelGateway().createChart({
      ...plan,
      kind: "pie",
      excelType: "pie",
      series: [plan.series[0]],
    }, { ...style, seriesColors: ["#640000"] });

    expect(fake.points[0].format.fill.setSolidColor).toHaveBeenCalledWith("#640000");
    expect(fake.points[1].format.fill.setSolidColor).toHaveBeenCalledWith("#8A2626");
  });

  it("does not access axes that a pie chart does not expose", async () => {
    const fake = makeChartHarness();
    Object.defineProperty(fake.chart, "axes", {
      get() {
        throw new Error("pie charts have no axes");
      },
    });
    stubExcel(fake.context);

    await expect(new ExcelGateway().createChart({
      ...plan,
      kind: "pie",
      excelType: "pie",
      series: [plan.series[0]],
    }, style)).resolves.toBeUndefined();
  });

  it("deletes a newly-created chart before rethrowing a typed runtime error", async () => {
    const fake = makeChartHarness();
    fake.chart.format.fill.setSolidColor.mockImplementation(() => {
      throw new Error("styling failed");
    });
    stubExcel(fake.context);

    const error = await new ExcelGateway().createChart(plan, style).catch((caught: unknown) => caught);

    expect(fake.chart.delete).toHaveBeenCalledOnce();
    expect(fake.sync.mock.invocationCallOrder.at(-1)).toBeGreaterThan(
      fake.chart.delete.mock.invocationCallOrder[0],
    );
    expect(error).toBeInstanceOf(AddinError);
    expect(error).toMatchObject({ code: "excel_runtime_error" });
  });

  it("preserves the failing series-style stage while rolling back the chart", async () => {
    const fake = makeChartHarness();
    const cause = new Error("series fill failed");
    fake.series[0].format.fill.setSolidColor.mockImplementation(() => {
      throw cause;
    });
    stubExcel(fake.context);

    const error = await new ExcelGateway().createChart(plan, style).catch((caught: unknown) => caught);

    expect(fake.chart.delete).toHaveBeenCalledOnce();
    expect(error).toBeInstanceOf(AddinError);
    expect(error).toMatchObject({
      code: "excel_runtime_error",
      details: { stage: "series-style", cause },
    });
  });
});

describe("ExcelGateway.applyTablePlan", () => {
  it("applies the strict standard table format and hides worksheet gridlines", async () => {
    const makeFont = () => ({ color: "", bold: false, size: 0 });
    const makeFill = () => ({ color: "" });
    const borders = new Map(
      ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"]
        .map((index) => [index, { style: "Continuous" }]),
    );
    const header = { format: { fill: makeFill(), font: makeFont() } };
    const columnFormats = Array.from({ length: 3 }, () => ({
      columnWidth: 200,
      autofitColumns: vi.fn(),
      load: vi.fn(),
    }));
    const selectedRange = {
      format: {
        fill: makeFill(),
        font: makeFont(),
        horizontalAlignment: "",
        verticalAlignment: "",
        wrapText: false,
        rowHeight: 0,
        borders: { getItem: vi.fn((index: string) => borders.get(index)) },
      },
      getRow: vi.fn(() => header),
      getColumn: vi.fn((index: number) => ({ format: columnFormats[index] })),
    };
    const worksheet = { showGridlines: true, getRange: vi.fn(() => selectedRange) };
    const sync = vi.fn().mockResolvedValue(undefined);
    stubExcel({ workbook: { worksheets: { getItem: vi.fn(() => worksheet) } }, sync });

    await new ExcelGateway().applyTablePlan({
      kind: "standard",
      worksheetName: "Data",
      address: "'Data'!$A$1:$C$4",
      rowCount: 4,
      columnCount: 3,
      header: { fill: "#8A2626", fontColor: "#FFFFFF", bold: true, fontSize: 8 },
      body: { fill: "#FFFFFF", fontColor: "#000000", bold: false, fontSize: 8 },
      horizontalAlignment: "left",
      verticalAlignment: "center",
      wrapText: true,
      rowHeight: 16,
      clearBorders: true,
      hideWorksheetGridlines: true,
      rowFills: [],
      preserve: ["values", "formulas", "numberFormats", "merges", "conditionalFormats"],
    });

    expect(selectedRange.format).toMatchObject({
      fill: { color: "#FFFFFF" },
      font: { color: "#000000", bold: false, size: 8 },
      horizontalAlignment: "Left",
      verticalAlignment: "Center",
      wrapText: true,
      rowHeight: 16,
    });
    expect(header.format).toMatchObject({
      fill: { color: "#8A2626" },
      font: { color: "#FFFFFF", bold: true, size: 8 },
    });
    expect([...borders.values()].every((border) => border.style === "None")).toBe(true);
    expect(worksheet.showGridlines).toBe(false);
    expect(columnFormats.every((format) => format.autofitColumns.mock.calls.length === 1)).toBe(true);
    expect(columnFormats.every((format) => format.columnWidth === 180)).toBe(true);
  });

  it("applies zebra fills only to rows on the worksheet captured by the plan", async () => {
    const fillWrites: Array<{ row: number; color: string }> = [];
    const unexpectedWrite = vi.fn(() => {
      throw new Error("zebra must not change this format");
    });
    const makeRow = (row: number) => ({
      format: {
        fill: {
          set color(color: string) {
            fillWrites.push({ row, color });
          },
        },
        font: { set color(_value: string) { unexpectedWrite(); }, set bold(_value: boolean) { unexpectedWrite(); }, set size(_value: number) { unexpectedWrite(); } },
        borders: { getItem: unexpectedWrite },
        set horizontalAlignment(_value: string) { unexpectedWrite(); },
      },
    });
    const selectedRange = {
      getRow: vi.fn(makeRow),
      format: {
        fill: { set color(_value: string) { unexpectedWrite(); } },
        font: { set color(_value: string) { unexpectedWrite(); }, set bold(_value: boolean) { unexpectedWrite(); }, set size(_value: number) { unexpectedWrite(); } },
        borders: { getItem: unexpectedWrite },
        autofitColumns: unexpectedWrite,
        autofitRows: unexpectedWrite,
        set horizontalAlignment(_value: string) { unexpectedWrite(); },
      },
    };
    const dataWorksheet = { getRange: vi.fn(() => selectedRange) };
    const activeWorksheet = { getRange: vi.fn(() => { throw new Error("active sheet must not be used"); }) };
    const sync = vi.fn().mockResolvedValue(undefined);
    const getItem = vi.fn(() => dataWorksheet);
    stubExcel({
      workbook: {
        worksheets: { getItem, getActiveWorksheet: vi.fn(() => activeWorksheet) },
      },
      sync,
    });

    await new ExcelGateway().applyTablePlan({
      kind: "zebra",
      worksheetName: "Data",
      address: "'Data'!$A$1:$C$3",
      rowCount: 3,
      columnCount: 3,
      rowFills: [{ rowOffset: 1, fill: "#FFFFFF" }, { rowOffset: 2, fill: "#F5F5F5" }],
      preserve: ["values", "formulas", "numberFormats", "merges", "conditionalFormats", "fonts", "borders", "alignment"],
    });

    expect(getItem).toHaveBeenCalledWith("Data");
    expect(dataWorksheet.getRange).toHaveBeenCalledWith("$A$1:$C$3");
    expect(fillWrites).toEqual([{ row: 1, color: "#FFFFFF" }, { row: 2, color: "#F5F5F5" }]);
    expect(unexpectedWrite).not.toHaveBeenCalled();
    expect(sync).toHaveBeenCalledOnce();
  });
});

function stubExcel(context: object): void {
  vi.stubGlobal("Excel", {
    run: vi.fn(async (callback: (value: object) => Promise<unknown>) => callback(context)),
    ChartType: {
      columnClustered: "ColumnClustered",
      xyscatter: "XYScatter",
      pie: "Pie",
    },
    ChartSeriesBy: { columns: "Columns", rows: "Rows" },
    ChartLegendPosition: { top: "Top", bottom: "Bottom", left: "Left", right: "Right" },
    ChartTrendlineType: { linear: "Linear" },
  });
}

function makeChartHarness(name = "Data") {
  const makeFill = () => ({ setSolidColor: vi.fn() });
  const points = [
    { format: { fill: makeFill() } },
    { format: { fill: makeFill() } },
    { format: { fill: makeFill() } },
  ];
  const makeSeries = (seriesPoints: typeof points = []) => {
    const trendline = { format: { line: { color: "", weight: 0 } } };
    return {
    name: "",
    delete: vi.fn(),
    setXAxisValues: vi.fn(),
    setValues: vi.fn(),
    smooth: false,
    markerStyle: "Automatic",
    markerBackgroundColor: "",
    markerForegroundColor: "",
    format: { fill: makeFill(), line: { color: "", weight: 0, clear: vi.fn() } },
    trendline,
    trendlines: { add: vi.fn(() => trendline) },
    points: { items: seriesPoints, load: vi.fn() },
    };
  };
  const series = [makeSeries(points), makeSeries()];
  const defaultSeries = series[0];
  const addedSeries: ReturnType<typeof makeSeries>[] = [];
  const categoryAxis = {
    visible: false,
    position: "",
    majorTickMark: "",
    tickLabelPosition: "",
    numberFormat: "",
    title: { visible: true },
    format: { font: { name: "", size: 0, color: "" } },
    majorGridlines: { visible: true, format: { line: { color: "" } } },
    minorGridlines: { visible: true, format: { line: { color: "" } } },
  };
  const valueAxis = {
    visible: false,
    position: "",
    majorTickMark: "",
    tickLabelPosition: "",
    numberFormat: "",
    title: { visible: true },
    format: { font: { name: "", size: 0, color: "" } },
    majorGridlines: { visible: true, format: { line: { color: "" } } },
    minorGridlines: { visible: true, format: { line: { color: "" } } },
  };
  const chart = {
    width: 0,
    height: 0,
    left: 0,
    top: 0,
    delete: vi.fn(),
    title: { visible: false, text: "", format: { font: { name: "", size: 0, color: "" } } },
    legend: { visible: false, position: "", overlay: true, format: { font: { name: "", size: 0, color: "" } } },
    dataLabels: { showValue: false, format: { font: { name: "", size: 0, color: "" } } },
    format: {
      fill: makeFill(),
      border: { clear: vi.fn() },
      font: { name: "", size: 0, color: "" },
    },
    plotArea: { format: { fill: makeFill() } },
    axes: { categoryAxis, valueAxis },
    series: {
      items: series,
      load: vi.fn(),
      add: vi.fn((name: string) => {
        const item = makeSeries();
        item.name = name;
        addedSeries.push(item);
        return item;
      }),
    },
  };
  const sourceRange = { left: 20, top: 40, width: 300, height: 60, load: vi.fn() };
  const ranges = new Map<string, object>([
    [plan.sourceAddress, sourceRange],
    ["$A$1:$C$4", sourceRange],
    ["A1:C4", sourceRange],
  ]);
  const getRange = vi.fn((address: string) => {
    if (!ranges.has(address)) {
      ranges.set(address, { address, load: vi.fn() });
    }
    return ranges.get(address);
  });
  const chartsAdd = vi.fn(() => chart);
  const sync = vi.fn().mockResolvedValue(undefined);
  const worksheet = { name, load: vi.fn(), getRange, charts: { add: chartsAdd } };
  return {
    context: {
      workbook: {
        worksheets: {
          getActiveWorksheet: vi.fn(() => worksheet),
          getItem: vi.fn(() => worksheet),
        },
      },
      sync,
    },
    chart,
    series,
    defaultSeries,
    addedSeries,
    points,
    sourceRange,
    ranges,
    chartsAdd,
    sync,
    categoryAxis,
    valueAxis,
    worksheet,
  };
}
