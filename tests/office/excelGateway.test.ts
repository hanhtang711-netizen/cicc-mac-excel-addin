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
  sizePreset: "medium",
  title: "Weekly prices",
  showDataLabels: true,
  addLinearTrendline: false,
};

const style: ChartStylePlan = {
  seriesColors: ["#640000", "#8A2626"],
  widthPoints: 326,
  heightPoints: 190,
  legendPosition: "bottom",
  chartAreaFill: "#FFFFFF",
  plotAreaFill: "#FFFFFF",
  showOuterBorder: false,
  textSizePoints: 8,
  majorGridlineColor: "#D9D9D9",
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

    await gateway.createChart(switchedPlan, style);

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
      width: 326,
      height: 190,
      left: 338,
      top: 40,
    });
    expect(fake.chart.title).toMatchObject({ visible: true, text: "Weekly prices" });
    expect(fake.chart.legend).toMatchObject({ visible: true, position: "Bottom" });
    expect(fake.series[0].format.fill.setSolidColor).toHaveBeenCalledWith("#640000");
    expect(fake.series[1].format.fill.setSolidColor).toHaveBeenCalledWith("#8A2626");
    expect(fake.chart.dataLabels.showValue).toBe(true);
    expect(fake.valueAxis.numberFormat).toBe("0.0%");
    expect(fake.categoryAxis.numberFormat).toBe("yyyy-mm-dd");
    expect(fake.valueAxis.majorGridlines.format.line.color).toBe("#D9D9D9");
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
    }, style);

    expect(fake.chartsAdd).toHaveBeenCalledWith("XYScatter", fake.sourceRange, "Columns");
    expect(fake.defaultSeries.delete).toHaveBeenCalledOnce();
    expect(fake.addedSeries).toHaveLength(2);
    expect(fake.addedSeries[0].name).toBe("Pulp");
    expect(fake.addedSeries[0].setXAxisValues).toHaveBeenCalledWith(fake.ranges.get("$A$2:$A$4"));
    expect(fake.addedSeries[0].setValues).toHaveBeenCalledWith(fake.ranges.get("$B$2:$B$4"));
    expect(fake.addedSeries[1].setValues).toHaveBeenCalledWith(fake.ranges.get("$C$2:$C$4"));
    expect(fake.addedSeries[0].trendlines.add).toHaveBeenCalledWith("Linear");
    expect(fake.addedSeries[1].trendlines.add).toHaveBeenCalledWith("Linear");
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
  const makeSeries = (seriesPoints: typeof points = []) => ({
    name: "",
    delete: vi.fn(),
    setXAxisValues: vi.fn(),
    setValues: vi.fn(),
    format: { fill: makeFill(), line: { color: "" } },
    trendlines: { add: vi.fn() },
    points: { items: seriesPoints, load: vi.fn() },
  });
  const series = [makeSeries(points), makeSeries()];
  const defaultSeries = series[0];
  const addedSeries: ReturnType<typeof makeSeries>[] = [];
  const categoryAxis = {
    numberFormat: "",
    format: { font: { size: 0 } },
    majorGridlines: { format: { line: { color: "" } } },
  };
  const valueAxis = {
    numberFormat: "",
    format: { font: { size: 0 } },
    majorGridlines: { format: { line: { color: "" } } },
  };
  const chart = {
    width: 0,
    height: 0,
    left: 0,
    top: 0,
    delete: vi.fn(),
    title: { visible: false, text: "", format: { font: { size: 0 } } },
    legend: { visible: false, position: "", format: { font: { size: 0 } } },
    dataLabels: { showValue: false, format: { font: { size: 0 } } },
    format: {
      fill: makeFill(),
      border: { clear: vi.fn() },
      font: { size: 0 },
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
