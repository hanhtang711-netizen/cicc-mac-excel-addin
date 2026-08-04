import { AddinError } from "../core/errors";
import type {
  ChartPlan,
  ChartStylePlan,
  StandardTableFormatPlan,
  SelectionSnapshot,
  TableFormatPlan,
  TableHorizontalAlignment,
} from "../core/types";
import { CICC_SERIES_COLORS } from "../charts/chartStyle";

export class ExcelGateway {
  async readSelection(): Promise<SelectionSnapshot> {
    return Excel.run(async (context) => {
      const areas = context.workbook.getSelectedRanges();
      areas.load("areaCount");
      await context.sync();

      if (areas.areaCount !== 1) {
        throw new AddinError("invalid_selection");
      }

      const range = context.workbook.getSelectedRange();
      range.load("address,rowIndex,columnIndex,rowCount,columnCount,values,text,numberFormat");
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();

      return snapshotFromRange(sheet.name, range);
    });
  }

  async createChart(plan: ChartPlan, style: ChartStylePlan): Promise<void> {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(plan.worksheetName);
      const sourceRange = worksheet.getRange(localAddress(plan.sourceAddress));
      sourceRange.load("left,top,width,height");
      await context.sync();

      let chart: Excel.Chart | undefined;
      try {
        chart = worksheet.charts.add(
          resolveChartType(plan.excelType),
          sourceRange,
          plan.orientation === "rows" ? Excel.ChartSeriesBy.rows : Excel.ChartSeriesBy.columns,
        );

        chart.series.load("items");
        await context.sync();

        if (plan.kind === "scatterTrend") {
          for (const existingSeries of chart.series.items) {
            existingSeries.delete();
          }
          addScatterSeries(worksheet, chart, plan, style);
        } else {
          applySeriesColors(chart.series.items, style.seriesColors);
        }

        if (plan.kind === "pie" || plan.kind === "pieExploded") {
          await colorPiePoints(context, chart);
        }

        applyChartStyle(chart, plan, style, sourceRange);
        await context.sync();
      } catch (cause) {
        if (chart !== undefined) {
          try {
            chart.delete();
            await context.sync();
          } catch (rollbackError) {
            throw new AddinError("excel_runtime_error", { cause, rollbackError });
          }
        }
        throw new AddinError("excel_runtime_error", cause);
      }
    });
  }

  async applyTablePlan(plan: TableFormatPlan): Promise<void> {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(plan.worksheetName);
      const range = worksheet.getRange(localAddress(plan.address));

      if (plan.kind === "zebra") {
        applyZebraFills(range, plan);
        await context.sync();
        return;
      }

      applyStandardFormats(range, plan);
      await autofitWithinBounds(context, range, plan);
    });
  }
}

function snapshotFromRange(sheetName: string, range: Excel.Range): SelectionSnapshot {
  return {
    worksheetName: sheetName,
    address: range.address,
    rowIndex: range.rowIndex,
    columnIndex: range.columnIndex,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
    values: range.values.map((row) => row.slice()),
    texts: range.text.map((row) => row.slice()),
    numberFormats: range.numberFormat.map((row) => row.slice()),
  };
}

function resolveChartType(excelType: string): Excel.ChartType {
  switch (excelType) {
    case "columnClustered": return Excel.ChartType.columnClustered;
    case "columnStacked": return Excel.ChartType.columnStacked;
    case "line": return Excel.ChartType.line;
    case "lineMarkers": return Excel.ChartType.lineMarkers;
    case "pie": return Excel.ChartType.pie;
    case "barClustered": return Excel.ChartType.barClustered;
    case "xyscatter": return Excel.ChartType.xyscatter;
    case "pieExploded": return Excel.ChartType.pieExploded;
    case "columnStacked100": return Excel.ChartType.columnStacked100;
    case "lineStacked": return Excel.ChartType.lineStacked;
    default: throw new Error(`unsupported_chart_type:${excelType}`);
  }
}

function addScatterSeries(
  worksheet: Excel.Worksheet,
  chart: Excel.Chart,
  plan: ChartPlan,
  style: ChartStylePlan,
): void {
  plan.series.forEach((seriesPlan, index) => {
    const series = chart.series.add(seriesPlan.name, index);
    series.setXAxisValues(worksheet.getRange(localAddress(seriesPlan.categoryAddress)));
    series.setValues(worksheet.getRange(localAddress(seriesPlan.valuesAddress)));
    setSeriesColor(series, style.seriesColors[index]);
    if (plan.addLinearTrendline) {
      series.trendlines.add(Excel.ChartTrendlineType.linear);
    }
  });
}

function localAddress(address: string): string {
  const separator = address.lastIndexOf("!");
  return separator < 0 ? address : address.slice(separator + 1);
}

function applyZebraFills(range: Excel.Range, plan: Extract<TableFormatPlan, { kind: "zebra" }>): void {
  plan.rowFills.forEach((rowFill) => {
    range.getRow(rowFill.rowOffset).format.fill.color = rowFill.fill;
  });
}

function applyStandardFormats(range: Excel.Range, plan: StandardTableFormatPlan): void {
  range.format.fill.color = plan.body.fill;
  range.format.font.color = plan.body.fontColor;
  range.format.font.bold = plan.body.bold;
  range.format.font.size = plan.body.fontSize;

  plan.columnAlignments.forEach((alignment, column) => {
    range.getColumn(column).format.horizontalAlignment = excelAlignment(alignment);
  });
  applyBorders(range, plan);

  const header = range.getRow(0);
  header.format.fill.color = plan.header.fill;
  header.format.font.color = plan.header.fontColor;
  header.format.font.bold = plan.header.bold;
  header.format.font.size = plan.header.fontSize;
  header.format.horizontalAlignment = excelAlignment(plan.header.horizontalAlignment);
}

function excelAlignment(alignment: TableHorizontalAlignment): "Left" | "Center" | "Right" {
  switch (alignment) {
    case "left": return "Left";
    case "center": return "Center";
    case "right": return "Right";
  }
}

function applyBorders(range: Excel.Range, plan: StandardTableFormatPlan): void {
  const borderIndexes = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"] as const;
  borderIndexes.forEach((index) => {
    const border = range.format.borders.getItem(index);
    border.color = plan.border.color;
    border.style = "Continuous";
    border.weight = "Thin";
  });
}

async function autofitWithinBounds(
  context: Excel.RequestContext,
  range: Excel.Range,
  plan: StandardTableFormatPlan,
): Promise<void> {
  const columns = Array.from({ length: plan.columnCount }, (_, index) => range.getColumn(index).format);
  const rows = Array.from({ length: plan.rowCount }, (_, index) => range.getRow(index).format);

  columns.forEach((format) => {
    format.autofitColumns();
    format.load("columnWidth");
  });
  rows.forEach((format) => {
    format.autofitRows();
    format.load("rowHeight");
  });
  await context.sync();

  columns.forEach((format) => {
    if (format.columnWidth > 180) {
      format.columnWidth = 180;
    }
  });
  rows.forEach((format) => {
    if (format.rowHeight > 45) {
      format.rowHeight = 45;
    }
  });
  await context.sync();
}

function applySeriesColors(series: readonly Excel.ChartSeries[], colors: readonly string[]): void {
  series.forEach((item, index) => setSeriesColor(item, colors[index]));
}

function setSeriesColor(series: Excel.ChartSeries, color: string | undefined): void {
  if (color === undefined) {
    return;
  }
  series.format.fill.setSolidColor(color);
  series.format.line.color = color;
}

async function colorPiePoints(
  context: Excel.RequestContext,
  chart: Excel.Chart,
): Promise<void> {
  const pieSeries = chart.series.items[0];
  if (pieSeries === undefined) {
    return;
  }
  pieSeries.points.load("items");
  await context.sync();
  pieSeries.points.items.forEach((point, index) => {
    const color = CICC_SERIES_COLORS[index % CICC_SERIES_COLORS.length];
    if (color !== undefined) {
      point.format.fill.setSolidColor(color);
    }
  });
}

function applyChartStyle(
  chart: Excel.Chart,
  plan: ChartPlan,
  style: ChartStylePlan,
  sourceRange: Excel.Range,
): void {
  chart.width = style.widthPoints;
  chart.height = style.heightPoints;
  chart.left = style.placement.side === "right"
    ? sourceRange.left + sourceRange.width + style.placement.gutterPoints
    : sourceRange.left;
  chart.top = style.placement.side === "below"
    ? sourceRange.top + sourceRange.height + style.placement.gutterPoints
    : sourceRange.top;

  chart.title.visible = style.showTitle;

  chart.legend.visible = style.legendPosition !== "none";
  if (style.legendPosition !== "none") {
    chart.legend.position = resolveLegendPosition(style.legendPosition);
  }

  chart.dataLabels.showValue = style.showDataLabels;
  chart.format.fill.setSolidColor(style.chartAreaFill);
  chart.plotArea.format.fill.setSolidColor(style.plotAreaFill);
  if (!style.showOuterBorder) {
    chart.format.border.clear();
  }
  chart.format.font.size = style.textSizePoints;
  chart.title.format.font.size = style.textSizePoints;
  chart.legend.format.font.size = style.textSizePoints;
  chart.dataLabels.format.font.size = style.textSizePoints;
  if (plan.kind !== "pie" && plan.kind !== "pieExploded") {
    chart.axes.categoryAxis.format.font.size = style.textSizePoints;
    chart.axes.valueAxis.format.font.size = style.textSizePoints;
    chart.axes.valueAxis.majorGridlines.format.line.color = style.majorGridlineColor;

    if (style.valueAxisNumberFormat !== undefined) {
      chart.axes.valueAxis.numberFormat = style.valueAxisNumberFormat;
    }
    if (style.categoryAxisNumberFormat !== undefined) {
      chart.axes.categoryAxis.numberFormat = style.categoryAxisNumberFormat;
    }
  }
}

function resolveLegendPosition(position: Exclude<ChartStylePlan["legendPosition"], "none">): Excel.ChartLegendPosition {
  switch (position) {
    case "top": return Excel.ChartLegendPosition.top;
    case "bottom": return Excel.ChartLegendPosition.bottom;
    case "left": return Excel.ChartLegendPosition.left;
    case "right": return Excel.ChartLegendPosition.right;
  }
}
