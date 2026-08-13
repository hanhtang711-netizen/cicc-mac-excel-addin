import { AddinError } from "../core/errors";
import type {
  ChartPlan,
  ChartStylePlan,
  StandardTableFormatPlan,
  SelectionSnapshot,
  TableFormatPlan,
} from "../core/types";
import { CICC_SERIES_COLORS } from "../charts/chartStyle";

type ChartWriteStage = "create" | "series-style" | "common-style" | "final-sync";

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

  /** 桥接自动执行：按工作表名+地址直读快照，不经过"当前选区"。
   *
   * 坑位（260814）：select() + 立即读当前选区存在 UI 竞态——select 尚未
   * 生效时 getSelectedRange() 抛 RichApi ItemNotFound（AMZN 18-sheet 底稿
   * 复现）。直读 getRange(address) 不依赖选区/激活，无竞态。 */
  async readRange(sheetName: string, address: string): Promise<SelectionSnapshot> {
    try {
      return await this.readRangeOnce(sheetName, address);
    } catch (cause) {
      // 坑位（260814）：连续两个 Excel.run 批次时，第二个批次的
      // worksheets.getItem 会命中 Excel 未刷新的集合缓存抛 ItemNotFound。
      // 等 1 秒让缓存刷新后重试一次——重试期间 getItem 必然已就绪。
      await new Promise((resolve) => setTimeout(resolve, 1000));
      try {
        return await this.readRangeOnce(sheetName, address);
      } catch (retryCause) {
        // 诊断包装：两次失败位置与参数一并回报
        throw new AddinError("excel_runtime_error", {
          stage: "readRange",
          sheet: sheetName,
          address,
          cause,
          retryCause,
        });
      }
    }
  }

  private async readRangeOnce(sheetName: string, address: string): Promise<SelectionSnapshot> {
    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const range = sheet.getRange(address);
      range.load("address,rowIndex,columnIndex,rowCount,columnCount,values,text,numberFormat");
      await context.sync();
      return snapshotFromRange(sheetName, range);
    });
  }

  async createChart(plan: ChartPlan, style: ChartStylePlan): Promise<void> {
    await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(plan.worksheetName);
      const sourceRange = worksheet.getRange(localAddress(plan.sourceAddress));
      let chart: Excel.Chart | undefined;
      let stage: ChartWriteStage = "create";
      try {
        sourceRange.load("left,top,width,height");
        await context.sync();
        chart = worksheet.charts.add(
          resolveChartType(plan.excelType),
          sourceRange,
          plan.orientation === "rows" ? Excel.ChartSeriesBy.rows : Excel.ChartSeriesBy.columns,
        );

        chart.series.load("items");
        await context.sync();

        // These are non-negotiable CICC defaults. Write them in isolated
        // batches before optional series/axis styling: on Mac, an unsupported
        // styling property must not leave Excel's placeholder title or labels.
        await applyEssentialChartCleanup(context, chart, style);

        // A chart that Excel has created successfully is more valuable than a
        // perfect-looking chart that is rolled back because one optional style
        // property is unavailable on a particular Mac build. Keep creation
        // fatal, but make all CICC visual refinements best-effort.
        try {
          stage = "series-style";
          if (plan.kind === "scatterTrend") {
            for (const existingSeries of chart.series.items) {
              existingSeries.delete();
            }
            addScatterSeries(worksheet, chart, plan, style);
          } else {
            // 坑位（260813 实测 Mac Excel 16.112）：charts.add 会把类别列
            // （columns）或表头行（rows）建成"幽灵系列"——系列数多 1，幽灵
            // 位于集合末尾，带工作簿主题 accent4 色（默认主题下是紫 #8064A2，
            // 不属于中金色板）。删除必须走独立 sync 批次：同批次 delete 会让
            // 剩余系列引用失效，导致整批回滚（260813 第一次修复翻车教训）。
            let seriesItems = chart.series.items;
            if (seriesItems.length === plan.series.length + 1) {
              const ghost = seriesItems[seriesItems.length - 1];
              ghost.delete();
              await context.sync();
              chart.series.load("items");
              await context.sync();
              seriesItems = chart.series.items;
            }
            applySeriesStyles(seriesItems, plan, style);
          }

          if (plan.kind === "pie" || plan.kind === "pieExploded") {
            await colorPiePoints(context, chart);
          }

          stage = "common-style";
          applyChartStyle(chart, plan, style, sourceRange);
          stage = "final-sync";
          await context.sync();
        } catch {
          // Leave the native chart in place if a non-essential format request
          // fails. This is required for cross-version Mac Excel compatibility.
        }

        await applyEssentialChartCleanup(context, chart, style);
      } catch (cause) {
        const failureStage = stage;
        if (chart !== undefined) {
          try {
            chart.delete();
            await context.sync();
          } catch (rollbackError) {
            throw new AddinError("excel_runtime_error", {
              stage: failureStage,
              cause,
              rollbackError,
            });
          }
        }
        throw new AddinError("excel_runtime_error", { stage: failureStage, cause });
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
      if (plan.hideWorksheetGridlines) {
        worksheet.showGridlines = false;
      }
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
    const color = style.seriesColors[index];
    series.setXAxisValues(worksheet.getRange(localAddress(seriesPlan.categoryAddress)));
    series.setValues(worksheet.getRange(localAddress(seriesPlan.valuesAddress)));
    applySeriesStyle(series, color, plan, style);
    if (plan.addLinearTrendline) {
      const trendline = series.trendlines.add(Excel.ChartTrendlineType.linear);
      if (color !== undefined) {
        trendline.format.line.color = color;
        trendline.format.line.weight = style.lineWidthPoints;
      }
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
  // Office.js 只有西文字体槽 API（font.name 写 latin）；中文字体槽（ea=黑体）
  // 由工作簿主题提供——底稿须加载中金配色主题（见 webext_inject 主题注入）。
  range.format.font.name = "Arial";

  range.format.horizontalAlignment = "Left";
  range.format.verticalAlignment = "Center";
  range.format.wrapText = plan.wrapText;
  range.format.rowHeight = plan.rowHeight;
  clearBorders(range);

  const header = range.getRow(0);
  header.format.fill.color = plan.header.fill;
  header.format.font.color = plan.header.fontColor;
  header.format.font.bold = plan.header.bold;
  header.format.font.size = plan.header.fontSize;
  header.format.font.name = "Arial";
}

function clearBorders(range: Excel.Range): void {
  const borderIndexes = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"] as const;
  borderIndexes.forEach((index) => {
    const border = range.format.borders.getItem(index);
    border.style = "None";
  });
}

async function autofitWithinBounds(
  context: Excel.RequestContext,
  range: Excel.Range,
  plan: StandardTableFormatPlan,
): Promise<void> {
  const columns = Array.from({ length: plan.columnCount }, (_, index) => range.getColumn(index).format);
  columns.forEach((format) => {
    format.autofitColumns();
    format.load("columnWidth");
  });
  await context.sync();

  columns.forEach((format) => {
    if (format.columnWidth > 180) {
      format.columnWidth = 180;
    }
  });
  await context.sync();
}

function applySeriesStyles(
  series: readonly Excel.ChartSeries[],
  plan: ChartPlan,
  style: ChartStylePlan,
): void {
  series.forEach((item, index) => applySeriesStyle(item, style.seriesColors[index], plan, style));
}

function applySeriesStyle(
  series: Excel.ChartSeries,
  color: string | undefined,
  plan: ChartPlan,
  style: ChartStylePlan,
): void {
  if (color === undefined) {
    return;
  }

  if (style.seriesStyle === "fill-no-border") {
    series.format.fill.setSolidColor(color);
    // Mac Excel may reject line.clear() for a newly-created column series.
    // The rejected batch can then leave later series in theme colours. Excel
    // defaults column outlines to none, so only set the reliable fill here.
    return;
  }
  if (style.seriesStyle === "pie-points") {
    return;
  }

  series.format.line.color = color;
  series.format.line.weight = style.lineWidthPoints;
  series.smooth = style.smoothLines;
  series.markerStyle = plan.kind === "lineMarkers" || style.seriesStyle === "scatter"
    ? "Automatic"
    : "None";
  if (series.markerStyle !== "None") {
    series.markerBackgroundColor = color;
    series.markerForegroundColor = color;
  }
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
      point.format.border.color = color;
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

  chart.legend.visible = style.legendPosition !== "none";
  if (style.legendPosition !== "none") {
    chart.legend.position = resolveLegendPosition(style.legendPosition);
    chart.legend.overlay = style.legendOverlay;
  }

  chart.format.fill.setSolidColor(style.chartAreaFill);
  chart.plotArea.format.fill.setSolidColor(style.plotAreaFill);
  if (!style.showOuterBorder) {
    chart.format.border.clear();
  }
  applyChartFont(chart.format.font, style.textSizePoints);
  applyChartFont(chart.title.format.font, style.textSizePoints);
  applyChartFont(chart.legend.format.font, style.legendFontSizePoints);
  applyChartFont(chart.dataLabels.format.font, style.textSizePoints);
  if (plan.kind !== "pie" && plan.kind !== "pieExploded") {
    applyAxisStyle(
      chart.axes.categoryAxis,
      "Minimum",
      "Outside",
      style.categoryAxisNumberFormat,
      style,
    );
    applyAxisStyle(
      chart.axes.valueAxis,
      "Automatic",
      "None",
      style.valueAxisNumberFormat,
      style,
    );
  }
}

async function applyEssentialChartCleanup(
  context: Excel.RequestContext,
  chart: Excel.Chart,
  style: ChartStylePlan,
): Promise<void> {
  if (!style.showTitle) {
    try {
      chart.title.text = "";
      await context.sync();
    } catch {
      // Continue with the visibility write below.
    }
  }

  try {
    // This must be the last title write. Assigning title.text can make Excel
    // recreate the default title placeholder on Mac.
    chart.title.visible = style.showTitle;
    await context.sync();
  } catch {
    // The chart remains usable even if its host does not expose this setting.
  }

  try {
    chart.dataLabels.showValue = style.showDataLabels;
    await context.sync();
  } catch {
    // The chart remains usable even if its host does not expose this setting.
  }
}

function applyChartFont(font: Excel.ChartFont, size: number): void {
  font.name = "Arial";
  font.size = size;
  font.color = "#000000";
}

function applyAxisStyle(
  axis: Excel.ChartAxis,
  position: "Minimum" | "Automatic",
  tickMark: "Outside" | "None",
  numberFormat: string | undefined,
  style: ChartStylePlan,
): void {
  axis.visible = true;
  axis.position = position;
  axis.majorTickMark = tickMark;
  axis.tickLabelPosition = "NextToAxis";
  axis.title.visible = false;
  axis.majorGridlines.visible = style.showGridlines;
  axis.minorGridlines.visible = false;
  axis.format.line.color = "#BFBFBF";
  axis.format.line.weight = 0.75;
  applyChartFont(axis.format.font, style.textSizePoints);
  if (numberFormat !== undefined) {
    axis.numberFormat = numberFormat;
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
