export type SeriesOrientation = "columns" | "rows";
export type SeriesOrientationMode = "auto" | SeriesOrientation;

export type ChartKind =
  | "column"
  | "columnStacked"
  | "line"
  | "lineMarkers"
  | "pie"
  | "bar"
  | "scatterTrend"
  | "pieExploded"
  | "columnStacked100"
  | "lineStacked";

export type ChartSizePreset = "small" | "medium" | "large" | "custom";

export type LegendPosition = "bottom" | "top" | "left" | "right" | "none";

export type CellKind = "blank" | "text" | "number" | "date" | "error";

export interface SelectionSnapshot {
  worksheetName: string;
  address: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  values: unknown[][];
  texts: string[][];
  numberFormats: string[][];
}

export interface ParsedSeries {
  name: string;
  valuesAddress: string;
  valueColumnOffset: number;
  valueKinds: CellKind[];
}

export interface ParsedSelection {
  worksheetName: string;
  sourceAddress: string;
  title?: string;
  headerRows: number;
  categoryColumnOffset: number;
  categoryAddress: string;
  categoryKinds: CellKind[];
  orientation: SeriesOrientation;
  series: ParsedSeries[];
  numberFormats: string[][];
}

export interface ChartOptions {
  title?: string;
  sizePreset?: ChartSizePreset;
  widthCm?: number;
  heightCm?: number;
  legendPosition?: LegendPosition;
  orientation?: SeriesOrientationMode;
  showDataLabels?: boolean;
  addTrendline?: boolean;
}

export interface ChartSeriesPlan {
  name: string;
  categoryAddress: string;
  valuesAddress: string;
}

export interface ChartPlan {
  kind: ChartKind;
  excelType: string;
  worksheetName: string;
  sourceAddress: string;
  orientation: SeriesOrientation;
  series: ChartSeriesPlan[];
  addLinearTrendline: boolean;
}

export type ChartSeriesStyle = "fill-no-border" | "line" | "scatter" | "pie-points";

export interface ChartPlacementPlan {
  side: "right" | "below";
  gutterPoints: number;
}

export interface ChartStyleInput {
  kind: ChartKind;
  seriesCount: number;
  sourceFormat: string;
  categoryFormat?: string;
  selectionColumn?: number;
  selectionColumnCount?: number;
  estimatedChartColumns?: number;
}

export interface ChartStylePlan {
  seriesColors: string[];
  seriesStyle: ChartSeriesStyle;
  widthPoints: number;
  heightPoints: number;
  legendPosition: LegendPosition;
  legendOverlay: boolean;
  legendFontSizePoints: number;
  chartAreaFill: string;
  plotAreaFill: string;
  showOuterBorder: boolean;
  showTitle: boolean;
  showDataLabels: boolean;
  showGridlines: boolean;
  lineWidthPoints: number;
  smoothLines: boolean;
  textSizePoints: number;
  majorGridlineColor: string;
  valueAxisNumberFormat: string | undefined;
  categoryAxisNumberFormat: string | undefined;
  warnings: string[];
  placement: ChartPlacementPlan;
}

export type TablePreservedProperty =
  | "values"
  | "formulas"
  | "numberFormats"
  | "merges"
  | "conditionalFormats"
  | "fonts"
  | "borders"
  | "alignment";

export type TableHorizontalAlignment = "left" | "center" | "right";

export interface TableCellFormat {
  fill: string;
  fontColor: string;
  bold: boolean;
  fontSize: number;
}

export interface TableBorderFormat {
  color: string;
  style: "continuous";
  weight: "thin";
}

interface TablePlanBase {
  worksheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
}

export interface StandardTableFormatPlan extends TablePlanBase {
  kind: "standard";
  header: TableCellFormat & { horizontalAlignment: "center" };
  body: TableCellFormat;
  columnAlignments: TableHorizontalAlignment[];
  border: TableBorderFormat;
  rowFills: [];
  preserve: Array<Exclude<TablePreservedProperty, "fonts" | "borders" | "alignment">>;
}

export interface ZebraTableFormatPlan extends TablePlanBase {
  kind: "zebra";
  rowFills: Array<{ rowOffset: number; fill: "#FFFFFF" | "#F5F5F5" }>;
  preserve: TablePreservedProperty[];
}

export type TableFormatPlan = StandardTableFormatPlan | ZebraTableFormatPlan;

export interface FeedbackPort {
  showError(error: unknown): Promise<void>;
  showWarnings(codes: string[]): Promise<void>;
}
