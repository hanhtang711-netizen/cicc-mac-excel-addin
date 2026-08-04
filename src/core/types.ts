export type SeriesOrientation = "columns" | "rows";

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
  orientation?: SeriesOrientation;
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
  sizePreset: ChartSizePreset;
  widthCm?: number;
  heightCm?: number;
  title?: string;
  showDataLabels: boolean;
  addLinearTrendline: boolean;
}

export interface ChartPlacementPlan {
  side: "right" | "below";
  gutterPoints: number;
}

export interface ChartStyleInput {
  seriesCount: number;
  options: ChartOptions;
  sourceFormat: string;
  categoryFormat?: string;
  selectionColumn?: number;
  selectionColumnCount?: number;
  estimatedChartColumns?: number;
}

export interface ChartStylePlan {
  seriesColors: string[];
  widthPoints: number;
  heightPoints: number;
  legendPosition: LegendPosition;
  chartAreaFill: string;
  plotAreaFill: string;
  showOuterBorder: boolean;
  textSizePoints: number;
  majorGridlineColor: string;
  valueAxisNumberFormat: string | undefined;
  categoryAxisNumberFormat: string | undefined;
  warnings: string[];
  placement: ChartPlacementPlan;
}

export interface FeedbackPort {
  showError(error: unknown): Promise<void>;
  showWarnings(codes: string[]): Promise<void>;
}
