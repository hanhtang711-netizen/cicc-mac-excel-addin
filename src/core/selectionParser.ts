import { AddinError } from "./errors";
import type {
  CellKind,
  ParsedSelection,
  ParsedSeries,
  SelectionSnapshot,
  SeriesOrientation,
} from "./types";

export function quoteSheetName(name: string): string {
  return `'${name.replaceAll("'", "''")}'`;
}

export function parseSelection(
  snapshot: SelectionSnapshot,
  orientation: SeriesOrientation = "columns",
): ParsedSelection {
  assertRectangular(snapshot);

  if (snapshot.rowCount < 2 || snapshot.columnCount < 2) {
    throw new AddinError("selection_too_small", {
      rowCount: snapshot.rowCount,
      columnCount: snapshot.columnCount,
    });
  }

  const title = titleInFirstRow(snapshot);
  const titleRows = title === undefined ? 0 : 1;
  const headerRow = titleRows;
  const hasHeader = isHeaderRow(snapshot, headerRow);
  const headerRows = titleRows + (hasHeader ? 1 : 0);

  if (orientation === "columns") {
    return parseColumns(snapshot, title, headerRows, hasHeader);
  }
  return parseRows(snapshot, title, headerRows, hasHeader);
}

function parseColumns(
  snapshot: SelectionSnapshot,
  title: string | undefined,
  headerRows: number,
  hasHeader: boolean,
): ParsedSelection {
  const dataStartRow = snapshot.rowIndex + headerRows;
  const dataEndRow = snapshot.rowIndex + snapshot.rowCount - 1;
  const categoryColumn = snapshot.columnIndex;
  const series: ParsedSeries[] = [];

  if (dataStartRow > dataEndRow) {
    throw new AddinError("unsupported_layout");
  }

  for (let offset = 1; offset < snapshot.columnCount; offset += 1) {
    const column = snapshot.columnIndex + offset;
    series.push({
      name: hasHeader ? displayText(snapshot.texts[headerRows - 1][offset]) : `Series ${offset}`,
      valuesAddress: rangeAddress(snapshot.worksheetName, dataStartRow, column, dataEndRow, column),
      valueColumnOffset: offset,
      valueKinds: columnKinds(snapshot, headerRows, offset),
    });
  }

  return {
    worksheetName: snapshot.worksheetName,
    sourceAddress: snapshot.address,
    ...(title === undefined ? {} : { title }),
    headerRows,
    categoryColumnOffset: 0,
    categoryAddress: rangeAddress(snapshot.worksheetName, dataStartRow, categoryColumn, dataEndRow, categoryColumn),
    categoryKinds: columnKinds(snapshot, headerRows, 0),
    orientation: "columns",
    series,
    numberFormats: snapshot.numberFormats,
  };
}

function parseRows(
  snapshot: SelectionSnapshot,
  title: string | undefined,
  headerRows: number,
  hasHeader: boolean,
): ParsedSelection {
  const firstDataRow = snapshot.rowIndex + headerRows;
  const endColumn = snapshot.columnIndex + snapshot.columnCount - 1;
  const categoryRowOffset = hasHeader ? headerRows - 1 : 0;
  const categoryRow = snapshot.rowIndex + categoryRowOffset;
  const series: ParsedSeries[] = [];

  for (let offset = headerRows; offset < snapshot.rowCount; offset += 1) {
    const row = snapshot.rowIndex + offset;
    series.push({
      name: displayText(snapshot.texts[offset][0]) || `Series ${offset - headerRows + 1}`,
      valuesAddress: rangeAddress(snapshot.worksheetName, row, snapshot.columnIndex + 1, row, endColumn),
      valueColumnOffset: 1,
      valueKinds: rowKinds(snapshot, offset, 1),
    });
  }

  if (firstDataRow > snapshot.rowIndex + snapshot.rowCount - 1) {
    throw new AddinError("unsupported_layout");
  }

  return {
    worksheetName: snapshot.worksheetName,
    sourceAddress: snapshot.address,
    ...(title === undefined ? {} : { title }),
    headerRows,
    categoryColumnOffset: 0,
    categoryAddress: rangeAddress(
      snapshot.worksheetName,
      categoryRow,
      snapshot.columnIndex + 1,
      categoryRow,
      endColumn,
    ),
    categoryKinds: rowKinds(snapshot, categoryRowOffset, 1),
    orientation: "rows",
    series,
    numberFormats: snapshot.numberFormats,
  };
}

function assertRectangular(snapshot: SelectionSnapshot): void {
  const matrices: Array<readonly unknown[][]> = [snapshot.values, snapshot.texts, snapshot.numberFormats];
  const isMatrix = matrices.every(
    (matrix) =>
      matrix.length === snapshot.rowCount &&
      matrix.every((row) => Array.isArray(row) && row.length === snapshot.columnCount),
  );

  if (!isMatrix || snapshot.rowCount < 1 || snapshot.columnCount < 1) {
    throw new AddinError("invalid_selection");
  }
}

function titleInFirstRow(snapshot: SelectionSnapshot): string | undefined {
  const kinds = snapshot.values[0].map((value, column) =>
    getCellKind(value, snapshot.texts[0][column], snapshot.numberFormats[0][column]),
  );
  const textColumns = kinds
    .map((kind, column) => (kind === "text" ? column : -1))
    .filter((column) => column >= 0);

  if (textColumns.length !== 1 || kinds.some((kind) => kind !== "text" && kind !== "blank")) {
    return undefined;
  }
  return displayText(snapshot.texts[0][textColumns[0]]);
}

function isHeaderRow(snapshot: SelectionSnapshot, row: number): boolean {
  if (row >= snapshot.rowCount) {
    return false;
  }

  const headerKinds = snapshot.values[row].map((value, column) =>
    getCellKind(value, snapshot.texts[row][column], snapshot.numberFormats[row][column]),
  );
  const textCount = headerKinds.filter((kind) => kind === "text").length;
  const hasTextMajority = textCount > headerKinds.length / 2;
  if (row === snapshot.rowCount - 1) {
    return hasTextMajority;
  }

  const dataCells = snapshot.values.slice(row + 1).flat();
  const dataKinds = dataCells.map((value, index) => {
    const dataRow = row + 1 + Math.floor(index / snapshot.columnCount);
    const dataColumn = index % snapshot.columnCount;
    return getCellKind(value, snapshot.texts[dataRow][dataColumn], snapshot.numberFormats[dataRow][dataColumn]);
  });
  const numericOrDateCount = dataKinds.filter((kind) => kind === "number" || kind === "date").length;

  return hasTextMajority && numericOrDateCount > dataKinds.length / 2;
}

export function getCellKind(value: unknown, text: string, numberFormat: string): CellKind {
  if (isBlank(value, text)) {
    return "blank";
  }
  if (typeof value === "string" && value.trim().startsWith("#")) {
    return "error";
  }
  if (value instanceof Date || (typeof value === "number" && isDateFormat(numberFormat))) {
    return "date";
  }
  if (typeof value === "number") {
    return "number";
  }
  return "text";
}

function columnKinds(snapshot: SelectionSnapshot, startRow: number, column: number): CellKind[] {
  return snapshot.values.slice(startRow).map((row, index) =>
    getCellKind(row[column], snapshot.texts[startRow + index][column], snapshot.numberFormats[startRow + index][column]),
  );
}

function rowKinds(snapshot: SelectionSnapshot, row: number, startColumn: number): CellKind[] {
  return snapshot.values[row].slice(startColumn).map((value, index) =>
    getCellKind(value, snapshot.texts[row][startColumn + index], snapshot.numberFormats[row][startColumn + index]),
  );
}

function isBlank(value: unknown, text: string): boolean {
  return (value === null || value === undefined || value === "") && displayText(text).length === 0;
}

function isDateFormat(numberFormat: string): boolean {
  return /(?:^|[^a-z])(d|m|y|h|s)(?:[^a-z]|$)/i.test(numberFormat);
}

function displayText(text: string): string {
  return text.trim();
}

function rangeAddress(sheetName: string, startRow: number, startColumn: number, endRow: number, endColumn: number): string {
  return `${quoteSheetName(sheetName)}!$${columnName(startColumn)}$${startRow + 1}:$${columnName(endColumn)}$${endRow + 1}`;
}

function columnName(index: number): string {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}
