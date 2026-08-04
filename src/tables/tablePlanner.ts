import { getCellKind } from "../core/selectionParser";
import type {
  CellKind,
  SelectionSnapshot,
  StandardTableFormatPlan,
  TableHorizontalAlignment,
  ZebraTableFormatPlan,
} from "../core/types";

const STANDARD_PRESERVE = ["values", "formulas", "numberFormats", "merges", "conditionalFormats"] as const;
const ZEBRA_PRESERVE = [...STANDARD_PRESERVE, "fonts", "borders", "alignment"] as const;

export function buildStandardTablePlan(snapshot: SelectionSnapshot): StandardTableFormatPlan {
  return {
    kind: "standard",
    worksheetName: snapshot.worksheetName,
    address: snapshot.address,
    rowCount: snapshot.rowCount,
    columnCount: snapshot.columnCount,
    header: {
      fill: "#8A2626",
      fontColor: "#FFFFFF",
      bold: true,
      fontSize: 10.5,
      horizontalAlignment: "center",
    },
    body: { fill: "#FFFFFF", fontColor: "#000000", bold: false, fontSize: 10.5 },
    columnAlignments: Array.from({ length: snapshot.columnCount }, (_, column) => columnAlignment(snapshot, column)),
    border: { color: "#D9D9D9", style: "continuous", weight: "thin" },
    rowFills: [],
    preserve: [...STANDARD_PRESERVE],
  };
}

export function buildZebraPlan(snapshot: SelectionSnapshot): ZebraTableFormatPlan {
  const firstDataRow = hasHeader(snapshot) ? 1 : 0;
  const rowFills = Array.from({ length: Math.max(snapshot.rowCount - firstDataRow, 0) }, (_, index) => ({
    rowOffset: firstDataRow + index,
    fill: index % 2 === 0 ? "#FFFFFF" : "#F5F5F5",
  } as const));

  return {
    kind: "zebra",
    worksheetName: snapshot.worksheetName,
    address: snapshot.address,
    rowCount: snapshot.rowCount,
    columnCount: snapshot.columnCount,
    rowFills,
    preserve: [...ZEBRA_PRESERVE],
  };
}

function columnAlignment(snapshot: SelectionSnapshot, column: number): TableHorizontalAlignment {
  const kinds = dataCellKinds(snapshot, column);
  if (kinds.length > 0 && kinds.every((kind) => kind === "date")) {
    return "center";
  }
  if (kinds.length > 0 && kinds.every((kind) => kind === "number")) {
    return "right";
  }
  return "left";
}

function hasHeader(snapshot: SelectionSnapshot): boolean {
  if (snapshot.rowCount < 2 || snapshot.columnCount < 1) {
    return false;
  }

  const headerKinds = snapshot.values[0].map((value, column) =>
    getCellKind(value, snapshot.texts[0][column], snapshot.numberFormats[0][column]),
  );
  const textCount = headerKinds.filter((kind) => kind === "text").length;
  const hasTextMajority = textCount > snapshot.columnCount / 2;
  const hasFollowingData = snapshot.values.slice(1).some((row, rowOffset) =>
    row.some((value, column) =>
      getCellKind(value, snapshot.texts[rowOffset + 1][column], snapshot.numberFormats[rowOffset + 1][column]) !== "blank"),
  );

  return hasTextMajority && hasFollowingData;
}

function dataCellKinds(snapshot: SelectionSnapshot, column: number): CellKind[] {
  return snapshot.values.slice(1)
    .map((row, rowOffset) =>
      getCellKind(row[column], snapshot.texts[rowOffset + 1][column], snapshot.numberFormats[rowOffset + 1][column]))
    .filter((kind) => kind !== "blank");
}
