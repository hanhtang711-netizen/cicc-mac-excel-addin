import { getCellKind } from "../core/selectionParser";
import type {
  SelectionSnapshot,
  StandardTableFormatPlan,
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
      fontSize: 8,
    },
    body: { fill: "#FFFFFF", fontColor: "#000000", bold: false, fontSize: 8 },
    horizontalAlignment: "left",
    verticalAlignment: "center",
    wrapText: true,
    rowHeight: 16,
    clearBorders: true,
    hideWorksheetGridlines: true,
    rowFills: [],
    preserve: [...STANDARD_PRESERVE],
  };
}

export function buildZebraPlan(snapshot: SelectionSnapshot): ZebraTableFormatPlan {
  const firstDataRow = hasHeader(snapshot) ? 1 : 0;
  const rowFills = Array.from({ length: Math.max(snapshot.rowCount - firstDataRow, 0) }, (_, index) => ({
    rowOffset: firstDataRow + index,
    fill: index % 2 === 0 ? "#F5F5F5" : "#FFFFFF",
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
