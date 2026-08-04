import { describe, expect, it } from "vitest";
import { buildStandardTablePlan, buildZebraPlan } from "../../src/tables/tablePlanner";
import type { SelectionSnapshot } from "../../src/core/types";

const selectionSnapshot: SelectionSnapshot = {
  worksheetName: "Data",
  address: "'Data'!$A$1:$C$4",
  rowIndex: 0,
  columnIndex: 0,
  rowCount: 4,
  columnCount: 3,
  values: [["Date", "Revenue", "Margin"], [46023, 10, 0.1], [46030, 12, 0.12], [46037, 11, 0.11]],
  texts: [
    ["Date", "Revenue", "Margin"],
    ["2026-01-01", "10", "10.0%"],
    ["2026-01-08", "12", "12.0%"],
    ["2026-01-15", "11", "11.0%"],
  ],
  numberFormats: [
    ["General", "General", "General"],
    ["yyyy-mm-dd", "0", "0.0%"],
    ["yyyy-mm-dd", "0", "0.0%"],
    ["yyyy-mm-dd", "0", "0.0%"],
  ],
};

describe("table plans", () => {
  it("formats a red header and white body without zebra actions", () => {
    const plan = buildStandardTablePlan(selectionSnapshot);

    expect(plan).toMatchObject({
      kind: "standard",
      worksheetName: "Data",
      address: "'Data'!$A$1:$C$4",
      header: {
        fill: "#8A2626",
        fontColor: "#FFFFFF",
        bold: true,
        fontSize: 10.5,
        horizontalAlignment: "center",
      },
      body: { fill: "#FFFFFF", fontColor: "#000000", bold: false, fontSize: 10.5 },
      border: { color: "#D9D9D9", style: "continuous", weight: "thin" },
      columnAlignments: ["center", "right", "right"],
      preserve: ["values", "formulas", "numberFormats", "merges", "conditionalFormats"],
    });
    expect(plan.rowFills).toEqual([]);
  });

  it("zebra plan changes fills only and skips a detected header", () => {
    const plan = buildZebraPlan(selectionSnapshot);

    expect(plan).toMatchObject({
      kind: "zebra",
      worksheetName: "Data",
      address: "'Data'!$A$1:$C$4",
      rowFills: [
        { rowOffset: 1, fill: "#FFFFFF" },
        { rowOffset: 2, fill: "#F5F5F5" },
        { rowOffset: 3, fill: "#FFFFFF" },
      ],
      preserve: [
        "values",
        "formulas",
        "numberFormats",
        "merges",
        "conditionalFormats",
        "fonts",
        "borders",
        "alignment",
      ],
    });
  });

  it("includes the first row in zebra fills when it is not a header", () => {
    const plan = buildZebraPlan({
      ...selectionSnapshot,
      values: [[1, 10, 0.1], [2, 12, 0.12]],
      texts: [["1", "10", "10.0%"], ["2", "12", "12.0%"]],
      numberFormats: [["0", "0", "0.0%"], ["0", "0", "0.0%"]],
      rowCount: 2,
    });

    expect(plan.rowFills).toEqual([
      { rowOffset: 0, fill: "#FFFFFF" },
      { rowOffset: 1, fill: "#F5F5F5" },
    ]);
  });
});
