import { describe, expect, it } from "vitest";
import { parseSelection, quoteSheetName } from "../../src/core/selectionParser";

const snapshot = {
  worksheetName: "Data",
  address: "Data!A1:C4",
  rowIndex: 0,
  columnIndex: 0,
  rowCount: 4,
  columnCount: 3,
  values: [["Date", "Pulp", "Paper"], [1, 10, 20], [2, 11, 21], [3, 12, 22]],
  texts: [
    ["Date", "Pulp", "Paper"],
    ["2026-01-01", "10", "20"],
    ["2026-01-08", "11", "21"],
    ["2026-01-15", "12", "22"],
  ],
  numberFormats: [
    ["General", "General", "General"],
    ["yyyy-mm-dd", "0", "0"],
    ["yyyy-mm-dd", "0", "0"],
    ["yyyy-mm-dd", "0", "0"],
  ],
};

describe("parseSelection", () => {
  it("uses row one as headers and column one as categories", () => {
    const parsed = parseSelection(snapshot, "columns");

    expect(parsed.headerRows).toBe(1);
    expect(parsed.categoryColumnOffset).toBe(0);
    expect(parsed.series.map((series) => series.name)).toEqual(["Pulp", "Paper"]);
    expect(parsed.categoryAddress).toBe("'Data'!$A$2:$A$4");
    expect(parsed.series.map((series) => series.valuesAddress)).toEqual([
      "'Data'!$B$2:$B$4",
      "'Data'!$C$2:$C$4",
    ]);
  });

  it("treats a single-cell text row as a title and the next row as headers", () => {
    const withTitle = {
      ...snapshot,
      address: "Data!A1:C5",
      rowCount: 5,
      values: [["Weekly prices", null, null], ...snapshot.values],
      texts: [["Weekly prices", "", ""], ...snapshot.texts],
      numberFormats: [["General", "General", "General"], ...snapshot.numberFormats],
    };

    const parsed = parseSelection(withTitle, "columns");

    expect(parsed.title).toBe("Weekly prices");
    expect(parsed.headerRows).toBe(2);
  });

  it("rejects a one-cell selection", () => {
    const oneCell = {
      ...snapshot,
      address: "Data!A1",
      rowCount: 1,
      columnCount: 1,
      values: [[1]],
      texts: [["1"]],
      numberFormats: [["0"]],
    };

    expect(() => parseSelection(oneCell)).toThrow("selection_too_small");
  });

  it("rejects matrices that do not match the declared rectangle", () => {
    const malformed = { ...snapshot, texts: snapshot.texts.slice(0, 3) };

    expect(() => parseSelection(malformed)).toThrow("invalid_selection");
  });

  it("quotes sheet names and doubles embedded apostrophes", () => {
    expect(quoteSheetName("O'Brien Data")).toBe("'O''Brien Data'");
  });
});
