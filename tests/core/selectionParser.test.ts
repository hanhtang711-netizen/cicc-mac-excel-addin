import { describe, expect, it } from "vitest";
import { getCellKind, parseSelection, quoteSheetName } from "../../src/core/selectionParser";

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
    expect(parsed.categoryKinds).toEqual(["date", "date", "date"]);
    expect(parsed.series.map((series) => series.valueKinds)).toEqual([
      ["number", "number", "number"],
      ["number", "number", "number"],
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

  it("keeps the first row of a numeric selection when no header is present", () => {
    const numericSelection = {
      ...snapshot,
      rowCount: 3,
      columnCount: 2,
      values: [[1, 10], [2, 11], [3, 12]],
      texts: [["1", "10"], ["2", "11"], ["3", "12"]],
      numberFormats: [["0", "0"], ["0", "0"], ["0", "0"]],
    };

    const parsed = parseSelection(numericSelection, "columns");

    expect(parsed.headerRows).toBe(0);
    expect(parsed.title).toBeUndefined();
    expect(parsed.categoryAddress).toBe("'Data'!$A$1:$A$3");
    expect(parsed.series[0].valuesAddress).toBe("'Data'!$B$1:$B$3");
  });

  it("does not treat a numeric or date display text as a title", () => {
    const numericAndDate = {
      ...snapshot,
      rowCount: 2,
      columnCount: 2,
      values: [[45292, null], [45299, 10]],
      texts: [["2024-01-01", ""], ["2024-01-08", "10"]],
      numberFormats: [["yyyy-mm-dd", "General"], ["yyyy-mm-dd", "0"]],
    };

    const parsed = parseSelection(numericAndDate, "columns");

    expect(parsed.title).toBeUndefined();
    expect(parsed.headerRows).toBe(0);
  });

  it("rejects a title and header selection that has no data rows", () => {
    const titleAndHeaderOnly = {
      ...snapshot,
      rowCount: 2,
      columnCount: 2,
      values: [["Weekly prices", null], ["Date", "Pulp"]],
      texts: [["Weekly prices", ""], ["Date", "Pulp"]],
      numberFormats: [["General", "General"], ["General", "General"]],
    };

    expect(() => parseSelection(titleAndHeaderOnly, "columns")).toThrow("unsupported_layout");
  });
});

describe("getCellKind", () => {
  it("recognizes date tokens while ignoring quoted and escaped numeric-format literals", () => {
    expect(getCellKind(46023, "2026-01-01", "yyyy-mm-dd")).toBe("date");
    expect(getCellKind(100, "100.0", "\"USD\" #,##0.0")).toBe("number");
    expect(getCellKind(100, "100.0", "\"total \"\"m\"\"\" #,##0.0")).toBe("number");
    expect(getCellKind(100, "100.0", "0.0 \\m")).toBe("number");
  });
});
