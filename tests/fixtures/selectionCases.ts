import type { SelectionSnapshot } from "../../src/core/types";

const fixture = (
  address: string,
  values: unknown[][],
  texts: string[][],
  numberFormats: string[][],
): SelectionSnapshot => ({
  worksheetName: "Fixture",
  address,
  rowIndex: 0,
  columnIndex: 0,
  rowCount: values.length,
  columnCount: values[0]?.length ?? 0,
  values,
  texts,
  numberFormats,
});

export const dateSingleSeriesFixture = fixture(
  "Fixture!A1:B4",
  [["Date", "Price"], [46023, 100], [46030, 102], [46037, 101]],
  [["Date", "Price"], ["2026-01-01", "100"], ["2026-01-08", "102"], ["2026-01-15", "101"]],
  [["General", "General"], ["yyyy-mm-dd", "0"], ["yyyy-mm-dd", "0"], ["yyyy-mm-dd", "0"]],
);

export const dateMultiSeriesFixture = fixture(
  "Fixture!A1:C4",
  [["Date", "Alpha", "Beta"], [46023, 10, 20], [46030, 12, 19], [46037, 11, 22]],
  [["Date", "Alpha", "Beta"], ["2026-01-01", "10", "20"], ["2026-01-08", "12", "19"], ["2026-01-15", "11", "22"]],
  [["General", "General", "General"], ["yyyy-mm-dd", "0", "0"], ["yyyy-mm-dd", "0", "0"], ["yyyy-mm-dd", "0", "0"]],
);

export const textCategoriesFixture = fixture(
  "Fixture!A1:C4",
  [["Region", "Plan", "Actual"], ["North", 10, 9], ["East", 12, 13], ["South", 8, 8]],
  [["Region", "Plan", "Actual"], ["North", "10", "9"], ["East", "12", "13"], ["South", "8", "8"]],
  Array.from({ length: 4 }, () => ["General", "0", "0"]),
);

export const percentageFixture = fixture(
  "Fixture!A1:C4",
  [["Period", "Margin", "Mix"], [1, 0.12, 0.4], [2, 0.1, 0.45], [3, 0.14, 0.42]],
  [["Period", "Margin", "Mix"], ["Q1", "12.0%", "40.0%"], ["Q2", "10.0%", "45.0%"], ["Q3", "14.0%", "42.0%"]],
  [["General", "General", "General"], ["0", "0.0%", "0.0%"], ["0", "0.0%", "0.0%"], ["0", "0.0%", "0.0%"]],
);

export const negativeValuesFixture = fixture(
  "Fixture!A1:B4",
  [["Period", "Change"], [1, -3], [2, 2], [3, -1]],
  [["Period", "Change"], ["Q1", "-3"], ["Q2", "2"], ["Q3", "-1"]],
  [["General", "General"], ["0", "0;[Red]-0"], ["0", "0;[Red]-0"], ["0", "0;[Red]-0"]],
);

export const blankValuesFixture = fixture(
  "Fixture!A1:C5",
  [["Date", "Alpha", "Beta"], [46023, 10, 20], [46030, null, 21], [46037, 12, ""], [46044, 13, 23]],
  [["Date", "Alpha", "Beta"], ["2026-01-01", "10", "20"], ["2026-01-08", "", "21"], ["2026-01-15", "12", ""], ["2026-01-22", "13", "23"]],
  [["General", "General", "General"], ["yyyy-mm-dd", "0", "0"], ["yyyy-mm-dd", "0", "0"], ["yyyy-mm-dd", "0", "0"], ["yyyy-mm-dd", "0", "0"]],
);

export const eightSeriesFixture = fixture(
  "Fixture!A1:I4",
  [
    ["Period", "S1", "S2", "S3", "S4", "S5", "S6", "S7", "S8"],
    [1, 1, 2, 3, 4, 5, 6, 7, 8],
    [2, 2, 3, 4, 5, 6, 7, 8, 9],
    [3, 3, 4, 5, 6, 7, 8, 9, 10],
  ],
  [
    ["Period", "S1", "S2", "S3", "S4", "S5", "S6", "S7", "S8"],
    ["P1", "1", "2", "3", "4", "5", "6", "7", "8"],
    ["P2", "2", "3", "4", "5", "6", "7", "8", "9"],
    ["P3", "3", "4", "5", "6", "7", "8", "9", "10"],
  ],
  [
    ["General", "General", "General", "General", "General", "General", "General", "General", "General"],
    ...Array.from({ length: 3 }, () => ["0", "0", "0", "0", "0", "0", "0", "0", "0"]),
  ],
);

export const titleRowFixture = fixture(
  "Fixture!A1:C5",
  [["Weekly Prices", null, null], ["Date", "Alpha", "Beta"], [46023, 10, 20], [46030, 11, 19], [46037, 12, 21]],
  [["Weekly Prices", "", ""], ["Date", "Alpha", "Beta"], ["2026-01-01", "10", "20"], ["2026-01-08", "11", "19"], ["2026-01-15", "12", "21"]],
  [["General", "General", "General"], ["General", "General", "General"], ["yyyy-mm-dd", "0", "0"], ["yyyy-mm-dd", "0", "0"], ["yyyy-mm-dd", "0", "0"]],
);

export const horizontalOrientationFixture = fixture(
  "Fixture!A1:D3",
  [["Series", "Q1", "Q2", "Q3"], ["Alpha", 10, 11, 12], ["Beta", 20, 19, 21]],
  [["Series", "Q1", "Q2", "Q3"], ["Alpha", "10", "11", "12"], ["Beta", "20", "19", "21"]],
  [["General", "General", "General", "General"], ["General", "0", "0", "0"], ["General", "0", "0", "0"]],
);

export const macLineRegressionFixture = fixture(
  "Fixture!A1:C13",
  [
    [null, "A", "B"],
    ...Array.from({ length: 12 }, (_, index) => [36892 + index * 365, index + 1, 5]),
  ],
  [
    ["", "A", "B"],
    ...Array.from({ length: 12 }, (_, index) => [
      `${2001 + index}-01-01`,
      String(index + 1),
      "5",
    ]),
  ],
  [
    ["General", "General", "General"],
    ...Array.from({ length: 12 }, () => ["yyyy-mm-dd", "0", "0"]),
  ],
);

export const invalidScatterXFixture = fixture(
  "Fixture!A1:B4",
  [["Label", "Value"], ["A", 1], ["B", 2], ["C", 3]],
  [["Label", "Value"], ["A", "1"], ["B", "2"], ["C", "3"]],
  Array.from({ length: 4 }, () => ["General", "0"]),
);
