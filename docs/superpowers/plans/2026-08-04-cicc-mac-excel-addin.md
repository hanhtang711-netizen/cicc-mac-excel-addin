# CICC Mac Excel Add-in Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Mac-compatible Office.js Excel add-in that creates ten native CICC-styled chart types from the current selection and exposes separate “Format Table” and “Zebra Stripe” commands.

**Architecture:** Keep selection parsing, chart planning, style planning, and table-format planning as pure TypeScript modules. A thin Office.js gateway reads the active selection and applies those plans to Excel; Ribbon command handlers and the advanced task pane call the same application services. Build the add-in as a three-page Vite site (`commands.html`, `taskpane.html`, and `feedback.html`) plus a generated manifest whose base URL is injected for local or production HTTPS hosting.

**Tech Stack:** TypeScript, Office.js Excel JavaScript API, Vite, Vitest, jsdom, `@types/office-js`, `office-addin-manifest`, and `sharp` for deterministic Ribbon icon generation.

## Global Constraints

- Target host: Microsoft Excel for Mac 16.111.2; Windows Excel compatibility is allowed but not the primary acceptance environment.
- Set the manifest minimum requirement to `ExcelApi 1.9`; runtime capability checks still guard exploded-pie and trendline operations.
- Use Office.js only; do not add VBA, `.xlam`, VSTO, COM, Excel-DNA, or XLL code.
- Workbook values, formulas, number formats, and chart source data must remain local to Excel and must not be sent over the network.
- Chart types are exactly: clustered column, stacked column, line, line with markers, pie, clustered bar, XY scatter with linear trendline, exploded pie, 100% stacked column, and stacked line.
- Do not implement organization charts or formatting of existing charts.
- CICC series colors are exactly `#640000`, `#8A2626`, `#3D889A`, `#646C86`, `#BE995D`, and `#DD965D`, in that order and cycling after six series.
- Chart text targets 8pt. Font-family application is best effort and must never block chart creation.
- “Format Table” creates a `#8A2626` header with white text and a white body; it must not apply zebra stripes.
- “Zebra Stripe” changes data-row fills only, alternating white and `#F5F5F5`, without overwriting a detected header.
- Neither table command may change values, formulas, number formats, merge state, or conditional-formatting rules.
- Direct chart buttons use defaults; the advanced task pane controls title, size, legend, row/column orientation, labels, and scatter trendline.
- Production resources must load from one fixed HTTPS base URL. Local development uses `https://localhost:3000`.
- No accounts, telemetry SDKs, upload endpoints, external data queries, model comparison, version service, Word, or PowerPoint features.

---

## Planned File Structure

```text
.
├── commands.html                       # Office command runtime page
├── taskpane.html                       # Advanced chart task pane page
├── feedback.html                       # Same-origin error/warning dialog page
├── package.json                        # Scripts and pinned dependency ranges
├── package-lock.json                   # Reproducible dependency lock
├── tsconfig.json                       # Strict TypeScript settings
├── vite.config.ts                      # Multi-page HTTPS build
├── manifest/
│   ├── manifest.template.xml           # Ribbon definition with BASE_URL token
│   └── manifest.dev.xml                # Generated localhost manifest
├── public/assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   └── icon-80.png
├── scripts/
│   ├── generate-icons.mjs              # Creates branded PNG icons
│   ├── render-manifest-lib.mjs         # Pure manifest renderer and validation
│   ├── render-manifest.mjs             # CLI manifest generator
│   └── package-release.mjs             # Creates deployable static bundle
├── src/
│   ├── commands.ts                     # Office action registration
│   ├── taskpane.ts                     # Advanced pane initialization
│   ├── taskpane.css                    # Advanced pane styles
│   ├── feedback.ts                     # Safe dialog message rendering
│   ├── feedback.css                    # Minimal dialog styles
│   ├── app/
│   │   ├── chartService.ts             # Parse → plan → create orchestration
│   │   ├── tableService.ts             # Table-plan orchestration
│   │   └── userFeedback.ts             # Error/warning presentation
│   ├── charts/
│   │   ├── chartCatalog.ts             # Ten chart definitions and action IDs
│   │   ├── chartPlanner.ts             # Pure chart plan construction
│   │   └── chartStyle.ts               # Pure CICC style and placement plan
│   ├── core/
│   │   ├── errors.ts                   # Typed error codes and user copy
│   │   ├── selectionParser.ts           # Pure selection classification
│   │   └── types.ts                     # Shared contracts
│   ├── office/
│   │   ├── capability.ts               # Requirement-set checks
│   │   └── excelGateway.ts             # All Excel.run calls
│   └── tables/
│       └── tablePlanner.ts              # Standard-table and zebra plans
├── tests/
│   ├── app/
│   ├── charts/
│   ├── core/
│   ├── fixtures/
│   ├── manifest/
│   ├── office/
│   └── tables/
└── docs/
    ├── DEPLOY.md                        # HTTPS deployment and manifest generation
    ├── INSTALL_MAC.md                   # Mac sideload, restart, and uninstall
    ├── PRIVACY.md                       # Local-only workbook-data guarantee
    └── TEST_CHECKLIST.md                # Excel 16.111.2 manual acceptance matrix
```

### Task 1: Buildable Office.js Project and Deterministic Manifest

**Files:**
- Create: `package.json`
- Create: `package-lock.json`
- Create: `tsconfig.json`
- Create: `vite.config.ts`
- Create: `commands.html`
- Create: `taskpane.html`
- Create: `manifest/manifest.template.xml`
- Create: `scripts/render-manifest-lib.mjs`
- Create: `scripts/render-manifest.mjs`
- Create: `scripts/generate-icons.mjs`
- Create: `tests/manifest/renderManifest.test.ts`
- Generate: `public/assets/icon-16.png`
- Generate: `public/assets/icon-32.png`
- Generate: `public/assets/icon-64.png`
- Generate: `public/assets/icon-80.png`
- Generate: `manifest/manifest.dev.xml`

**Interfaces:**
- Produces: `renderManifest(template: string, baseUrl: string): string`.
- Produces: Vite entry points `/commands.html` and `/taskpane.html` used by the manifest.
- Consumes: no project code.

- [ ] **Step 1: Write the failing manifest-renderer tests**

```ts
import { describe, expect, it } from "vitest";
import { renderManifest } from "../../scripts/render-manifest-lib.mjs";

describe("renderManifest", () => {
  it("replaces every BASE_URL token and keeps command action names", () => {
    const xml = renderManifest(
      "<Source>{{BASE_URL}}/commands.html</Source><Function>createColumnChart</Function>",
      "https://localhost:3000/",
    );
    expect(xml).toContain("https://localhost:3000/commands.html");
    expect(xml).toContain("createColumnChart");
    expect(xml).not.toContain("{{BASE_URL}}");
  });

  it("rejects non-HTTPS production URLs", () => {
    expect(() => renderManifest("{{BASE_URL}}", "http://example.com")).toThrow(
      "Base URL must use HTTPS",
    );
  });
});
```

- [ ] **Step 2: Add the package and TypeScript/Vite configuration, then install dependencies**

Create scripts with these exact names: `dev`, `build`, `test`, `typecheck`, `icons`, `manifest:dev`, and `manifest:validate`. Install `typescript`, `vite`, `vitest`, `jsdom`, `@types/office-js`, `@vitejs/plugin-basic-ssl`, `office-addin-manifest`, and `sharp` as development dependencies, and commit the resulting lockfile.

```bash
npm install --save-dev typescript vite vitest jsdom @types/office-js @vitejs/plugin-basic-ssl office-addin-manifest sharp
npm test -- tests/manifest/renderManifest.test.ts
```

Expected: FAIL because `render-manifest-lib.mjs` does not exist.

- [ ] **Step 3: Implement strict manifest rendering**

```js
export function renderManifest(template, rawBaseUrl) {
  const baseUrl = rawBaseUrl.replace(/\/$/, "");
  const isLocalHttps = baseUrl === "https://localhost:3000";
  if (!baseUrl.startsWith("https://") || (!isLocalHttps && baseUrl.includes("localhost"))) {
    throw new Error("Base URL must use HTTPS");
  }
  const rendered = template.replaceAll("{{BASE_URL}}", baseUrl);
  if (rendered.includes("{{BASE_URL}}")) {
    throw new Error("Manifest contains an unresolved BASE_URL token");
  }
  return rendered;
}
```

The CLI reads `manifest/manifest.template.xml`, takes base URL and output path as arguments, calls `renderManifest`, and writes UTF-8 XML. The template declares a Workbook host, ReadWriteDocument permission, `commands.html` as the function file, `taskpane.html` as the advanced-pane URL, the ten chart actions, two table actions, and 16/32/64/80px icon resources.

- [ ] **Step 4: Add minimal HTML entry points and deterministic icons**

Both HTML files must load Office.js from `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`. `commands.html` imports `/src/commands.ts`; `taskpane.html` imports `/src/taskpane.ts` and `/src/taskpane.css`.

`generate-icons.mjs` uses `sharp` to render a square `#640000` background with a white “C” monogram at 16, 32, 64, and 80px. Do not download icon assets.

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
<script type="module" src="/src/commands.ts"></script>
```

- [ ] **Step 5: Run foundation checks**

```bash
npm run icons
npm run manifest:dev
npm run manifest:validate
npm test -- tests/manifest/renderManifest.test.ts
npm run typecheck
npm run build
```

Expected: all commands exit 0; `manifest/manifest.dev.xml` references only `https://localhost:3000` URLs; Vite emits both HTML pages.

- [ ] **Step 6: Commit the foundation**

```bash
git add package.json package-lock.json tsconfig.json vite.config.ts commands.html taskpane.html manifest scripts public/assets tests/manifest
git commit -m "build: scaffold Office.js Excel add-in"
```

### Task 2: Selection Snapshot and Deterministic Parser

**Files:**
- Create: `src/core/types.ts`
- Create: `src/core/selectionParser.ts`
- Create: `src/core/errors.ts`
- Create: `tests/core/selectionParser.test.ts`

**Interfaces:**
- Produces: `parseSelection(snapshot: SelectionSnapshot, orientation?: SeriesOrientation): ParsedSelection`.
- Produces: `SelectionSnapshot`, `ParsedSelection`, `ParsedSeries`, `CellKind`, and `SeriesOrientation` types.
- Produces: `FeedbackPort` with asynchronous `showError` and `showWarnings` methods.
- Produces: `AddinError` with stable `code` and `details` fields.
- Consumes: no Office global; the parser must run in Vitest without Excel.

- [ ] **Step 1: Write parser tests for normal headers, title rows, and invalid shapes**

```ts
import { describe, expect, it } from "vitest";
import { parseSelection } from "../../src/core/selectionParser";

const snapshot = {
  worksheetName: "Data",
  address: "Data!A1:C4",
  rowIndex: 0,
  columnIndex: 0,
  rowCount: 4,
  columnCount: 3,
  values: [["Date", "Pulp", "Paper"], [1, 10, 20], [2, 11, 21], [3, 12, 22]],
  texts: [["Date", "Pulp", "Paper"], ["2026-01-01", "10", "20"], ["2026-01-08", "11", "21"], ["2026-01-15", "12", "22"]],
  numberFormats: [["General", "General", "General"], ["yyyy-mm-dd", "0", "0"], ["yyyy-mm-dd", "0", "0"], ["yyyy-mm-dd", "0", "0"]],
};

describe("parseSelection", () => {
  it("uses row one as headers and column one as categories", () => {
    const parsed = parseSelection(snapshot, "columns");
    expect(parsed.headerRows).toBe(1);
    expect(parsed.categoryColumnOffset).toBe(0);
    expect(parsed.series.map((series) => series.name)).toEqual(["Pulp", "Paper"]);
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
    expect(() => parseSelection({ ...snapshot, rowCount: 1, columnCount: 1 })).toThrow(
      "selection_too_small",
    );
  });
});
```

- [ ] **Step 2: Run the parser tests and confirm failure**

```bash
npm test -- tests/core/selectionParser.test.ts
```

Expected: FAIL because the parser and types do not exist.

- [ ] **Step 3: Define shared contracts and stable error codes**

```ts
export type SeriesOrientation = "columns" | "rows";
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
}

export interface ParsedSelection {
  worksheetName: string;
  sourceAddress: string;
  title?: string;
  headerRows: number;
  categoryColumnOffset: number;
  categoryAddress: string;
  orientation: SeriesOrientation;
  series: ParsedSeries[];
  numberFormats: string[][];
}

export interface FeedbackPort {
  showError(error: unknown): Promise<void>;
  showWarnings(codes: string[]): Promise<void>;
}
```

`AddinError` codes must include `invalid_selection`, `selection_too_small`, `unsupported_layout`, `pie_requires_one_series`, `scatter_requires_numeric_x`, `unsupported_api`, `protected_sheet`, and `excel_runtime_error`.

- [ ] **Step 4: Implement parsing and address construction**

Implement these exact rules: reject non-rectangular matrices; detect a title only when the first row has exactly one non-empty text cell; detect a header when the candidate header row is majority text and the following data cells are majority numeric/date; keep blanks as blanks; quote worksheet names with doubled apostrophes; and produce absolute A1 addresses for categories and every series.

```ts
export function quoteSheetName(name: string): string {
  return `'${name.replaceAll("'", "''")}'`;
}
```

- [ ] **Step 5: Run tests and typecheck**

```bash
npm test -- tests/core/selectionParser.test.ts
npm run typecheck
```

Expected: PASS.

- [ ] **Step 6: Commit the parser**

```bash
git add src/core tests/core
git commit -m "feat: parse Excel selections for charting"
```

### Task 3: Ten-Chart Catalog and Chart Planner

**Files:**
- Create: `src/charts/chartCatalog.ts`
- Create: `src/charts/chartPlanner.ts`
- Create: `tests/charts/chartPlanner.test.ts`
- Modify: `src/core/types.ts`

**Interfaces:**
- Consumes: `ParsedSelection` and `SeriesOrientation` from Task 2.
- Produces: `ChartKind`, `ChartOptions`, `ChartPlan`, `ChartSeriesPlan`, and `buildChartPlan(parsed, kind, options)`.
- Produces: `CHART_CATALOG`, the single source of truth for action IDs, labels, and Excel chart type keys.

- [ ] **Step 1: Write chart mapping and validation tests**

```ts
import { describe, expect, it } from "vitest";
import { buildChartPlan } from "../../src/charts/chartPlanner";

describe("buildChartPlan", () => {
  it.each([
    ["column", "columnClustered"],
    ["columnStacked", "columnStacked"],
    ["line", "line"],
    ["lineMarkers", "lineMarkers"],
    ["pie", "pie"],
    ["bar", "barClustered"],
    ["scatterTrend", "xyscatter"],
    ["pieExploded", "pieExploded"],
    ["columnStacked100", "columnStacked100"],
    ["lineStacked", "lineStacked"],
  ] as const)("maps %s to %s", (kind, excelType) => {
    expect(buildChartPlan(parsedSelection, kind, {}).excelType).toBe(excelType);
  });

  it("rejects a pie selection with two value series", () => {
    expect(() => buildChartPlan(parsedSelection, "pie", {})).toThrow(
      "pie_requires_one_series",
    );
  });
});
```

Define `parsedSelection` in the test with two named series, an absolute category address, and absolute series addresses.

- [ ] **Step 2: Run the planner test and confirm failure**

```bash
npm test -- tests/charts/chartPlanner.test.ts
```

Expected: FAIL because the chart planner does not exist.

- [ ] **Step 3: Implement the catalog and chart contracts**

```ts
export const CHART_CATALOG = {
  column: { actionId: "createColumnChart", label: "普通柱形图", excelType: "columnClustered" },
  columnStacked: { actionId: "createStackedColumnChart", label: "堆积柱形图", excelType: "columnStacked" },
  line: { actionId: "createLineChart", label: "普通折线图", excelType: "line" },
  lineMarkers: { actionId: "createMarkedLineChart", label: "带标记折线图", excelType: "lineMarkers" },
  pie: { actionId: "createPieChart", label: "饼图", excelType: "pie" },
  bar: { actionId: "createBarChart", label: "条形图", excelType: "barClustered" },
  scatterTrend: { actionId: "createScatterTrendChart", label: "散点图＋趋势线", excelType: "xyscatter" },
  pieExploded: { actionId: "createExplodedPieChart", label: "分离饼图", excelType: "pieExploded" },
  columnStacked100: { actionId: "createPercentageStackedChart", label: "百分比堆积图", excelType: "columnStacked100" },
  lineStacked: { actionId: "createStackedLineChart", label: "堆积折线图", excelType: "lineStacked" },
} as const;
```

`ChartOptions` includes optional `title`, `sizePreset`, `widthCm`, `heightCm`, `legendPosition`, `orientation`, `showDataLabels`, and `addTrendline` fields. Direct commands pass an empty object and receive documented defaults.

- [ ] **Step 4: Implement chart-specific validation and plans**

For pie types, require exactly one series. For scatter, require a numeric/date X range and at least one numeric Y series. Set `addLinearTrendline` true by default only for `scatterTrend`. Do not rewrite percentage-stacked source values. Preserve source series order.

```ts
if ((kind === "pie" || kind === "pieExploded") && parsed.series.length !== 1) {
  throw new AddinError("pie_requires_one_series");
}
const addLinearTrendline = kind === "scatterTrend" && options.addTrendline !== false;
```

- [ ] **Step 5: Run chart-planner tests**

```bash
npm test -- tests/charts/chartPlanner.test.ts
npm run typecheck
```

Expected: PASS for all ten mappings and validation cases.

- [ ] **Step 6: Commit the chart planner**

```bash
git add src/charts src/core/types.ts tests/charts/chartPlanner.test.ts
git commit -m "feat: plan ten CICC chart types"
```

### Task 4: CICC Chart Style, Axis Format, and Placement Plans

**Files:**
- Create: `src/charts/chartStyle.ts`
- Create: `tests/charts/chartStyle.test.ts`
- Modify: `src/core/types.ts`

**Interfaces:**
- Consumes: `ChartPlan`, `ChartOptions`, parsed number formats, and selection bounds.
- Produces: `buildChartStylePlan(input): ChartStylePlan`.
- Produces: `CICC_SERIES_COLORS`, `SIZE_PRESETS_CM`, and `cmToPoints`.

- [ ] **Step 1: Write exact style tests**

```ts
import { describe, expect, it } from "vitest";
import {
  buildChartStylePlan,
  CICC_SERIES_COLORS,
  cmToPoints,
} from "../../src/charts/chartStyle";

describe("CICC chart style", () => {
  it("cycles the six exact series colors", () => {
    const style = buildChartStylePlan({ seriesCount: 8, options: {}, sourceFormat: "0.0%" });
    expect(style.seriesColors).toEqual([
      ...CICC_SERIES_COLORS,
      "#640000",
      "#8A2626",
    ]);
    expect(style.warnings).toContain("series_palette_reused");
  });

  it("uses the medium preset and percentage precision", () => {
    const style = buildChartStylePlan({ seriesCount: 2, options: {}, sourceFormat: "0.0%" });
    expect(style.widthPoints).toBeCloseTo(cmToPoints(11.5), 4);
    expect(style.heightPoints).toBeCloseTo(cmToPoints(6.7), 4);
    expect(style.valueAxisNumberFormat).toBe("0.0%");
    expect(style.legendPosition).toBe("bottom");
  });
});
```

- [ ] **Step 2: Run the style tests and confirm failure**

```bash
npm test -- tests/charts/chartStyle.test.ts
```

Expected: FAIL because `chartStyle.ts` does not exist.

- [ ] **Step 3: Implement exact style constants and size validation**

```ts
export const CICC_SERIES_COLORS = [
  "#640000",
  "#8A2626",
  "#3D889A",
  "#646C86",
  "#BE995D",
  "#DD965D",
] as const;

export const SIZE_PRESETS_CM = {
  small: { width: 9, height: 5.5 },
  medium: { width: 11.5, height: 6.7 },
  large: { width: 15, height: 8.7 },
} as const;

export const cmToPoints = (centimeters: number): number => centimeters * 28.3464567;
```

Reject custom dimensions that are non-finite or not greater than zero. Default to medium size, bottom legend, white chart/plot background, no decorative outer border, 8pt text, and light-gray major gridlines.

- [ ] **Step 4: Implement format and placement decisions**

Use `yyyy-mm-dd` when the category format is date-like but cannot be reused. Preserve `0%`, `0.0%`, or `0.00%` based on the first value-series format. Put the chart to the right of the selection with an 18pt gutter; only place it below when adding the chart width would cross Excel's final column boundary.

```ts
const valueAxisNumberFormat = /^0(?:\.0{1,2})?%$/.test(sourceFormat)
  ? sourceFormat
  : undefined;
const placeBelow = selectionColumn + selectionColumnCount + estimatedChartColumns > 16384;
```

- [ ] **Step 5: Run style tests and typecheck**

```bash
npm test -- tests/charts/chartStyle.test.ts
npm run typecheck
```

Expected: PASS.

- [ ] **Step 6: Commit the chart style planner**

```bash
git add src/charts/chartStyle.ts src/core/types.ts tests/charts/chartStyle.test.ts
git commit -m "feat: define CICC chart styling"
```

### Task 5: Office.js Gateway for Native Chart Creation

**Files:**
- Create: `src/office/excelGateway.ts`
- Create: `src/app/chartService.ts`
- Create: `tests/office/excelGateway.test.ts`
- Create: `tests/app/chartService.test.ts`

**Interfaces:**
- Consumes: `parseSelection`, `buildChartPlan`, and `buildChartStylePlan`.
- Produces: `ExcelGateway.readSelection(): Promise<SelectionSnapshot>`.
- Produces: `ExcelGateway.createChart(plan: ChartPlan, style: ChartStylePlan): Promise<void>`.
- Produces: `ChartService.create(kind: ChartKind, options?: ChartOptions): Promise<ServiceResult>`.

- [ ] **Step 1: Write orchestration and rollback tests**

```ts
import { describe, expect, it, vi } from "vitest";
import { ChartService } from "../../src/app/chartService";

describe("ChartService", () => {
  it("reads, parses, plans, and creates one chart", async () => {
    const gateway = {
      readSelection: vi.fn().mockResolvedValue(selectionSnapshot),
      createChart: vi.fn().mockResolvedValue(undefined),
    };
    const service = new ChartService(gateway);
    const result = await service.create("column");
    expect(result.ok).toBe(true);
    expect(gateway.createChart).toHaveBeenCalledTimes(1);
  });
});
```

In `excelGateway.test.ts`, use a fake chart object whose styling call throws, and assert that `chart.delete()` is called before the error is rethrown.

- [ ] **Step 2: Run gateway tests and confirm failure**

```bash
npm test -- tests/app/chartService.test.ts tests/office/excelGateway.test.ts
```

Expected: FAIL because the gateway and service do not exist.

- [ ] **Step 3: Implement selection reads in one Excel.run transaction**

Load `areaCount` from `context.workbook.getSelectedRanges()` first and throw `invalid_selection` unless it equals one. Then load `address`, `rowIndex`, `columnIndex`, `rowCount`, `columnCount`, `values`, `text`, and `numberFormat` from `context.workbook.getSelectedRange()`, plus the active worksheet name. Return plain arrays and strings after `context.sync()`; never expose Office proxy objects outside the callback.

```ts
async readSelection(): Promise<SelectionSnapshot> {
  return Excel.run(async (context) => {
    const areas = context.workbook.getSelectedRanges();
    areas.load("areaCount");
    await context.sync();
    if (areas.areaCount !== 1) throw new AddinError("invalid_selection");
    const range = context.workbook.getSelectedRange();
    range.load("address,rowIndex,columnIndex,rowCount,columnCount,values,text,numberFormat");
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();
    return snapshotFromRange(sheet.name, range);
  });
}
```

Define `snapshotFromRange(sheetName, range)` as a private helper in `excelGateway.ts`; it copies only the loaded scalar properties and 2D arrays into a new `SelectionSnapshot` object.

- [ ] **Step 4: Implement native chart creation**

For non-scatter charts, call `worksheet.charts.add` with the mapped `Excel.ChartType`, selected range, and `Excel.ChartSeriesBy.rows` or `.columns`. For scatter charts, create an empty XY chart and add each Y series with explicit category and value formulas. Apply series colors in source order, color pie points individually, add linear trendlines to scatter series, and set chart title, legend, axes, labels, width, height, left, and top from the plans.

```ts
const chartType = Excel.ChartType[plan.excelType as keyof typeof Excel.ChartType];
const seriesBy = plan.orientation === "rows"
  ? Excel.ChartSeriesBy.rows
  : Excel.ChartSeriesBy.columns;
```

Keep the newly created chart proxy in scope. If any later API call or `context.sync()` fails, call `chart.delete()`, sync once, and throw a typed `excel_runtime_error`.

- [ ] **Step 5: Implement ChartService**

`ChartService.create` runs the pure parser and planners before any write, calls the gateway only after validation succeeds, and returns `{ ok: true, warnings }` or throws `AddinError`. It must not catch and discard typed validation errors.

```ts
async create(kind: ChartKind, options: ChartOptions = {}): Promise<ServiceResult> {
  const snapshot = await this.gateway.readSelection();
  const parsed = parseSelection(snapshot, options.orientation);
  const plan = buildChartPlan(parsed, kind, options);
  const style = buildChartStylePlan({
    seriesCount: plan.series.length,
    options,
    sourceFormat: parsed.numberFormats[parsed.headerRows]?.[1] ?? "General",
    categoryFormat: parsed.numberFormats[parsed.headerRows]?.[0] ?? "General",
    selectionColumn: snapshot.columnIndex,
    selectionColumnCount: snapshot.columnCount,
  });
  await this.gateway.createChart(plan, style);
  return { ok: true, warnings: style.warnings };
}
```

- [ ] **Step 6: Run chart-service checks**

```bash
npm test -- tests/app/chartService.test.ts tests/office/excelGateway.test.ts
npm run typecheck
```

Expected: PASS, including rollback coverage.

- [ ] **Step 7: Commit native chart creation**

```bash
git add src/office/excelGateway.ts src/app/chartService.ts tests/office tests/app/chartService.test.ts
git commit -m "feat: create native Excel charts"
```

### Task 6: Independent Standard-Table and Zebra-Stripe Commands

**Files:**
- Create: `src/tables/tablePlanner.ts`
- Create: `src/app/tableService.ts`
- Create: `tests/tables/tablePlanner.test.ts`
- Create: `tests/app/tableService.test.ts`
- Modify: `src/office/excelGateway.ts`

**Interfaces:**
- Consumes: `SelectionSnapshot` from Task 2.
- Produces: `buildStandardTablePlan(snapshot): TableFormatPlan`.
- Produces: `buildZebraPlan(snapshot): TableFormatPlan`.
- Produces: `ExcelGateway.applyTablePlan(plan): Promise<void>`.
- Produces: `TableService.formatStandard()` and `TableService.applyZebra()`.

- [ ] **Step 1: Write tests proving the two commands remain separate**

```ts
import { describe, expect, it } from "vitest";
import { buildStandardTablePlan, buildZebraPlan } from "../../src/tables/tablePlanner";

describe("table plans", () => {
  it("formats a red header and white body without zebra actions", () => {
    const plan = buildStandardTablePlan(selectionSnapshot);
    expect(plan.header.fill).toBe("#8A2626");
    expect(plan.header.fontColor).toBe("#FFFFFF");
    expect(plan.body.fill).toBe("#FFFFFF");
    expect(plan.rowFills).toEqual([]);
  });

  it("zebra plan changes fills only and skips a detected header", () => {
    const plan = buildZebraPlan(selectionSnapshot);
    expect(plan.preserve).toEqual([
      "values",
      "formulas",
      "numberFormats",
      "merges",
      "conditionalFormats",
      "fonts",
      "borders",
      "alignment",
    ]);
    expect(plan.rowFills[0]).toEqual({ rowOffset: 1, fill: "#FFFFFF" });
    expect(plan.rowFills[1]).toEqual({ rowOffset: 2, fill: "#F5F5F5" });
  });
});
```

- [ ] **Step 2: Run table tests and confirm failure**

```bash
npm test -- tests/tables/tablePlanner.test.ts tests/app/tableService.test.ts
```

Expected: FAIL because the table modules do not exist.

- [ ] **Step 3: Implement pure table plans**

Standard format overwrites direct font, fill, border, and alignment properties only. It sets the first row to red/white/bold/centered at 10.5pt, the body to white at 10.5pt, text columns left, numeric columns right, date columns center, and thin light-gray structural borders. It never includes a zebra row-fill list.

Zebra detection skips row one when it is majority text and the following rows contain data. Its plan contains row fill actions only: first data row white, second data row `#F5F5F5`, then alternating.

```ts
export function buildStandardTablePlan(snapshot: SelectionSnapshot): TableFormatPlan {
  return {
    address: snapshot.address,
    header: { fill: "#8A2626", fontColor: "#FFFFFF", bold: true, fontSize: 10.5 },
    body: { fill: "#FFFFFF", fontColor: "#000000", bold: false, fontSize: 10.5 },
    rowFills: [],
    preserve: ["values", "formulas", "numberFormats", "merges", "conditionalFormats"],
  };
}
```

- [ ] **Step 4: Apply plans without writing workbook content**

In `ExcelGateway.applyTablePlan`, get the current range, apply direct formats in bulk, then cap each autofitted column at 180 points and each autofitted row at 45 points. Never assign `values`, `formulas`, `numberFormat`, merge APIs, or conditional-format collections. Zebra execution must assign only `RangeFormat.fill.color` on data-row subranges.

```ts
async applyTablePlan(plan: TableFormatPlan): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(plan.address);
    applyDirectFormats(range, plan);
    await context.sync();
  });
}
```

- [ ] **Step 5: Implement TableService validation**

Reject empty, discontiguous, whole-row, whole-column, or whole-sheet selections before applying a plan. Call `readSelection` once and `applyTablePlan` once per user action.

```ts
async formatStandard(): Promise<ServiceResult> {
  const snapshot = await this.gateway.readSelection();
  validateTableSelection(snapshot);
  await this.gateway.applyTablePlan(buildStandardTablePlan(snapshot));
  return { ok: true, warnings: [] };
}
```

- [ ] **Step 6: Run table checks**

```bash
npm test -- tests/tables/tablePlanner.test.ts tests/app/tableService.test.ts
npm run typecheck
```

Expected: PASS; the zebra test confirms no font, border, alignment, value, formula, or number-format mutation.

- [ ] **Step 7: Commit table commands**

```bash
git add src/tables src/app/tableService.ts src/office/excelGateway.ts tests/tables tests/app/tableService.test.ts
git commit -m "feat: add independent CICC table styles"
```

### Task 7: Ribbon Command Registration and Action Wiring

**Files:**
- Create: `src/app/commandHandlers.ts`
- Create: `src/app/userFeedback.ts`
- Create: `feedback.html`
- Create: `src/feedback.ts`
- Create: `src/feedback.css`
- Create: `tests/app/commandHandlers.test.ts`
- Modify: `src/commands.ts`
- Modify: `vite.config.ts`
- Modify: `manifest/manifest.template.xml`
- Modify: `tests/manifest/renderManifest.test.ts`

**Interfaces:**
- Consumes: `CHART_CATALOG`, `ChartService`, `TableService`, and `FeedbackPort` from Task 2.
- Produces: `registerCommandHandlers(dependencies): void`.
- Produces: `createCommandDependencies()` as the single runtime composition root.
- Produces: thirteen manifest actions: ten chart actions, `formatCiccTable`, `applyZebraStripe`, and `openAdvancedChartPane`.

- [ ] **Step 1: Write action-registration and completion tests**

```ts
import { describe, expect, it, vi } from "vitest";
import { createCommandHandlers } from "../../src/app/commandHandlers";

describe("command handlers", () => {
  it("always completes the Office event", async () => {
    const completed = vi.fn();
    const services = {
      charts: { create: vi.fn().mockRejectedValue(new Error("failed")) },
      tables: { formatStandard: vi.fn(), applyZebra: vi.fn() },
      feedback: { showError: vi.fn(), showWarnings: vi.fn() },
    };
    const handlers = createCommandHandlers(services);
    await handlers.createColumnChart({ completed });
    expect(completed).toHaveBeenCalledTimes(1);
    expect(services.feedback.showError).toHaveBeenCalledTimes(1);
  });
});
```

- [ ] **Step 2: Run handler tests and confirm failure**

```bash
npm test -- tests/app/commandHandlers.test.ts tests/manifest/renderManifest.test.ts
```

Expected: FAIL because the handlers do not exist and the full manifest actions are not yet defined.

- [ ] **Step 3: Implement one shared command wrapper**

```ts
async function runCommand(
  event: Office.AddinCommands.Event,
  operation: () => Promise<{ warnings?: string[] }>,
  feedback: FeedbackPort,
): Promise<void> {
  try {
    const result = await operation();
    await feedback.showWarnings(result.warnings ?? []);
  } catch (error) {
    await feedback.showError(error);
  } finally {
    event.completed();
  }
}
```

Create ten chart handlers from `CHART_CATALOG` plus explicit standard-table and zebra handlers. No handler may duplicate parsing or Office.js formatting logic.

- [ ] **Step 4: Register Office actions on readiness**

`src/commands.ts` constructs the gateway, services, and `DialogFeedback`, calls `Office.actions.associate(actionId, handler)` for all twelve execute-function actions, and performs no workbook operation during startup. `openAdvancedChartPane` remains a manifest `ShowTaskpane` action and is not associated in JavaScript.

`DialogFeedback` implements `FeedbackPort` with a same-origin `feedback.html` dialog. The dialog obtains a message code from query parameters, maps it to hardcoded local copy, assigns it through `textContent`, and never renders workbook-derived HTML. Add `feedback.html` as Vite's third build entry.

```ts
Office.onReady(() => {
  const handlers = createCommandHandlers(createCommandDependencies());
  for (const [actionId, handler] of Object.entries(handlers)) {
    Office.actions.associate(actionId, handler);
  }
});
```

- [ ] **Step 5: Complete Ribbon XML and verify action-name parity**

Use a `中金工具` tab with a `生成图表` group and a `表格` group. Put the ten chart buttons in a menu, add a separate “高级生成” task-pane button, and add “格式化表格” plus “斑马纹” execute-function buttons. Extend the manifest test to compare all execute-function names with `CHART_CATALOG` plus the two table action IDs.

```xml
<Control xsi:type="Button" id="FormatCiccTable">
  <Label resid="FormatCiccTable.Label" />
  <Action xsi:type="ExecuteFunction">
    <FunctionName>formatCiccTable</FunctionName>
  </Action>
</Control>
```

- [ ] **Step 6: Run command and manifest checks**

```bash
npm run manifest:dev
npm run manifest:validate
npm test -- tests/app/commandHandlers.test.ts tests/manifest/renderManifest.test.ts
npm run typecheck
```

Expected: PASS; every execute-function action has exactly one associated handler.

- [ ] **Step 7: Commit Ribbon wiring**

```bash
git add src/commands.ts src/app/commandHandlers.ts src/app/userFeedback.ts feedback.html src/feedback.ts src/feedback.css vite.config.ts manifest/manifest.template.xml manifest/manifest.dev.xml tests/app/commandHandlers.test.ts tests/manifest/renderManifest.test.ts
git commit -m "feat: wire CICC Ribbon commands"
```

### Task 8: Advanced Chart Task Pane

**Files:**
- Create: `src/app/advancedOptions.ts`
- Create: `tests/app/advancedOptions.test.ts`
- Create: `tests/app/taskpane.test.ts`
- Modify: `taskpane.html`
- Modify: `src/taskpane.ts`
- Modify: `src/taskpane.css`

**Interfaces:**
- Consumes: `CHART_CATALOG`, `ChartOptions`, `ChartService`, and `FeedbackPort`.
- Produces: `readAdvancedOptions(form: HTMLFormElement): { kind: ChartKind; options: ChartOptions }`.
- Produces: a task pane that calls the same `ChartService.create` used by Ribbon commands.

- [ ] **Step 1: Write DOM tests for defaults and custom size validation**

```ts
import { beforeEach, describe, expect, it } from "vitest";
import { readAdvancedOptions } from "../../src/app/advancedOptions";

describe("advanced chart options", () => {
  beforeEach(() => {
    document.body.innerHTML = `
      <form id="advanced-chart-form">
        <select name="kind"><option value="column" selected>普通柱形图</option></select>
        <select name="sizePreset"><option value="medium" selected>中</option></select>
        <select name="legendPosition"><option value="bottom" selected>下</option></select>
        <select name="orientation"><option value="columns" selected>按列</option></select>
        <input name="title" value="Weekly prices">
        <input name="showDataLabels" type="checkbox">
        <input name="addTrendline" type="checkbox" checked>
      </form>`;
  });

  it("returns approved defaults", () => {
    const form = document.querySelector("form") as HTMLFormElement;
    expect(readAdvancedOptions(form)).toEqual({
      kind: "column",
      options: {
        title: "Weekly prices",
        sizePreset: "medium",
        legendPosition: "bottom",
        orientation: "columns",
        showDataLabels: false,
        addTrendline: true,
      },
    });
  });
});
```

- [ ] **Step 2: Run task-pane tests and confirm failure**

```bash
npm test -- tests/app/advancedOptions.test.ts tests/app/taskpane.test.ts
```

Expected: FAIL because advanced-option parsing and the pane are not implemented.

- [ ] **Step 3: Build the accessible form**

The form includes chart type, optional title, small/medium/large/custom size, custom width and height in centimeters, legend position, series orientation, data-label checkbox, trendline checkbox, Generate button, and an `aria-live="polite"` status region. Use plain TypeScript and CSS; do not add React or another UI framework.

```html
<form id="advanced-chart-form">
  <label>图表类型<select name="kind" required></select></label>
  <label>标题<input name="title" type="text"></label>
  <label>尺寸<select name="sizePreset" required></select></label>
  <button type="submit">生成图表</button>
</form>
<p id="status" aria-live="polite"></p>
```

- [ ] **Step 4: Implement option parsing and conditional fields**

Show custom width/height only when `sizePreset` is `custom`. Show the trendline option only for `scatterTrend`. Reject non-positive custom dimensions before calling the service. Trim empty titles to `undefined`.

```ts
const titleInput = form.elements.namedItem("title") as HTMLInputElement;
const sizeInput = form.elements.namedItem("sizePreset") as HTMLSelectElement;
const legendInput = form.elements.namedItem("legendPosition") as HTMLSelectElement;
const orientationInput = form.elements.namedItem("orientation") as HTMLSelectElement;
const widthInput = form.elements.namedItem("widthCm") as HTMLInputElement;
const heightInput = form.elements.namedItem("heightCm") as HTMLInputElement;
const title = titleInput.value.trim();
const sizePreset = sizeInput.value as ChartOptions["sizePreset"];
const legendPosition = legendInput.value as ChartOptions["legendPosition"];
const orientation = orientationInput.value as ChartOptions["orientation"];
const widthCm = Number(widthInput.value);
const heightCm = Number(heightInput.value);
const options: ChartOptions = { title: title || undefined, sizePreset, legendPosition, orientation };
if (sizePreset === "custom" && (!(widthCm > 0) || !(heightCm > 0))) {
  throw new AddinError("unsupported_layout", { reason: "invalid_custom_size" });
}
```

- [ ] **Step 5: Connect the pane to Office and ChartService**

On `Office.onReady`, enable the form only when the host is Excel. On submit, disable Generate, show “正在生成图表…”, call `ChartService.create`, show success or warning copy, and re-enable the button in `finally`. Do not transmit selection data through `fetch`, XMLHttpRequest, beacons, or form actions.

```ts
form.addEventListener("submit", async (event) => {
  event.preventDefault();
  submitButton.disabled = true;
  status.textContent = "正在生成图表…";
  try {
    const request = readAdvancedOptions(form);
    await chartService.create(request.kind, request.options);
    status.textContent = "图表已生成。";
  } finally {
    submitButton.disabled = false;
  }
});
```

- [ ] **Step 6: Run pane checks**

```bash
npm test -- tests/app/advancedOptions.test.ts tests/app/taskpane.test.ts
npm run typecheck
npm run build
```

Expected: PASS; built task-pane code contains no network call for workbook data.

- [ ] **Step 7: Commit the advanced pane**

```bash
git add taskpane.html src/taskpane.ts src/taskpane.css src/app/advancedOptions.ts tests/app/advancedOptions.test.ts tests/app/taskpane.test.ts
git commit -m "feat: add advanced chart task pane"
```

### Task 9: Capabilities, User-Facing Errors, and Non-Blocking Warnings

**Files:**
- Create: `src/office/capability.ts`
- Create: `tests/office/capability.test.ts`
- Create: `tests/app/userFeedback.test.ts`
- Modify: `src/app/userFeedback.ts`
- Modify: `src/feedback.ts`
- Modify: `src/core/errors.ts`
- Modify: `src/app/chartService.ts`
- Modify: `src/app/tableService.ts`

**Interfaces:**
- Produces: `getCapabilities(requirements): AddinCapabilities`.
- Produces: `DialogFeedback.showError(error): Promise<void>` and `showWarnings(codes): Promise<void>`.
- Consumes: error codes from Task 2 and warning codes from Task 4.

- [ ] **Step 1: Write capability and message-copy tests**

```ts
import { describe, expect, it } from "vitest";
import { messageForError } from "../../src/app/userFeedback";

describe("user feedback", () => {
  it("gives a corrective message for multi-series pie data", () => {
    expect(messageForError({ code: "pie_requires_one_series" })).toBe(
      "饼图只能使用一个数值系列。请缩小选区，或在高级生成中指定一个系列。",
    );
  });

  it("explains unsupported Excel APIs without faking a chart", () => {
    expect(messageForError({ code: "unsupported_api" })).toContain("升级 Excel");
  });
});
```

Add this capability assertion in `capability.test.ts`:

```ts
it("requires ExcelApi 1.9 for all target capabilities", () => {
  const requirements = { isSetSupported: vi.fn().mockReturnValue(true) };
  expect(getCapabilities(requirements)).toEqual({ base: true, explodedPie: true, trendlines: true });
  expect(requirements.isSetSupported).toHaveBeenCalledWith("ExcelApi", "1.9");
});
```

- [ ] **Step 2: Run feedback tests and confirm failure**

```bash
npm test -- tests/office/capability.test.ts tests/app/userFeedback.test.ts
```

Expected: FAIL because capability and feedback modules do not exist.

- [ ] **Step 3: Implement startup capability checks**

Require `ExcelApi 1.9` for the add-in. Check `Office.context.requirements.isSetSupported("ExcelApi", "1.9")` at startup and use explicit capability booleans for exploded-pie and trendline operations before execution. Return stable booleans rather than leaking `Office.context.requirements` throughout the codebase. The target Excel 16.111.2 must pass every capability.

```ts
export interface RequirementChecker {
  isSetSupported(name: string, minimumVersion: string): boolean;
}

export function getCapabilities(requirements: RequirementChecker): AddinCapabilities {
  const excelApi19 = requirements.isSetSupported("ExcelApi", "1.9");
  return { base: excelApi19, explodedPie: excelApi19, trendlines: excelApi19 };
}
```

- [ ] **Step 4: Implement Chinese error and warning copy**

Map every `AddinError` code to one message with a corrective action. Map `series_palette_reused` to a non-blocking warning explaining that colors repeat after six series. Preserve diagnostic details for console logging, but never include cell values, formulas, or workbook names in logs.

```ts
const ERROR_MESSAGES: Record<AddinErrorCode, string> = {
  invalid_selection: "请选择一个连续的矩形区域后重试。",
  selection_too_small: "图表数据至少需要两行两列。",
  unsupported_layout: "当前数据排列无法识别，请在高级生成中指定系列方向。",
  pie_requires_one_series: "饼图只能使用一个数值系列。请缩小选区，或在高级生成中指定一个系列。",
  scatter_requires_numeric_x: "散点图第一列必须是数值或日期型 X 轴。",
  unsupported_api: "当前 Excel 版本不支持此图表类型，请升级 Excel。",
  protected_sheet: "当前工作表受保护，无法写入图表或格式。",
  excel_runtime_error: "Excel 未能完成操作，请检查选区后重试。",
};
```

- [ ] **Step 5: Present command feedback safely**

Advanced-pane actions use the pane's inline status region. Execute-function commands show dialogs only for errors and non-blocking warnings, using an HTTPS page from the same add-in origin. Do not show a success dialog after every direct command.

```ts
const dialogUrl = new URL("/feedback.html", this.origin);
dialogUrl.searchParams.set("code", messageCode);
Office.context.ui.displayDialogAsync(dialogUrl.toString(), { height: 24, width: 32 });
```

- [ ] **Step 6: Run capability and feedback checks**

```bash
npm test -- tests/office/capability.test.ts tests/app/userFeedback.test.ts
npm run typecheck
```

Expected: PASS; unsupported exploded-pie execution fails before creating any chart.

- [ ] **Step 7: Commit capability and feedback behavior**

```bash
git add src/office/capability.ts src/app/userFeedback.ts src/feedback.ts src/core/errors.ts src/app/chartService.ts src/app/tableService.ts tests/office/capability.test.ts tests/app/userFeedback.test.ts
git commit -m "feat: add Excel capability and error handling"
```

### Task 10: End-to-End Fixtures, Release Bundle, Documentation, and Mac Acceptance

**Files:**
- Create: `tests/fixtures/selectionCases.ts`
- Create: `tests/app/workflows.test.ts`
- Create: `scripts/package-release.mjs`
- Create: `tests/manifest/packageRelease.test.ts`
- Create: `docs/DEPLOY.md`
- Create: `docs/INSTALL_MAC.md`
- Create: `docs/PRIVACY.md`
- Create: `docs/TEST_CHECKLIST.md`
- Modify: `package.json`

**Interfaces:**
- Consumes: all production modules and generated Vite assets.
- Produces: `npm run release -- --base-url "$ADDIN_BASE_URL" --out release/`.
- Produces: a release directory containing static site files plus a production manifest.

- [ ] **Step 1: Add representative, sanitized selection fixtures**

Create fixture shapes corresponding to the reference workbook's representative regions without committing the 13MB source workbook: date plus one series, date plus multiple series, text categories, percentages, negative values, blank values, eight series, a title row, horizontal orientation, and invalid scatter X values. Keep only short synthetic labels and numbers in Git.

```ts
export const dateSingleSeriesFixture: SelectionSnapshot = {
  worksheetName: "Fixture",
  address: "Fixture!A1:B4",
  rowIndex: 0,
  columnIndex: 0,
  rowCount: 4,
  columnCount: 2,
  values: [["Date", "Price"], [1, 100], [2, 102], [3, 101]],
  texts: [["Date", "Price"], ["2026-01-01", "100"], ["2026-01-08", "102"], ["2026-01-15", "101"]],
  numberFormats: [["General", "General"], ["yyyy-mm-dd", "0"], ["yyyy-mm-dd", "0"], ["yyyy-mm-dd", "0"]],
};
```

- [ ] **Step 2: Write end-to-end service tests**

```ts
import { describe, expect, it, vi } from "vitest";
import { ChartService } from "../../src/app/chartService";
import { TableService } from "../../src/app/tableService";
import type { SelectionSnapshot } from "../../src/core/types";

const createFixtureGateway = (snapshot: SelectionSnapshot) => ({
  readSelection: vi.fn().mockResolvedValue(snapshot),
  createChart: vi.fn().mockResolvedValue(undefined),
  applyTablePlan: vi.fn().mockResolvedValue(undefined),
});

describe("chart workflows", () => {
  it.each([
    "column",
    "columnStacked",
    "line",
    "lineMarkers",
    "bar",
    "scatterTrend",
    "columnStacked100",
    "lineStacked",
  ] as const)("creates %s from a valid multi-series fixture", async (kind) => {
    const gateway = createFixtureGateway(dateMultiSeriesFixture);
    await new ChartService(gateway).create(kind);
    expect(gateway.createChart).toHaveBeenCalledTimes(1);
  });

  it.each(["pie", "pieExploded"] as const)("creates %s from one series", async (kind) => {
    const gateway = createFixtureGateway(dateSingleSeriesFixture);
    await new ChartService(gateway).create(kind);
    expect(gateway.createChart).toHaveBeenCalledTimes(1);
  });

  it("keeps standard-table and zebra actions separate", async () => {
    const gateway = createFixtureGateway(dateMultiSeriesFixture);
    const tables = new TableService(gateway);
    await tables.formatStandard();
    await tables.applyZebra();
    expect(gateway.applyTablePlan).toHaveBeenCalledTimes(2);
    expect(gateway.applyTablePlan.mock.calls[0][0].rowFills).toEqual([]);
    expect(gateway.applyTablePlan.mock.calls[1][0].rowFills.length).toBeGreaterThan(0);
  });
});
```

- [ ] **Step 3: Run workflow tests**

```bash
npm test -- tests/app/workflows.test.ts
```

Expected: PASS for all ten chart types and both table-command combinations.

- [ ] **Step 4: Write the failing release-packager tests**

```ts
import { describe, expect, it } from "vitest";
import { validateProductionBaseUrl } from "../../scripts/package-release.mjs";

describe("release packager", () => {
  it("accepts a fixed HTTPS origin", () => {
    expect(validateProductionBaseUrl("https://excel-addon.example.com")).toBe(
      "https://excel-addon.example.com",
    );
  });

  it("rejects localhost and insecure origins", () => {
    expect(() => validateProductionBaseUrl("https://localhost:3000")).toThrow();
    expect(() => validateProductionBaseUrl("http://excel-addon.example.com")).toThrow();
  });
});
```

- [ ] **Step 5: Run the release test and confirm failure**

```bash
npm test -- tests/manifest/packageRelease.test.ts
```

Expected: FAIL because `package-release.mjs` does not exist.

- [ ] **Step 6: Implement a fail-closed release packager**

The packager validates an HTTPS base URL, runs or verifies the Vite build, renders the production manifest, copies `dist/` into the requested output directory, and writes the manifest beside it. It must reject localhost for production and reject output if any `{{BASE_URL}}` token remains.

```js
export function validateProductionBaseUrl(rawUrl) {
  const url = new URL(rawUrl);
  if (url.protocol !== "https:" || url.hostname === "localhost" || url.hostname === "127.0.0.1") {
    throw new Error("Production base URL must be a non-local HTTPS origin");
  }
  return url.origin;
}
```

- [ ] **Step 7: Write deployment, privacy, installation, and test documentation**

`DEPLOY.md` documents static HTTPS hosting, cache invalidation, manifest regeneration, and rollback to the previous static bundle. `INSTALL_MAC.md` gives exact Mac Excel sideload, restart, trust, and uninstall steps. `PRIVACY.md` states that workbook data remains inside Excel and lists the absence of accounts, telemetry, and uploads. `TEST_CHECKLIST.md` contains one checkbox for every chart type, advanced option, invalid-selection case, table command, restart persistence, and privacy verification.

```markdown
## Mac 验收

- [ ] Excel 重启后仍显示“中金工具”页签
- [ ] “格式化表格”产生红色表头和白色数据区
- [ ] “斑马纹”只改变数据行填充
- [ ] 工作簿公式和数字格式前后完全一致
```

- [ ] **Step 8: Run the complete automated gate**

```bash
npm test
npm run typecheck
npm run build
npm run manifest:dev
npm run manifest:validate
```

Expected: every command exits 0 with no skipped tests and no unresolved manifest token.

- [ ] **Step 9: Perform the Mac Excel 16.111.2 smoke test**

Start the local HTTPS site, sideload `manifest/manifest.dev.xml`, restart Excel, and verify the `中金工具` tab. In a copy of `/Users/sky/Desktop/个人/实习作品集/01_CICC/造纸周度数据库_260125.xlsx`, test representative selections from `Summary`, `Summary new ver`, and `国内木片价格（周度）`. Confirm native editability, correct series/category mapping, exact six-color order, chart rollback on invalid inputs, red-header/white-body formatting, independent zebra fills, and unchanged formulas/number formats.

- [ ] **Step 10: Package an authorized production release**

After the user identifies or authorizes the HTTPS host, deploy the static `dist/` contents, then run the release packager with that exact origin. Validate the generated production manifest and sideload it once on the target Mac. This step changes external deployment state and therefore requires the user's authorization at execution time.

- [ ] **Step 11: Commit release readiness**

```bash
git add tests/fixtures tests/app/workflows.test.ts scripts/package-release.mjs tests/manifest/packageRelease.test.ts docs package.json package-lock.json
git commit -m "docs: add release and Mac acceptance workflow"
```

## Plan Self-Review Results

- Spec coverage: all ten chart types, direct and advanced generation, exact color palette, title/size/legend/orientation/label/trendline options, two independent table commands, privacy, HTTPS deployment, errors, rollback, and Mac Excel 16.111.2 acceptance have an owning task.
- Scope check: selection parsing, chart generation, table formatting, Ribbon UI, and deployment are tightly coupled parts of one Excel add-in and produce one independently testable application; splitting them into separate project plans would delay the first usable build.
- Type consistency: `SelectionSnapshot` flows from `ExcelGateway` to `parseSelection`; `ParsedSelection` flows to `buildChartPlan`; `ChartPlan` plus `ChartStylePlan` flow back to `ExcelGateway.createChart`; table commands use `TableFormatPlan` and never reuse chart contracts.
- Privacy check: no task adds a data API, telemetry library, workbook-content log, or upload mechanism.
- Scope exclusions: organization charts, formatting existing charts, alternate table shades, and Windows-only plugin code remain absent from every implementation task.
