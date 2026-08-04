# CICC Skill Parity and Usability Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (- [ ]) syntax for tracking.

**Goal:** Make all supported chart and table commands reliable in Excel for Mac and strictly match the approved CICC Skill style.

**Architecture:** Keep the parser → planner → style plan → Excel gateway flow. Add automatic data orientation, replace the universal series-color operation with chart-type-specific formatting, lock visual options to one CICC contract, and preserve transactional rollback with stage diagnostics.

**Tech Stack:** TypeScript 7, Office.js, Vite 8, Vitest 4, jsdom, Microsoft Excel for Mac 16.111.2

## Global Constraints

- Source of truth: 技能_CICC图表+底稿_260803.md, build_amzn_cicc_xlsx.py, and verify_amzn_charts.py.
- Fixed chart size: 16×9 cm.
- No chart title, axis title, data label, source text, gridline, or outer border.
- Bottom, non-overlay, black 9pt legend; set Arial where Office.js has one font slot.
- Columns/bars: solid fill and cleared outline. Lines: colored 1.5pt smooth line and no series fill.
- Keep full series names, including parenthetical units and currency.
- Standard table and zebra commands stay independent.
- Standard table: 8pt, left aligned, no borders, red header, white body, row height 16, worksheet gridlines off.
- Zebra: first data row #F5F5F5, next row white, fill-only changes.
- Do not modify the original 造纸周度数据库_260125.xlsx.
- Preserve the existing uncommitted development-certificate changes until their separate final commit.

---

### Task 1: Automatic Orientation and Exact Mac Regression Fixture

**Files:**
- Modify: src/core/types.ts
- Modify: src/core/selectionParser.ts
- Modify: tests/fixtures/selectionCases.ts
- Modify: tests/core/selectionParser.test.ts
- Modify: tests/app/workflows.test.ts

**Interfaces:**
- Produce SeriesOrientationMode = "auto" | "columns" | "rows".
- Produce parseSelection(snapshot, orientation = "auto"): ParsedSelection.

- [ ] **Step 1: Add the exact A1:C13 fixture**

~~~ts
export const macLineRegressionFixture = fixture(
  "Fixture!A1:C13",
  [
    [null, "A", "B"],
    ...Array.from({ length: 12 }, (_, index) => [36892 + index * 365, index + 1, 5]),
  ],
  [
    ["", "A", "B"],
    ...Array.from({ length: 12 }, (_, index) => [
      String(2001 + index) + "-01-01",
      String(index + 1),
      "5",
    ]),
  ],
  [
    ["General", "General", "General"],
    ...Array.from({ length: 12 }, () => ["yyyy-mm-dd", "0", "0"]),
  ],
);
~~~

- [ ] **Step 2: Write failing auto-orientation tests**

~~~ts
expect(parseSelection(macLineRegressionFixture)).toMatchObject({
  orientation: "columns",
  categoryAddress: "'Fixture'!$A$2:$A$13",
  series: [
    { name: "A", valuesAddress: "'Fixture'!$B$2:$B$13" },
    { name: "B", valuesAddress: "'Fixture'!$C$2:$C$13" },
  ],
});

expect(parseSelection(horizontalOrientationFixture)).toMatchObject({
  orientation: "rows",
  categoryAddress: "'Fixture'!$B$1:$D$1",
});
~~~

- [ ] **Step 3: Verify failure**

~~~bash
npm test -- tests/core/selectionParser.test.ts tests/app/workflows.test.ts
~~~

Expected: the horizontal fixture fails because the current default is always columns.

- [ ] **Step 4: Implement deterministic inference**

Add SeriesOrientationMode to src/core/types.ts. In parseSelection, resolve explicit rows/columns directly. For auto: use columns when only the top row supplies text labels, rows when only the first column supplies text labels, and use the longer data dimension as the category direction when both edges contain text labels. Use columns for a square tie.

~~~ts
function inferOrientation(snapshot: SelectionSnapshot): SeriesOrientation {
  const topLabels = countTopRowSeriesLabels(snapshot);
  const leftLabels = countFirstColumnSeriesLabels(snapshot);
  if (topLabels > 0 && leftLabels === 0) return "columns";
  if (leftLabels > 0 && topLabels === 0) return "rows";
  return snapshot.columnCount > snapshot.rowCount ? "rows" : "columns";
}
~~~

- [ ] **Step 5: Verify and commit**

~~~bash
npm test -- tests/core/selectionParser.test.ts tests/app/workflows.test.ts
npm run typecheck
git add src/core/types.ts src/core/selectionParser.ts tests/fixtures/selectionCases.ts tests/core/selectionParser.test.ts tests/app/workflows.test.ts
git commit -m "fix: infer chart series orientation"
~~~

---

### Task 2: Lock the CICC Chart Plan and Style Contract

**Files:**
- Modify: src/core/types.ts
- Modify: src/charts/chartPlanner.ts
- Modify: src/charts/chartStyle.ts
- Modify: src/app/chartService.ts
- Modify: tests/charts/chartPlanner.test.ts
- Modify: tests/charts/chartStyle.test.ts
- Modify: tests/app/workflows.test.ts

**Interfaces:**
- ChartOptions retains orientation only.
- ChartSeriesStyle is fill-no-border, line, scatter, or pie-points.
- buildChartStylePlan consumes kind, count, formats, and placement inputs.

- [ ] **Step 1: Write failing strict-contract tests**

~~~ts
const plan = buildChartPlan(parseSelection(titleRowFixture), "lineMarkers");
expect(plan).not.toHaveProperty("title");
expect(plan).not.toHaveProperty("showDataLabels");
expect(plan).not.toHaveProperty("sizePreset");

const style = buildChartStylePlan({
  kind: "line",
  seriesCount: 2,
  sourceFormat: "0.0%",
  categoryFormat: "yyyy-mm-dd",
  selectionColumn: 0,
  selectionColumnCount: 3,
});
expect(style).toMatchObject({
  seriesStyle: "line",
  widthPoints: cmToPoints(16),
  heightPoints: cmToPoints(9),
  legendPosition: "bottom",
  legendOverlay: false,
  legendFontSizePoints: 9,
  showTitle: false,
  showDataLabels: false,
  showGridlines: false,
  lineWidthPoints: 1.5,
  smoothLines: true,
});
~~~

- [ ] **Step 2: Verify failure**

~~~bash
npm test -- tests/charts/chartPlanner.test.ts tests/charts/chartStyle.test.ts tests/app/workflows.test.ts
~~~

Expected: old title/size/label fields, 11.5×6.7 size, 8pt text, and gray-gridline fields cause failures.

- [ ] **Step 3: Narrow types and plans**

~~~ts
export interface ChartOptions {
  orientation?: SeriesOrientationMode;
}

export type ChartSeriesStyle =
  | "fill-no-border"
  | "line"
  | "scatter"
  | "pie-points";
~~~

Remove visual overrides from ChartPlan. Keep addLinearTrendline true only for scatterTrend.

- [ ] **Step 4: Implement fixed style constants**

~~~ts
const CHART_WIDTH_CM = 16;
const CHART_HEIGHT_CM = 9;
const LEGEND_SIZE_POINTS = 9;
const LINE_WIDTH_POINTS = 1.5;
~~~

Map columns/bars to fill-no-border, line types to line, scatterTrend to scatter, and pie types to pie-points. Preserve the six-color palette, date/percentage formats, warning, and safe right/below placement.

- [ ] **Step 5: Update ChartService, verify, and commit**

~~~bash
npm test -- tests/charts/chartPlanner.test.ts tests/charts/chartStyle.test.ts tests/app/workflows.test.ts
npm run typecheck
git add src/core/types.ts src/charts/chartPlanner.ts src/charts/chartStyle.ts src/app/chartService.ts tests/charts/chartPlanner.test.ts tests/charts/chartStyle.test.ts tests/app/workflows.test.ts
git commit -m "refactor: lock CICC chart style contract"
~~~

---

### Task 3: Fix Mac Line Charts with Type-Specific Series Formatting

**Files:**
- Modify: src/office/excelGateway.ts
- Modify: tests/office/excelGateway.test.ts

**Interfaces:**
- Produce applySeriesStyles(series, plan, style).
- Line/scatter styling must never call ChartSeriesFormat.fill.setSolidColor.

- [ ] **Step 1: Write the failing Office.js boundary tests**

~~~ts
await new ExcelGateway().createChart(
  { ...plan, kind: "line", excelType: "line" },
  { ...style, seriesStyle: "line" },
);
expect(fake.series[0].format.fill.setSolidColor).not.toHaveBeenCalled();
expect(fake.series[0].format.line.color).toBe("#640000");
expect(fake.series[0].format.line.weight).toBe(1.5);
expect(fake.series[0].smooth).toBe(true);

await new ExcelGateway().createChart(plan, {
  ...style,
  seriesStyle: "fill-no-border",
});
expect(fake.series[0].format.fill.setSolidColor).toHaveBeenCalledWith("#640000");
expect(fake.series[0].format.line.clear).toHaveBeenCalledOnce();
~~~

- [ ] **Step 2: Verify failure**

~~~bash
npm test -- tests/office/excelGateway.test.ts
~~~

Expected: the line case calls fill.setSolidColor, matching Mac Excel HResult 0xA7120001.

- [ ] **Step 3: Expand the harness**

Add smooth, markerStyle, marker colors, line.weight, line.clear, and trendline.format.line to fake series objects.

- [ ] **Step 4: Replace setSeriesColor**

~~~ts
function applySeriesStyles(
  series: readonly Excel.ChartSeries[],
  plan: ChartPlan,
  style: ChartStylePlan,
): void {
  series.forEach((item, index) => {
    const color = style.seriesColors[index];
    if (color === undefined) return;
    if (style.seriesStyle === "fill-no-border") {
      item.format.fill.setSolidColor(color);
      item.format.line.clear();
      return;
    }
    item.format.line.color = color;
    item.format.line.weight = style.lineWidthPoints;
    item.smooth = style.smoothLines;
    item.markerStyle = plan.kind === "lineMarkers" ? "Automatic" : "None";
  });
}
~~~

Style scatter series and returned trendlines without accessing series fill. Keep pie-point fill separate.

- [ ] **Step 5: Verify and commit**

~~~bash
npm test -- tests/office/excelGateway.test.ts
npm run typecheck
git add src/office/excelGateway.ts tests/office/excelGateway.test.ts
git commit -m "fix: style chart series by type"
~~~

---

### Task 4: Apply Strict Common Chart Styling

**Files:**
- Modify: src/office/excelGateway.ts
- Modify: tests/office/excelGateway.test.ts

**Interfaces:**
- Produce applyCommonChartStyle and applyAxisStyle.
- Pie types never access axes.

- [ ] **Step 1: Write failing common-style assertions**

~~~ts
expect(fake.chart.title.visible).toBe(false);
expect(fake.chart.dataLabels.showValue).toBe(false);
expect(fake.chart.legend).toMatchObject({
  visible: true,
  position: "Bottom",
  overlay: false,
});
expect(fake.chart.legend.format.font).toMatchObject({
  name: "Arial",
  size: 9,
  color: "#000000",
});
expect(fake.categoryAxis).toMatchObject({
  visible: true,
  position: "Minimum",
  majorTickMark: "Outside",
  tickLabelPosition: "NextToAxis",
});
expect(fake.valueAxis.majorGridlines.visible).toBe(false);
expect(fake.valueAxis.minorGridlines.visible).toBe(false);
~~~

- [ ] **Step 2: Verify failure**

~~~bash
npm test -- tests/office/excelGateway.test.ts
~~~

Expected: old title/label behavior, missing overlay/axis settings, and visible gray gridlines fail.

- [ ] **Step 3: Implement strict common style**

Set fixed 16×9 points, white chart/plot fill, cleared chart border, hidden title and labels, bottom non-overlay legend, Arial 9pt black legend text.

For category and value axes set visible true, majorTickMark Outside, tickLabelPosition NextToAxis, hidden axis titles, hidden major/minor gridlines, Arial 8pt black text, bottom/minimum category positioning, and automatic primary value positioning. Preserve date and percentage number formats.

- [ ] **Step 4: Verify pie guard and rollback**

Keep the existing pie test that throws if axes are accessed. Ensure any style failure deletes the new chart and syncs the deletion.

- [ ] **Step 5: Verify and commit**

~~~bash
npm test -- tests/office/excelGateway.test.ts tests/charts/chartStyle.test.ts
npm run typecheck
git add src/office/excelGateway.ts tests/office/excelGateway.test.ts
git commit -m "feat: apply strict CICC chart styling"
~~~

---

### Task 5: Simplify the Advanced Pane

**Files:**
- Modify: src/app/advancedOptions.ts
- Modify: taskpane.html
- Modify: src/taskpane.ts
- Modify: src/app/userFeedback.ts
- Modify: tests/app/advancedOptions.test.ts
- Modify: tests/app/taskpane.test.ts
- Modify: tests/app/userFeedback.test.ts

**Interfaces:**
- readAdvancedOptions returns chart kind and auto/columns/rows only.

- [ ] **Step 1: Write failing simplified-form tests**

~~~ts
expect(readAdvancedOptions(form)).toEqual({
  kind: "column",
  options: { orientation: "auto" },
});
expect(form.elements.namedItem("title")).toBeNull();
expect(form.elements.namedItem("sizePreset")).toBeNull();
expect(form.elements.namedItem("legendPosition")).toBeNull();
expect(form.elements.namedItem("showDataLabels")).toBeNull();
~~~

- [ ] **Step 2: Verify failure**

~~~bash
npm test -- tests/app/advancedOptions.test.ts tests/app/taskpane.test.ts
~~~

Expected: current code requires removed visual controls.

- [ ] **Step 3: Implement the strict parser**

~~~ts
const orientations = ["auto", "columns", "rows"] as const;

return {
  kind: approvedChartKind(valueOf(form, "kind")),
  options: {
    orientation: approvedValue(
      valueOf(form, "orientation"),
      orientations,
      "invalid_orientation",
    ),
  },
};
~~~

- [ ] **Step 4: Simplify markup and initialization**

Keep chart type, orientation, submit button, and live status only. Remove title, size, legend, label, and trendline controls and all conditional-field handlers.

- [ ] **Step 5: Remove obsolete custom-size feedback, verify, and commit**

~~~bash
npm test -- tests/app/advancedOptions.test.ts tests/app/taskpane.test.ts tests/app/userFeedback.test.ts
npm run typecheck
git add src/app/advancedOptions.ts taskpane.html src/taskpane.ts src/app/userFeedback.ts tests/app/advancedOptions.test.ts tests/app/taskpane.test.ts tests/app/userFeedback.test.ts
git commit -m "refactor: simplify strict CICC chart pane"
~~~

---

### Task 6: Mirror Strict CICC Table Formatting

**Files:**
- Modify: src/core/types.ts
- Modify: src/tables/tablePlanner.ts
- Modify: src/office/excelGateway.ts
- Modify: tests/tables/tablePlanner.test.ts
- Modify: tests/app/tableService.test.ts
- Modify: tests/app/workflows.test.ts
- Modify: tests/office/excelGateway.test.ts

**Interfaces:**
- Standard plan uses fixed left/center alignment, 8pt, row height 16, clearBorders, and hideWorksheetGridlines.
- Zebra plan changes row fills only.

- [ ] **Step 1: Write failing parity tests**

~~~ts
expect(buildStandardTablePlan(selectionSnapshot)).toMatchObject({
  header: { fill: "#8A2626", fontColor: "#FFFFFF", bold: true, fontSize: 8 },
  body: { fill: "#FFFFFF", fontColor: "#000000", bold: false, fontSize: 8 },
  horizontalAlignment: "left",
  verticalAlignment: "center",
  rowHeight: 16,
  clearBorders: true,
  hideWorksheetGridlines: true,
});
expect(buildZebraPlan(selectionSnapshot).rowFills).toEqual([
  { rowOffset: 1, fill: "#F5F5F5" },
  { rowOffset: 2, fill: "#FFFFFF" },
  { rowOffset: 3, fill: "#F5F5F5" },
]);
~~~

- [ ] **Step 2: Verify failure**

~~~bash
npm test -- tests/tables/tablePlanner.test.ts tests/app/tableService.test.ts tests/app/workflows.test.ts
~~~

Expected: old 10.5pt, type-dependent alignment, thin borders, and white-first zebra fail.

- [ ] **Step 3: Implement fixed plans**

Remove per-column alignment and border color/weight. Return fixed 8pt styles, left/center alignment, wrap text, row height 16, cleared borders, hidden worksheet gridlines, and gray-first zebra rows.

- [ ] **Step 4: Implement gateway writes**

Apply range/header fonts and fills, set horizontal/vertical alignment and wrap text, clear EdgeTop/Bottom/Left/Right/InsideVertical/InsideHorizontal borders, set every selected row height to 16, and disable the active worksheet gridline view. Keep bounded column autofit and do not autofit row heights.

- [ ] **Step 5: Prove zebra is fill-only**

Keep unexpected-write spies for font, borders, alignment, and content. Expect gray/white/gray writes only.

- [ ] **Step 6: Verify and commit**

~~~bash
npm test -- tests/tables/tablePlanner.test.ts tests/app/tableService.test.ts tests/app/workflows.test.ts tests/office/excelGateway.test.ts
npm run typecheck
git add src/core/types.ts src/tables/tablePlanner.ts src/office/excelGateway.ts tests/tables/tablePlanner.test.ts tests/app/tableService.test.ts tests/app/workflows.test.ts tests/office/excelGateway.test.ts
git commit -m "feat: mirror CICC table styling"
~~~

---

### Task 7: Preserve Runtime Stage Diagnostics and Rollback

**Files:**
- Modify: src/office/excelGateway.ts
- Modify: tests/office/excelGateway.test.ts
- Modify: tests/app/userFeedback.test.ts

**Interfaces:**
- AddinError details contain stage, cause, and optional rollbackError.

- [ ] **Step 1: Write a failing stage test**

~~~ts
expect(error).toMatchObject({
  code: "excel_runtime_error",
  details: {
    stage: "series-style",
    cause: expect.any(Error),
  },
});
expect(fake.chart.delete).toHaveBeenCalledOnce();
~~~

- [ ] **Step 2: Verify failure**

~~~bash
npm test -- tests/office/excelGateway.test.ts tests/app/userFeedback.test.ts
~~~

Expected: current details contain only the raw cause.

- [ ] **Step 3: Add staged sync**

Track create, series-style, common-style, final-sync, and rollback. Set the stage before each queued call group and sync. On failure, delete and sync the new chart, then throw AddinError with the original stage/cause; include rollbackError if deletion fails.

- [ ] **Step 4: Keep feedback concise**

Prove that user-facing feedback remains the existing Excel runtime message and does not expose internal stage data.

- [ ] **Step 5: Verify and commit**

~~~bash
npm test -- tests/office/excelGateway.test.ts tests/app/userFeedback.test.ts
npm run typecheck
git add src/office/excelGateway.ts tests/office/excelGateway.test.ts tests/app/userFeedback.test.ts
git commit -m "fix: preserve chart failure stages"
~~~

---

### Task 8: Full Automated and Mac Excel Acceptance

**Files:**
- Modify: docs/TEST_CHECKLIST.md

**Interfaces:**
- Produce a green automated suite and recorded Mac acceptance against the test copy.

- [ ] **Step 1: Add exact checklist cases**

Document A1:C13 line regression, all ten chart buttons, column/row layouts, dates, text categories, percentages, negatives, blanks, palette reuse, strict style elements, both table buttons separately/sequentially, persistence after reopen, and log inspection.

- [ ] **Step 2: Run the full automated gate**

~~~bash
npm test
npm run typecheck
npm run build
npm run manifest:dev
npm run manifest:validate
git diff --check
~~~

Expected: every command succeeds.

- [ ] **Step 3: Verify trusted HTTPS**

~~~bash
curl --fail --silent --show-error --output /dev/null https://localhost:3000/taskpane.html
curl --fail --silent --show-error --output /dev/null https://localhost:3000/commands.html
~~~

Expected: both exit 0 without insecure TLS flags.

- [ ] **Step 4: Run the Mac A1:C13 regression**

Open /private/tmp/CICC_Addin_Smoke_260125.xlsx, select Sheet1!A1:C13, click 中金工具 → 普通折线图, and confirm a native chart remains present.

- [ ] **Step 5: Inspect the chart**

Confirm series A/B, 2001–2012 date categories, colors #640000/#8A2626, 1.5pt smooth lines, 16×9 cm, bottom non-overlay legend, and no title/labels/gridlines.

- [ ] **Step 6: Run representative chart/table acceptance**

Create every chart type from a compatible fixture. Run standard table and zebra separately and sequentially. Use only temporary/test copies.

- [ ] **Step 7: Save, reopen, and inspect Excel logs**

Save and reopen the temporary test workbook. Confirm chart persistence and no new negative-HResult ChartFill.setSolidColor call from a line-series path.

- [ ] **Step 8: Final verification and commit**

~~~bash
npm test
npm run typecheck
npm run build
npm run manifest:dev
npm run manifest:validate
git diff --check
git add docs/TEST_CHECKLIST.md
git commit -m "docs: record strict CICC Mac acceptance"
~~~

Expected: all checks pass; the original workbook remains unchanged.

---

## Completion Criteria

- A1:C13 creates and keeps a native two-series line chart in Mac Excel.
- Line/scatter series never invoke ChartSeriesFormat.fill.setSolidColor.
- Ten chart types and two table commands pass representative Mac acceptance.
- All visible chart/table output matches the Skill contract except the accepted Office.js dual-font limitation.
- Tests, typecheck, build, manifest render/validation, and diff-check pass.
- Original user files remain unchanged.
