# Task 7 Report: Ribbon Command Registration and Action Wiring

## Status

Implemented and locally verified. The official Microsoft manifest validator could not be reached from the sandbox because DNS resolution failed; the exact error is recorded below.

## Implementation

- Added a pure `createCommandHandlers(dependencies)` factory with exactly twelve unique ExecuteFunction handlers: the ten action IDs read from `CHART_CATALOG`, plus `formatCiccTable` and `applyZebraStripe`.
- Added `registerCommandHandlers(dependencies)`, which associates those twelve handlers exactly once. `openAdvancedChartPane` is not associated in JavaScript.
- Added one shared command wrapper. Service failures, warning feedback failures, and error feedback failures cannot prevent `event.completed()` from being called exactly once.
- Added `createCommandDependencies()` as the runtime composition root and delayed construction/registration until `Office.onReady`; startup performs no workbook operation.
- Added same-origin `DialogFeedback`, a third Vite entry (`feedback.html`), and a feedback page that maps a query-string code to hardcoded local Chinese copy using `textContent`. It performs no network requests and renders no workbook-derived HTML.
- Reworked the `中金工具` Ribbon into a `生成图表` group with a ten-item menu and a separate `高级生成` ShowTaskpane action whose `TaskpaneId` is `openAdvancedChartPane`; retained independent `格式化表格` and `斑马纹` buttons in the `表格` group.
- Regenerated `manifest/manifest.dev.xml` and added manifest/action parity tests.

## TDD Evidence

Initial focused run failed for the expected missing behavior:

- `tests/app/commandHandlers.test.ts`: module `src/app/commandHandlers.ts` did not exist.
- `tests/manifest/renderManifest.test.ts`: the manifest had no chart `Menu` or advanced ShowTaskpane action.

A later focused RED verified the exact thirteenth action ID: the test expected `openAdvancedChartPane` and failed against the initial `AdvancedChartPane`; the manifest was then corrected and the test passed.

## Final Verification

- `npm run manifest:dev`: PASS.
- `npm test -- tests/app/commandHandlers.test.ts tests/manifest/renderManifest.test.ts`: PASS, 22/22 tests.
- `npm test`: PASS, 84/84 tests across 10 files.
- `npm run typecheck`: PASS.
- `npm run build`: PASS; Vite produced all three HTML entries, including `dist/feedback.html`.
- `xmllint --noout manifest/manifest.template.xml manifest/manifest.dev.xml`: PASS.
- `git diff --check`: PASS.
- `npm run manifest:validate`: BLOCKED by environment/network, exit 1:

  `FetchError: request to https://validationgateway.omex.office.net/package/api/check?clientId=devx failed, reason: getaddrinfo ENOTFOUND validationgateway.omex.office.net`

## Self-review

- ExecuteFunction names match the catalog plus the two fixed table action IDs, with no duplicates.
- Each returned handler is a distinct function and maps to one existing service method; no parsing or Office.js formatting logic is duplicated.
- The ShowTaskpane action is the thirteenth manifest action and has no `FunctionName` or JavaScript association.
- Feedback input is reduced to stable local codes; unknown errors map to `excel_runtime_error`, and the dialog writes only hardcoded text through `textContent`.
- No unrelated source files were changed.

## Remaining concern

Official schema/service validation still needs to be rerun in an environment that can resolve and reach `validationgateway.omex.office.net`. Local XML well-formedness, rendering, tests, typecheck, and build all pass.
