import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { renderManifest } from "../../scripts/render-manifest-lib.mjs";
import { CHART_CATALOG } from "../../src/charts/chartCatalog";

const projectRoot = resolve(import.meta.dirname, "../..");

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

  it("keeps manifest execute actions in parity with the command catalog", async () => {
    const template = await readFile(resolve(projectRoot, "manifest/manifest.template.xml"), "utf8");
    const executeActions = Array.from(
      template.matchAll(
        /<Action xsi:type="ExecuteFunction">\s*<FunctionName>([^<]+)<\/FunctionName>\s*<\/Action>/g,
      ),
      ([, actionId]) => actionId,
    );

    expect(executeActions).toEqual([
      ...Object.values(CHART_CATALOG).map(({ actionId }) => actionId),
      "formatCiccTable",
      "applyZebraStripe",
    ]);
    expect(new Set(executeActions).size).toBe(executeActions.length);
  });

  it("uses a chart menu and a separate advanced-chart task-pane action", async () => {
    const template = await readFile(resolve(projectRoot, "manifest/manifest.template.xml"), "utf8");
    const advancedAction = template.match(
      /<Control xsi:type="Button" id="OpenAdvancedChartPane">[\s\S]*?<Action xsi:type="ShowTaskpane">[\s\S]*?<SourceLocation resid="Taskpane.Url"\s*\/>[\s\S]*?<\/Action>[\s\S]*?<\/Control>/,
    );

    expect(template).toContain('<Control xsi:type="Menu" id="CreateChartMenu">');
    expect(advancedAction).not.toBeNull();
    expect(advancedAction?.[0]).toContain("<TaskpaneId>openAdvancedChartPane</TaskpaneId>");
    expect(advancedAction?.[0]).not.toContain("<FunctionName>");
  });
});
