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
      "https://localhost:3001/",
    );
    expect(xml).toContain("https://localhost:3001/commands.html");
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

  it("assigns distinct semantic icons to each ribbon command", async () => {
    const template = await readFile(resolve(projectRoot, "manifest/manifest.template.xml"), "utf8");
    const iconResidFor = (controlId: string) => {
      const control = template.match(
        new RegExp(`<Control xsi:type="(?:Menu|Button)" id="${controlId}">[\\s\\S]*?<Icon><bt:Image size="16" resid="([^"]+)"`),
      );
      return control?.[1];
    };

    expect(iconResidFor("CreateChartMenu")).toBe("ChartIcon.16");
    expect(iconResidFor("OpenAdvancedChartPane")).toBe("AdvancedIcon.16");
    expect(iconResidFor("FormatCiccTable")).toBe("FormatIcon.16");
    expect(iconResidFor("ApplyZebraStripe")).toBe("ZebraIcon.16");
    for (const iconId of ["ChartIcon", "AdvancedIcon", "FormatIcon", "ZebraIcon"]) {
      for (const size of [16, 32, 64, 80]) {
        expect(template).toContain(`<bt:Image id="${iconId}.${size}"`);
      }
    }
  });

  it("uses the chart mark as the Excel developer-add-in icon", async () => {
    const template = await readFile(resolve(projectRoot, "manifest/manifest.template.xml"), "utf8");

    expect(template).toContain('<IconUrl DefaultValue="{{BASE_URL}}/assets/chart-32.png" />');
    expect(template).toContain('<HighResolutionIconUrl DefaultValue="{{BASE_URL}}/assets/chart-80.png" />');
  });
});
