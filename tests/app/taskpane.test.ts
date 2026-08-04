// @vitest-environment jsdom

import { beforeEach, describe, expect, it, vi } from "vitest";
import { createAdvancedChartService, initializeAdvancedChartPane } from "../../src/taskpane";
import { AddinError } from "../../src/core/errors";

function renderPane(): HTMLFormElement {
  document.body.innerHTML = `
    <main id="app">
      <form id="advanced-chart-form">
        <fieldset id="chart-form-controls" disabled>
          <select id="kind" name="kind"><option value="column">普通柱形图</option><option value="scatterTrend">散点图</option></select>
          <select id="size-preset" name="sizePreset"><option value="medium">中</option><option value="custom">自定义</option></select>
          <div id="custom-size-fields" hidden><input id="width-cm" name="widthCm"><input id="height-cm" name="heightCm"></div>
          <select id="legend-position" name="legendPosition"><option value="bottom">下</option></select>
          <select id="orientation" name="orientation"><option value="columns">按列</option></select>
          <input id="title" name="title">
          <input id="show-data-labels" name="showDataLabels" type="checkbox">
          <div id="trendline-field" hidden><input id="add-trendline" name="addTrendline" type="checkbox"></div>
          <button type="submit">生成图表</button>
        </fieldset>
      </form>
      <p id="status" aria-live="polite"></p>
    </main>`;
  return document.querySelector("form") as HTMLFormElement;
}

describe("advanced chart task pane", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    renderPane();
  });

  it("enables the form and shows only the fields relevant to the selected options", () => {
    initializeAdvancedChartPane({ create: vi.fn().mockResolvedValue({ ok: true, warnings: [] }) });
    const form = document.querySelector("form") as HTMLFormElement;
    const controls = document.querySelector("fieldset") as HTMLFieldSetElement;
    const kind = form.elements.namedItem("kind") as HTMLSelectElement;
    const size = form.elements.namedItem("sizePreset") as HTMLSelectElement;

    expect(controls.disabled).toBe(false);
    expect(document.querySelector<HTMLElement>("#custom-size-fields")?.hidden).toBe(true);
    expect(document.querySelector<HTMLElement>("#trendline-field")?.hidden).toBe(true);

    size.value = "custom";
    size.dispatchEvent(new Event("change"));
    kind.value = "scatterTrend";
    kind.dispatchEvent(new Event("change"));

    expect(document.querySelector<HTMLElement>("#custom-size-fields")?.hidden).toBe(false);
    expect(document.querySelector<HTMLElement>("#trendline-field")?.hidden).toBe(false);
  });

  it("announces success and restores the submit button after generation", async () => {
    const create = vi.fn().mockResolvedValue({ ok: true, warnings: [] });
    initializeAdvancedChartPane({ create });
    const form = document.querySelector("form") as HTMLFormElement;
    const button = form.querySelector("button") as HTMLButtonElement;

    form.dispatchEvent(new Event("submit", { cancelable: true }));
    expect(button.disabled).toBe(true);
    expect(document.querySelector("#status")?.textContent).toBe("正在生成图表…");

    await vi.waitFor(() => expect(button.disabled).toBe(false));
    expect(document.querySelector("#status")?.textContent).toBe("图表已生成。");
    expect(create).toHaveBeenCalledWith("column", expect.objectContaining({ sizePreset: "medium" }));
  });

  it("keeps the pane usable and gives an actionable error for invalid custom size", async () => {
    const create = vi.fn();
    initializeAdvancedChartPane({ create });
    const form = document.querySelector("form") as HTMLFormElement;
    const size = form.elements.namedItem("sizePreset") as HTMLSelectElement;
    size.value = "custom";
    size.dispatchEvent(new Event("change"));

    form.dispatchEvent(new Event("submit", { cancelable: true }));

    await vi.waitFor(() => expect((form.querySelector("button") as HTMLButtonElement).disabled).toBe(false));
    expect(create).not.toHaveBeenCalled();
    expect(document.querySelector("#status")?.textContent).toContain("请填写有效的自定义宽度和高度");
  });

  it("shows the shared unsupported-API correction inline", async () => {
    const create = vi.fn().mockRejectedValue(new AddinError("unsupported_api"));
    initializeAdvancedChartPane({ create });
    const form = document.querySelector("form") as HTMLFormElement;

    form.dispatchEvent(new Event("submit", { cancelable: true }));

    await vi.waitFor(() => expect((form.querySelector("button") as HTMLButtonElement).disabled).toBe(false));
    expect(document.querySelector("#status")?.textContent).toBe(
      "当前 Excel 版本不支持此图表类型，请升级 Excel。",
    );
  });

  it("shows the shared non-blocking warning inline", async () => {
    const create = vi.fn().mockResolvedValue({
      ok: true,
      warnings: ["series_palette_reused"],
    });
    initializeAdvancedChartPane({ create });
    const form = document.querySelector("form") as HTMLFormElement;

    form.dispatchEvent(new Event("submit", { cancelable: true }));

    await vi.waitFor(() => expect((form.querySelector("button") as HTMLButtonElement).disabled).toBe(false));
    expect(document.querySelector("#status")?.textContent).toBe(
      "图表已生成。系列超过六个，后续系列将循环使用中金配色。",
    );
  });

  it("computes one capability snapshot for the advanced-pane service", async () => {
    const requirements = { isSetSupported: vi.fn().mockReturnValue(false) };
    const service = createAdvancedChartService(requirements);

    const error = await service.create("column").catch((caught: unknown) => caught);

    expect(requirements.isSetSupported).toHaveBeenCalledOnce();
    expect(error).toMatchObject({ code: "unsupported_api" });
  });
});
