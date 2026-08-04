// @vitest-environment jsdom

import { beforeEach, describe, expect, it } from "vitest";
import { readAdvancedOptions } from "../../src/app/advancedOptions";

function createForm(): HTMLFormElement {
  document.body.innerHTML = `
    <form id="advanced-chart-form">
      <select name="kind"><option value="column" selected>普通柱形图</option></select>
      <select name="sizePreset"><option value="medium" selected>中</option></select>
      <input name="widthCm" value="">
      <input name="heightCm" value="">
      <select name="legendPosition"><option value="bottom" selected>下</option></select>
      <select name="orientation"><option value="columns" selected>按列</option></select>
      <input name="title" value="Weekly prices">
      <input name="showDataLabels" type="checkbox">
      <input name="addTrendline" type="checkbox" checked>
    </form>`;
  return document.querySelector("form") as HTMLFormElement;
}

describe("advanced chart options", () => {
  beforeEach(() => {
    createForm();
  });

  it("returns approved defaults", () => {
    expect(readAdvancedOptions(createForm())).toEqual({
      kind: "column",
      options: {
        title: "Weekly prices",
        sizePreset: "medium",
        legendPosition: "bottom",
        orientation: "columns",
        showDataLabels: false,
      },
    });
  });

  it("trims an empty title and omits a trendline for non-scatter charts", () => {
    const form = createForm();
    (form.elements.namedItem("title") as HTMLInputElement).value = "   ";

    expect(readAdvancedOptions(form)).toEqual({
      kind: "column",
      options: {
        title: undefined,
        sizePreset: "medium",
        legendPosition: "bottom",
        orientation: "columns",
        showDataLabels: false,
      },
    });
  });

  it("accepts finite positive custom dimensions", () => {
    const form = createForm();
    setValue(form, "sizePreset", "custom");
    setValue(form, "widthCm", "14.5");
    setValue(form, "heightCm", "8");

    expect(readAdvancedOptions(form)).toMatchObject({
      options: { sizePreset: "custom", widthCm: 14.5, heightCm: 8 },
    });
  });

  it.each(["0", "-1", "Infinity", "not-a-number"])("rejects invalid custom width %s", (width) => {
    const form = createForm();
    setValue(form, "sizePreset", "custom");
    setValue(form, "widthCm", width);
    setValue(form, "heightCm", "8");

    expect(() => readAdvancedOptions(form)).toThrow("unsupported_layout");
  });

  it.each([
    ["kind", "made-up"],
    ["sizePreset", "giant"],
    ["legendPosition", "center"],
    ["orientation", "diagonal"],
  ])("rejects an unsupported %s value", (name, value) => {
    const form = createForm();
    const select = form.elements.namedItem(name) as HTMLSelectElement;
    const option = document.createElement("option");
    option.value = value;
    option.selected = true;
    select.append(option);

    expect(() => readAdvancedOptions(form)).toThrow("unsupported_layout");
  });
});

function setValue(form: HTMLFormElement, name: string, value: string): void {
  const input = form.elements.namedItem(name) as HTMLInputElement | HTMLSelectElement;
  if (input instanceof HTMLSelectElement && !Array.from(input.options).some((option) => option.value === value)) {
    input.append(new Option(value, value));
  }
  input.value = value;
}
