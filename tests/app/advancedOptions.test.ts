// @vitest-environment jsdom

import { beforeEach, describe, expect, it } from "vitest";
import { readAdvancedOptions } from "../../src/app/advancedOptions";

function createForm(): HTMLFormElement {
  document.body.innerHTML = `
    <form id="advanced-chart-form">
      <select name="kind"><option value="column" selected>普通柱形图</option></select>
      <select name="orientation"><option value="auto" selected>自动</option><option value="columns">按列</option><option value="rows">按行</option></select>
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
      options: { orientation: "auto" },
    });
  });

  it("accepts an explicit row orientation", () => {
    const form = createForm();
    setValue(form, "orientation", "rows");

    expect(readAdvancedOptions(form)).toEqual({
      kind: "column",
      options: { orientation: "rows" },
    });
  });

  it("contains no visual override controls", () => {
    const form = createForm();
    expect(form.elements.namedItem("title")).toBeNull();
    expect(form.elements.namedItem("sizePreset")).toBeNull();
    expect(form.elements.namedItem("legendPosition")).toBeNull();
    expect(form.elements.namedItem("showDataLabels")).toBeNull();
  });

  it.each([
    ["kind", "made-up"],
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
