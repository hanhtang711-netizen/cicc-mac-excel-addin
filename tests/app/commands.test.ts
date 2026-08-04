// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest";

describe("command composition root", () => {
  afterEach(() => {
    vi.resetModules();
    vi.unstubAllGlobals();
  });

  it("computes one capability snapshot and shares it with chart and table services", async () => {
    vi.stubGlobal("Office", { onReady: vi.fn() });
    const { createCommandDependencies } = await import("../../src/commands");
    const requirements = { isSetSupported: vi.fn().mockReturnValue(false) };

    const dependencies = createCommandDependencies(requirements);
    const chartError = await dependencies.charts.create("column").catch((caught: unknown) => caught);
    const tableError = await dependencies.tables.formatStandard().catch((caught: unknown) => caught);

    expect(requirements.isSetSupported).toHaveBeenCalledOnce();
    expect(chartError).toMatchObject({ code: "unsupported_api" });
    expect(tableError).toMatchObject({ code: "unsupported_api" });
  });
});
