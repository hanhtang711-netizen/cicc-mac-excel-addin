import { afterEach, describe, expect, it, vi } from "vitest";
import { CHART_CATALOG } from "../../src/charts/chartCatalog";
import {
  createCommandHandlers,
  registerCommandHandlers,
} from "../../src/app/commandHandlers";

const chartActionIds = Object.values(CHART_CATALOG).map(({ actionId }) => actionId);
const executeActionIds = [...chartActionIds, "formatCiccTable", "applyZebraStripe"];

describe("command handlers", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("creates exactly one handler for each chart and table execute action", () => {
    const handlers = createCommandHandlers(makeDependencies());

    expect(Object.keys(handlers)).toEqual(executeActionIds);
    expect(new Set(Object.values(handlers)).size).toBe(executeActionIds.length);
    expect(handlers).not.toHaveProperty("openAdvancedChartPane");
  });

  it.each(Object.entries(CHART_CATALOG))(
    "routes the %s action to ChartService.create",
    async (kind, definition) => {
      const dependencies = makeDependencies();
      const handlers = createCommandHandlers(dependencies);
      const event = makeEvent();

      await handlers[definition.actionId]?.(event);

      expect(dependencies.charts.create).toHaveBeenCalledOnce();
      expect(dependencies.charts.create).toHaveBeenCalledWith(kind);
      expect(event.completed).toHaveBeenCalledOnce();
    },
  );

  it.each([
    ["formatCiccTable", "formatStandard"],
    ["applyZebraStripe", "applyZebra"],
  ] as const)("routes %s to TableService.%s", async (actionId, method) => {
    const dependencies = makeDependencies();
    const handlers = createCommandHandlers(dependencies);
    const event = makeEvent();

    await handlers[actionId]?.(event);

    expect(dependencies.tables[method]).toHaveBeenCalledOnce();
    expect(event.completed).toHaveBeenCalledOnce();
  });

  it("shows returned warnings and completes exactly once", async () => {
    const dependencies = makeDependencies();
    dependencies.charts.create.mockResolvedValue({ ok: true, warnings: ["series_palette_reused"] });
    const event = makeEvent();

    await createCommandHandlers(dependencies).createColumnChart?.(event);

    expect(dependencies.feedback.showWarnings).toHaveBeenCalledWith(["series_palette_reused"]);
    expect(event.completed).toHaveBeenCalledOnce();
  });

  it("does not open feedback after a successful command without warnings", async () => {
    const dependencies = makeDependencies();
    const event = makeEvent();

    await createCommandHandlers(dependencies).createColumnChart?.(event);

    expect(dependencies.feedback.showWarnings).not.toHaveBeenCalled();
    expect(dependencies.feedback.showError).not.toHaveBeenCalled();
    expect(event.completed).toHaveBeenCalledOnce();
  });

  it("reports service failures and completes exactly once", async () => {
    const dependencies = makeDependencies();
    const failure = new Error("failed");
    dependencies.charts.create.mockRejectedValue(failure);
    const event = makeEvent();

    await createCommandHandlers(dependencies).createColumnChart?.(event);

    expect(dependencies.feedback.showError).toHaveBeenCalledWith(failure);
    expect(event.completed).toHaveBeenCalledOnce();
  });

  it.each(["showWarnings", "showError"] as const)(
    "still resolves and completes exactly once when feedback.%s fails",
    async (method) => {
      const dependencies = makeDependencies();
      dependencies.feedback[method].mockRejectedValue(new Error("dialog failed"));
      if (method === "showError") {
        dependencies.charts.create.mockRejectedValue(new Error("service failed"));
      }
      const event = makeEvent();

      await expect(
        createCommandHandlers(dependencies).createColumnChart?.(event),
      ).resolves.toBeUndefined();
      expect(event.completed).toHaveBeenCalledOnce();
    },
  );

  it("associates each execute action exactly once", () => {
    const associate = vi.fn();
    vi.stubGlobal("Office", { actions: { associate } });

    registerCommandHandlers(makeDependencies());

    expect(associate).toHaveBeenCalledTimes(executeActionIds.length);
    expect(associate.mock.calls.map(([actionId]) => actionId)).toEqual(executeActionIds);
    expect(new Set(associate.mock.calls.map(([actionId]) => actionId)).size).toBe(
      executeActionIds.length,
    );
  });
});

function makeDependencies() {
  return {
    charts: {
      create: vi.fn().mockResolvedValue({ ok: true, warnings: [] }),
    },
    tables: {
      formatStandard: vi.fn().mockResolvedValue({ ok: true, warnings: [] }),
      applyZebra: vi.fn().mockResolvedValue({ ok: true, warnings: [] }),
    },
    feedback: {
      showError: vi.fn().mockResolvedValue(undefined),
      showWarnings: vi.fn().mockResolvedValue(undefined),
    },
  };
}

function makeEvent(): Office.AddinCommands.Event {
  return { source: { id: "test-control" }, completed: vi.fn() };
}
