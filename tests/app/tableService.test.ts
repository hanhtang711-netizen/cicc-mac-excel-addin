import { describe, expect, it, vi } from "vitest";
import { TableService } from "../../src/app/tableService";
import { AddinError } from "../../src/core/errors";
import type { SelectionSnapshot } from "../../src/core/types";
import type { AddinCapabilities } from "../../src/office/capability";

const supportedCapabilities: AddinCapabilities = {
  base: true,
  explodedPie: true,
  trendlines: true,
};

const selectionSnapshot: SelectionSnapshot = {
  worksheetName: "Data",
  address: "Data!A1:B3",
  rowIndex: 0,
  columnIndex: 0,
  rowCount: 3,
  columnCount: 2,
  values: [["Date", "Revenue"], [46023, 10], [46030, 12]],
  texts: [["Date", "Revenue"], ["2026-01-01", "10"], ["2026-01-08", "12"]],
  numberFormats: [["General", "General"], ["yyyy-mm-dd", "0"], ["yyyy-mm-dd", "0"]],
};

describe("TableService", () => {
  it("reads and applies one independent standard-table plan", async () => {
    const gateway = makeGateway();

    const result = await new TableService(gateway, supportedCapabilities).formatStandard();

    expect(result).toEqual({ ok: true, warnings: [] });
    expect(gateway.readSelection).toHaveBeenCalledOnce();
    expect(gateway.applyTablePlan).toHaveBeenCalledOnce();
    expect(gateway.applyTablePlan).toHaveBeenCalledWith(expect.objectContaining({
      kind: "standard",
      rowFills: [],
    }));
  });

  it("reads and applies one zebra-only plan", async () => {
    const gateway = makeGateway();

    const result = await new TableService(gateway, supportedCapabilities).applyZebra();

    expect(result).toEqual({ ok: true, warnings: [] });
    expect(gateway.readSelection).toHaveBeenCalledOnce();
    expect(gateway.applyTablePlan).toHaveBeenCalledOnce();
    expect(gateway.applyTablePlan).toHaveBeenCalledWith(expect.objectContaining({
      kind: "zebra",
      rowFills: expect.arrayContaining([{ rowOffset: 1, fill: "#F5F5F5" }]),
    }));
  });

  it.each([
    ["empty", { ...selectionSnapshot, address: "Data!" , rowCount: 0, columnCount: 0, values: [], texts: [], numberFormats: [] }],
    ["discontiguous", { ...selectionSnapshot, address: "Data!A1:B2,D1:E2" }],
    ["whole row", { ...selectionSnapshot, address: "Data!1:3" }],
    ["whole column", { ...selectionSnapshot, address: "Data!A:B" }],
    ["whole sheet", { ...selectionSnapshot, address: "Data!A1:XFD1048576" }],
  ])("rejects a %s selection before applying a table plan", async (_label, snapshot) => {
    const gateway = makeGateway(snapshot);

    const error = await new TableService(gateway, supportedCapabilities).formatStandard().catch((caught: unknown) => caught);

    expect(error).toBeInstanceOf(AddinError);
    expect(error).toMatchObject({ code: "invalid_selection" });
    expect(gateway.applyTablePlan).not.toHaveBeenCalled();
  });

  it.each(["formatStandard", "applyZebra"] as const)(
    "rejects %s before reading when the base API is unsupported",
    async (method) => {
      const gateway = makeGateway();
      const service = new TableService(gateway, {
        base: false,
        explodedPie: false,
        trendlines: false,
      });

      const error = await service[method]().catch((caught: unknown) => caught);

      expect(error).toBeInstanceOf(AddinError);
      expect(error).toMatchObject({ code: "unsupported_api" });
      expect(gateway.readSelection).not.toHaveBeenCalled();
      expect(gateway.applyTablePlan).not.toHaveBeenCalled();
    },
  );
});

function makeGateway(snapshot: SelectionSnapshot = selectionSnapshot) {
  return {
    readSelection: vi.fn().mockResolvedValue(snapshot),
    applyTablePlan: vi.fn().mockResolvedValue(undefined),
  };
}
