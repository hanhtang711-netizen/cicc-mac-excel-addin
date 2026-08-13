import { describe, expect, it, vi } from "vitest";
import { ChartService } from "../../src/app/chartService";
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
  address: "Data!A1:C4",
  rowIndex: 0,
  columnIndex: 2,
  rowCount: 4,
  columnCount: 3,
  values: [["Date", "Pulp", "Paper"], [1, 10, 20], [2, 11, 21], [3, 12, 22]],
  texts: [
    ["Date", "Pulp", "Paper"],
    ["2026-01-01", "10.0%", "20.0%"],
    ["2026-01-08", "11.0%", "21.0%"],
    ["2026-01-15", "12.0%", "22.0%"],
  ],
  numberFormats: [
    ["General", "General", "General"],
    ["yyyy-mm-dd", "0.0%", "0.0%"],
    ["yyyy-mm-dd", "0.0%", "0.0%"],
    ["yyyy-mm-dd", "0.0%", "0.0%"],
  ],
};

describe("ChartService", () => {
  it("reads, parses, plans, and creates one chart", async () => {
    const gateway = {
      readSelection: vi.fn().mockResolvedValue(selectionSnapshot),
      readRange: vi.fn().mockResolvedValue(selectionSnapshot),
      createChart: vi.fn().mockResolvedValue(undefined),
    };
    const service = new ChartService(gateway, supportedCapabilities);

    const result = await service.create("column");

    expect(result).toEqual({ ok: true, warnings: [] });
    expect(gateway.createChart).toHaveBeenCalledOnce();
    expect(gateway.createChart).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: "column",
        excelType: "columnClustered",
        sourceAddress: "Data!A1:C4",
      }),
      expect.objectContaining({
        widthPoints: expect.any(Number),
        valueAxisNumberFormat: "0.0%",
        categoryAxisNumberFormat: "yyyy-mm-dd",
        placement: { side: "right", gutterPoints: 18 },
      }),
    );
  });

  it("preserves typed validation errors and performs no write", async () => {
    const gateway = {
      readSelection: vi.fn().mockResolvedValue({
        ...selectionSnapshot,
        values: [["Name", "Value"], ["alpha", 1], ["beta", 2]],
        texts: [["Name", "Value"], ["alpha", "1"], ["beta", "2"]],
        numberFormats: [["General", "General"], ["General", "General"], ["General", "General"]],
        rowCount: 3,
        columnCount: 2,
      }),
      readRange: vi.fn().mockResolvedValue(selectionSnapshot),
      createChart: vi.fn().mockResolvedValue(undefined),
    };
    const service = new ChartService(gateway, supportedCapabilities);

    const error = await service.create("scatterTrend").catch((caught: unknown) => caught);

    expect(error).toBeInstanceOf(AddinError);
    expect(error).toMatchObject({ code: "scatter_requires_numeric_x" });
    expect(gateway.createChart).not.toHaveBeenCalled();
  });

  it("rejects an unsupported base API before reading the selection", async () => {
    const gateway = makeGateway();
    const service = new ChartService(gateway, {
      base: false,
      explodedPie: false,
      trendlines: false,
    });

    const error = await service.create("column").catch((caught: unknown) => caught);

    expect(error).toBeInstanceOf(AddinError);
    expect(error).toMatchObject({ code: "unsupported_api" });
    expect(gateway.readSelection).not.toHaveBeenCalled();
    expect(gateway.createChart).not.toHaveBeenCalled();
  });

  it("rejects an unsupported exploded pie before reading the selection", async () => {
    const gateway = makeGateway();
    const service = new ChartService(gateway, {
      ...supportedCapabilities,
      explodedPie: false,
    });

    const error = await service.create("pieExploded").catch((caught: unknown) => caught);

    expect(error).toBeInstanceOf(AddinError);
    expect(error).toMatchObject({ code: "unsupported_api" });
    expect(gateway.readSelection).not.toHaveBeenCalled();
    expect(gateway.createChart).not.toHaveBeenCalled();
  });

  it("rejects an unsupported default scatter trendline before reading the selection", async () => {
    const gateway = makeGateway();
    const service = new ChartService(gateway, {
      ...supportedCapabilities,
      trendlines: false,
    });

    const error = await service.create("scatterTrend").catch((caught: unknown) => caught);

    expect(error).toBeInstanceOf(AddinError);
    expect(error).toMatchObject({ code: "unsupported_api" });
    expect(gateway.readSelection).not.toHaveBeenCalled();
    expect(gateway.createChart).not.toHaveBeenCalled();
  });

});

function makeGateway() {
  return {
    readSelection: vi.fn().mockResolvedValue(selectionSnapshot),
    readRange: vi.fn().mockResolvedValue(selectionSnapshot),
    createChart: vi.fn().mockResolvedValue(undefined),
  };
}
