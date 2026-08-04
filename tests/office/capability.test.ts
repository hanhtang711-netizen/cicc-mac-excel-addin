import { describe, expect, it, vi } from "vitest";
import { getCapabilities } from "../../src/office/capability";

describe("Excel API capabilities", () => {
  it("requires ExcelApi 1.9 once for all target capabilities", () => {
    const requirements = { isSetSupported: vi.fn().mockReturnValue(true) };

    expect(getCapabilities(requirements)).toEqual({
      base: true,
      explodedPie: true,
      trendlines: true,
    });
    expect(requirements.isSetSupported).toHaveBeenCalledOnce();
    expect(requirements.isSetSupported).toHaveBeenCalledWith("ExcelApi", "1.9");
  });

  it("disables every target capability when ExcelApi 1.9 is unavailable", () => {
    const requirements = { isSetSupported: vi.fn().mockReturnValue(false) };

    expect(getCapabilities(requirements)).toEqual({
      base: false,
      explodedPie: false,
      trendlines: false,
    });
    expect(requirements.isSetSupported).toHaveBeenCalledOnce();
  });
});
