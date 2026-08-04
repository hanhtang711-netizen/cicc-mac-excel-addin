import { describe, expect, it } from "vitest";
import { renderManifest } from "../../scripts/render-manifest-lib.mjs";

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
});
