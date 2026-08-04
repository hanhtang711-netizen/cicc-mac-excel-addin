import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

const projectRoot = resolve(import.meta.dirname, "../..");

describe("foundation configuration", () => {
  it("keeps Office.js declarations available to TypeScript", async () => {
    const tsconfig = JSON.parse(
      await readFile(resolve(projectRoot, "tsconfig.json"), "utf8"),
    ) as { compilerOptions: { types?: string[] } };

    expect(tsconfig.compilerOptions.types).toContain("office-js");
  });

  it("loads task-pane CSS directly from the HTML entry point", async () => {
    const taskpane = await readFile(resolve(projectRoot, "taskpane.html"), "utf8");
    const taskpaneScript = await readFile(
      resolve(projectRoot, "src/taskpane.ts"),
      "utf8",
    );

    expect(taskpane).toContain('<link rel="stylesheet" href="/src/taskpane.css" />');
    expect(taskpaneScript).not.toContain('import "./taskpane.css"');
  });
});
