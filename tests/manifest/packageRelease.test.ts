import { mkdtemp, mkdir, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { packageRelease, validateProductionBaseUrl } from "../../scripts/package-release.mjs";

const temporaryDirectories: string[] = [];

afterEach(async () => {
  const { rm } = await import("node:fs/promises");
  await Promise.all(temporaryDirectories.splice(0).map((path) => rm(path, { recursive: true, force: true })));
});

describe("release packager", () => {
  it("accepts a fixed HTTPS origin", () => {
    expect(validateProductionBaseUrl("https://excel-addon.example.com")).toBe(
      "https://excel-addon.example.com",
    );
    expect(validateProductionBaseUrl("https://excel-addon.example.com/")).toBe(
      "https://excel-addon.example.com",
    );
  });

  it.each([
    "http://excel-addon.example.com",
    "https://localhost:3000",
    "https://dev.localhost",
    "https://127.0.0.1",
    "https://127.0.0.2",
    "https://[::1]",
    "https://user:secret@excel-addon.example.com",
    "https://excel-addon.example.com/addin",
    "https://excel-addon.example.com/?channel=prod",
    "https://excel-addon.example.com/#latest",
    " https://excel-addon.example.com",
  ])("rejects non-production origin %s", (candidate) => {
    expect(() => validateProductionBaseUrl(candidate)).toThrow(
      "Production base URL must be a fixed, non-local HTTPS origin",
    );
  });

  it("copies a verified static build and renders a token-free production manifest", async () => {
    const projectRoot = await createProjectFixture();
    const outDir = join(projectRoot, "release-output");

    const result = await packageRelease({
      baseUrl: "https://excel-addon.example.com",
      outDir,
      projectRoot,
    });

    expect(result).toEqual({
      baseUrl: "https://excel-addon.example.com",
      outDir,
      siteDir: join(outDir, "site"),
      manifestPath: join(outDir, "manifest.production.xml"),
    });
    expect(await readFile(join(outDir, "site", "taskpane.html"), "utf8")).toBe("<main>ready</main>");
    expect(await readFile(join(outDir, "site", "assets", "app.js"), "utf8")).toBe("export const ready=true;");
    const manifest = await readFile(join(outDir, "manifest.production.xml"), "utf8");
    expect(manifest).toContain("https://excel-addon.example.com/taskpane.html");
    expect(manifest).not.toContain("{{BASE_URL}}");
  });

  it("fails before replacing an existing explicit output when the build is missing", async () => {
    const projectRoot = await createProjectFixture();
    const outDir = join(projectRoot, "release-output");
    const { rm } = await import("node:fs/promises");
    await rm(join(projectRoot, "dist"), { recursive: true });
    await mkdir(outDir);
    await writeFile(join(outDir, "keep.txt"), "existing", "utf8");

    await expect(packageRelease({
      baseUrl: "https://excel-addon.example.com",
      outDir,
      projectRoot,
    })).rejects.toThrow("Static build is missing");

    expect(await readFile(join(outDir, "keep.txt"), "utf8")).toBe("existing");
  });

  it("rejects output paths that could overwrite project inputs", async () => {
    const projectRoot = await createProjectFixture();

    await expect(packageRelease({
      baseUrl: "https://excel-addon.example.com",
      outDir: projectRoot,
      projectRoot,
    })).rejects.toThrow("Release output must not overwrite the project or build input");
    await expect(packageRelease({
      baseUrl: "https://excel-addon.example.com",
      outDir: join(projectRoot, "dist"),
      projectRoot,
    })).rejects.toThrow("Release output must not overwrite the project or build input");
  });
});

async function createProjectFixture(): Promise<string> {
  const projectRoot = await mkdtemp(join(tmpdir(), "cicc-release-test-"));
  temporaryDirectories.push(projectRoot);
  await mkdir(join(projectRoot, "dist", "assets"), { recursive: true });
  await mkdir(join(projectRoot, "manifest"), { recursive: true });
  await writeFile(join(projectRoot, "dist", "taskpane.html"), "<main>ready</main>", "utf8");
  await writeFile(join(projectRoot, "dist", "commands.html"), "<main>commands</main>", "utf8");
  await writeFile(join(projectRoot, "dist", "feedback.html"), "<main>feedback</main>", "utf8");
  await writeFile(join(projectRoot, "dist", "assets", "app.js"), "export const ready=true;", "utf8");
  await writeFile(
    join(projectRoot, "manifest", "manifest.template.xml"),
    '<SourceLocation DefaultValue="{{BASE_URL}}/taskpane.html" />',
    "utf8",
  );
  return projectRoot;
}
