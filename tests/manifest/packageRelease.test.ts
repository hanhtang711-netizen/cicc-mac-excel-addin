import { mkdtemp, mkdir, readFile, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, dirname, join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { packageRelease, validateProductionBaseUrl } from "../../scripts/package-release.mjs";

const temporaryDirectories: string[] = [];

afterEach(async () => {
  vi.doUnmock("node:fs/promises");
  vi.resetModules();
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
    "https://localhost.",
    "https://[::ffff:127.0.0.1]",
    "https://10.0.0.1",
    "https://169.254.1.1",
    "https://172.16.0.1",
    "https://192.168.1.10",
    "https://[fc00::1]",
    "https://[fe80::1]",
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

  it("rejects an output path that reaches the project through a symlinked parent", async () => {
    const projectRoot = await createProjectFixture();
    const aliasRoot = await mkdtemp(join(tmpdir(), "cicc-release-alias-"));
    temporaryDirectories.push(aliasRoot);
    const parentAlias = join(aliasRoot, "parent-alias");
    await symlink(dirname(projectRoot), parentAlias, "dir");
    const aliasedProjectPath = join(parentAlias, basename(projectRoot));

    await expect(packageRelease({
      baseUrl: "https://excel-addon.example.com",
      outDir: aliasedProjectPath,
      projectRoot,
    })).rejects.toThrow("Release output must not overwrite the project or build input");

    expect(await readFile(join(projectRoot, "manifest", "manifest.template.xml"), "utf8"))
      .toContain("{{BASE_URL}}");
  });

  it("restores the previous release when the staged-directory swap fails", async () => {
    const projectRoot = await createProjectFixture();
    const outDir = join(projectRoot, "release-output");
    await mkdir(outDir);
    await writeFile(join(outDir, "keep.txt"), "existing", "utf8");

    vi.resetModules();
    vi.doMock("node:fs/promises", async (importOriginal) => {
      const actual = await importOriginal<typeof import("node:fs/promises")>();
      return {
        ...actual,
        rename: async (from: Parameters<typeof actual.rename>[0], to: Parameters<typeof actual.rename>[1]) => {
          if (basename(String(from)).startsWith(".release-output.tmp-") && basename(String(to)) === "release-output") {
            throw new Error("simulated staged swap failure");
          }
          return actual.rename(from, to);
        },
      };
    });
    const { packageRelease: packageReleaseWithSwapFailure } = await import("../../scripts/package-release.mjs");

    await expect(packageReleaseWithSwapFailure({
      baseUrl: "https://excel-addon.example.com",
      outDir,
      projectRoot,
    })).rejects.toThrow("simulated staged swap failure");

    await expect(readFile(join(outDir, "keep.txt"), "utf8")).resolves.toBe("existing");
  });

  it("rejects a static build that omits a manifest-referenced icon", async () => {
    const projectRoot = await createProjectFixture();
    const { rm } = await import("node:fs/promises");
    await rm(join(projectRoot, "dist", "assets", "icon-80.png"));

    await expect(packageRelease({
      baseUrl: "https://excel-addon.example.com",
      outDir: join(projectRoot, "release-output"),
      projectRoot,
    })).rejects.toThrow("Static build is missing required entry: assets/icon-80.png");
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
  await Promise.all([16, 32, 64, 80].map((size) => (
    writeFile(join(projectRoot, "dist", "assets", `icon-${size}.png`), `icon-${size}`, "utf8")
  )));
  await writeFile(
    join(projectRoot, "manifest", "manifest.template.xml"),
    '<SourceLocation DefaultValue="{{BASE_URL}}/taskpane.html" />',
    "utf8",
  );
  return projectRoot;
}
