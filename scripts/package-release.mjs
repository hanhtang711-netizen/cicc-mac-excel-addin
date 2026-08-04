import { execFile } from "node:child_process";
import { randomUUID } from "node:crypto";
import { cp, lstat, mkdir, mkdtemp, readFile, realpath, rename, rm, stat, writeFile } from "node:fs/promises";
import { isIP } from "node:net";
import { basename, dirname, isAbsolute, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { renderManifest } from "./render-manifest-lib.mjs";

const execFileAsync = promisify(execFile);
const scriptPath = fileURLToPath(import.meta.url);
const defaultProjectRoot = resolve(dirname(scriptPath), "..");
const REQUIRED_STATIC_ENTRIES = [
  "taskpane.html",
  "commands.html",
  "feedback.html",
  "assets/icon-16.png",
  "assets/icon-32.png",
  "assets/icon-64.png",
  "assets/icon-80.png",
];

export function validateProductionBaseUrl(rawUrl) {
  if (typeof rawUrl !== "string" || rawUrl.length === 0 || rawUrl !== rawUrl.trim()) {
    throw productionUrlError();
  }
  let url;
  try {
    url = new URL(rawUrl);
  } catch {
    throw productionUrlError();
  }

  const hostname = normalizeHostname(url.hostname);
  const isLocal = hostname === "localhost" || hostname.endsWith(".localhost") ||
    hostname.endsWith(".local") || isIP(hostname) !== 0;
  const hasCredentials = url.username.length > 0 || url.password.length > 0;
  const hasNonOriginParts = url.pathname !== "/" || url.search.length > 0 || url.hash.length > 0;

  if (url.protocol !== "https:" || isLocal || hasCredentials || hasNonOriginParts) {
    throw productionUrlError();
  }
  return url.origin;
}

export async function packageRelease({ baseUrl, outDir, projectRoot = defaultProjectRoot }) {
  const normalizedBaseUrl = validateProductionBaseUrl(baseUrl);
  if (typeof outDir !== "string" || outDir.trim().length === 0) {
    throw new Error("Release output path is required");
  }

  const resolvedProjectRoot = resolve(projectRoot);
  const resolvedOutDir = resolve(resolvedProjectRoot, outDir);
  const distDir = resolve(resolvedProjectRoot, "dist");
  const manifestTemplatePath = resolve(resolvedProjectRoot, "manifest", "manifest.template.xml");
  await assertSafeOutputPath(resolvedOutDir, resolvedProjectRoot, distDir, manifestTemplatePath);

  const [template] = await Promise.all([
    readFile(manifestTemplatePath, "utf8").catch(() => {
      throw new Error(`Manifest template is missing: ${manifestTemplatePath}`);
    }),
    verifyStaticBuild(distDir),
  ]);
  const manifest = renderManifest(template, normalizedBaseUrl);
  if (manifest.includes("{{BASE_URL}}")) {
    throw new Error("Production manifest contains an unresolved BASE_URL token");
  }

  await mkdir(dirname(resolvedOutDir), { recursive: true });
  const stagedRelease = await mkdtemp(resolve(dirname(resolvedOutDir), `.${basename(resolvedOutDir)}.tmp-`));
  const stagedSite = resolve(stagedRelease, "site");

  try {
    await cp(distDir, stagedSite, { recursive: true, force: true });
    await writeFile(resolve(stagedRelease, "manifest.production.xml"), manifest, "utf8");
    await assertSafeOutputPath(resolvedOutDir, resolvedProjectRoot, distDir, manifestTemplatePath);
    await replaceReleaseDirectory(stagedRelease, resolvedOutDir);
  } finally {
    await rm(stagedRelease, { recursive: true, force: true });
  }

  return {
    baseUrl: normalizedBaseUrl,
    outDir: resolvedOutDir,
    siteDir: resolve(resolvedOutDir, "site"),
    manifestPath: resolve(resolvedOutDir, "manifest.production.xml"),
  };
}

export async function runReleaseCli(argv = process.argv.slice(2)) {
  const { baseUrl, outDir } = parseArguments(argv);
  validateProductionBaseUrl(baseUrl);
  await execFileAsync(process.platform === "win32" ? "npm.cmd" : "npm", ["run", "build"], {
    cwd: defaultProjectRoot,
  });
  const result = await packageRelease({ baseUrl, outDir, projectRoot: defaultProjectRoot });
  process.stdout.write(`Release site: ${result.siteDir}\nProduction manifest: ${result.manifestPath}\n`);
  return result;
}

function parseArguments(argv) {
  const values = new Map();
  for (let index = 0; index < argv.length; index += 2) {
    const flag = argv[index];
    const value = argv[index + 1];
    if ((flag !== "--base-url" && flag !== "--out") || value === undefined || values.has(flag)) {
      throw usageError();
    }
    values.set(flag, value);
  }
  if (argv.length !== 4 || !values.has("--base-url") || !values.has("--out")) {
    throw usageError();
  }
  return { baseUrl: values.get("--base-url"), outDir: values.get("--out") };
}

async function verifyStaticBuild(distDir) {
  const distStats = await stat(distDir).catch(() => undefined);
  if (!distStats?.isDirectory()) {
    throw new Error(`Static build is missing: ${distDir}`);
  }
  await Promise.all(REQUIRED_STATIC_ENTRIES.map(async (entry) => {
    const entryPath = resolve(distDir, entry);
    const entryStats = await stat(entryPath).catch(() => undefined);
    if (!entryStats?.isFile()) {
      throw new Error(`Static build is missing required entry: ${entry}`);
    }
  }));
}

async function assertSafeOutputPath(outDir, projectRoot, distDir, manifestTemplatePath) {
  const outputLeaf = await lstat(outDir).catch((error) => {
    if (isMissingPathError(error)) return undefined;
    throw error;
  });
  if (outputLeaf?.isSymbolicLink()) {
    throw new Error("Release output must not be a symbolic link");
  }

  const [canonicalOutDir, canonicalProjectRoot, canonicalDistDir, canonicalManifestDir] = await Promise.all([
    canonicalizePotentialPath(outDir),
    realpath(projectRoot),
    canonicalizePotentialPath(distDir),
    canonicalizePotentialPath(dirname(manifestTemplatePath)),
  ]);
  const isFilesystemRoot = dirname(canonicalOutDir) === canonicalOutDir;
  const overwritesProject = canonicalOutDir === canonicalProjectRoot ||
    isInside(canonicalOutDir, canonicalProjectRoot);
  const overwritesInput = canonicalOutDir === canonicalDistDir ||
    isInside(canonicalDistDir, canonicalOutDir) ||
    isInside(canonicalOutDir, canonicalDistDir) ||
    canonicalOutDir === canonicalManifestDir ||
    isInside(canonicalManifestDir, canonicalOutDir) ||
    isInside(canonicalOutDir, canonicalManifestDir);
  if (isFilesystemRoot || overwritesProject || overwritesInput) {
    throw new Error("Release output must not overwrite the project or build input");
  }
}

async function canonicalizePotentialPath(inputPath) {
  let cursor = resolve(inputPath);
  const missingSegments = [];

  while (true) {
    try {
      const canonicalAncestor = await realpath(cursor);
      return resolve(canonicalAncestor, ...missingSegments.reverse());
    } catch (error) {
      if (!isMissingPathError(error)) {
        throw error;
      }
      const parent = dirname(cursor);
      if (parent === cursor) {
        throw error;
      }
      missingSegments.push(basename(cursor));
      cursor = parent;
    }
  }
}

async function replaceReleaseDirectory(stagedRelease, outDir) {
  const existingOutput = await stat(outDir).catch((error) => {
    if (isMissingPathError(error)) return undefined;
    throw error;
  });
  if (existingOutput === undefined) {
    await rename(stagedRelease, outDir);
    return;
  }

  const backupDir = resolve(dirname(outDir), `.${basename(outDir)}.backup-${randomUUID()}`);
  await rename(outDir, backupDir);
  try {
    await rename(stagedRelease, outDir);
  } catch (swapError) {
    try {
      await rename(backupDir, outDir);
    } catch (rollbackError) {
      throw new AggregateError(
        [swapError, rollbackError],
        "Release replacement failed and the previous release could not be restored",
      );
    }
    throw swapError;
  }
  await rm(backupDir, { recursive: true, force: true });
}

function normalizeHostname(hostname) {
  return hostname.toLowerCase().replace(/^\[|\]$/g, "").replace(/\.+$/, "");
}

function isMissingPathError(error) {
  return error instanceof Error && "code" in error && error.code === "ENOENT";
}

function isInside(parent, candidate) {
  const pathFromParent = relative(parent, candidate);
  return pathFromParent.length > 0 && !pathFromParent.startsWith("..") && !isAbsolute(pathFromParent);
}

function productionUrlError() {
  return new Error("Production base URL must be a fixed, non-local HTTPS origin");
}

function usageError() {
  return new Error("Usage: npm run release -- --base-url <https-origin> --out <directory>");
}

if (process.argv[1] && resolve(process.argv[1]) === resolve(scriptPath)) {
  runReleaseCli().catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
}
