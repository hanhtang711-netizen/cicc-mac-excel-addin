import { readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { renderManifest } from "./render-manifest-lib.mjs";

const [baseUrl, outputPath] = process.argv.slice(2);
if (!baseUrl || !outputPath) {
  throw new Error("Usage: node scripts/render-manifest.mjs <base-url> <output-path>");
}

const templatePath = fileURLToPath(new URL("../manifest/manifest.template.xml", import.meta.url));
const template = await readFile(templatePath, "utf8");
const manifest = renderManifest(template, baseUrl);
await writeFile(outputPath, manifest, "utf8");
