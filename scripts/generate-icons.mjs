import { mkdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import sharp from "sharp";

const assetsDirectory = fileURLToPath(new URL("../public/assets/", import.meta.url));
const iconSizes = [16, 32, 64, 80];

await mkdir(assetsDirectory, { recursive: true });

for (const size of iconSizes) {
  const icon = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 80 80" width="${size}" height="${size}">
    <rect width="80" height="80" fill="#640000"/>
    <path d="M57 19a27 27 0 1 0 0 42" fill="none" stroke="#ffffff" stroke-width="10" stroke-linecap="square"/>
  </svg>`;
  await sharp(Buffer.from(icon)).png().toFile(`${assetsDirectory}/icon-${size}.png`);
}
