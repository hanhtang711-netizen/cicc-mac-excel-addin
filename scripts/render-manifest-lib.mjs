export function renderManifest(template, rawBaseUrl) {
  const baseUrl = rawBaseUrl.replace(/\/$/, "");
  const isLocalHttps = baseUrl === "https://localhost:3000";
  if (!baseUrl.startsWith("https://") || (!isLocalHttps && baseUrl.includes("localhost"))) {
    throw new Error("Base URL must use HTTPS");
  }

  const rendered = template.replaceAll("{{BASE_URL}}", baseUrl);
  if (rendered.includes("{{BASE_URL}}")) {
    throw new Error("Manifest contains an unresolved BASE_URL token");
  }

  return rendered;
}
