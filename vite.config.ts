import { resolve } from "node:path";
import { defineConfig } from "vite";
import { getHttpsServerOptions } from "office-addin-dev-certs";

export default defineConfig(async ({ command }) => {
  const useDevelopmentHttps = command === "serve" && process.env.VITEST === undefined;

  return {
    server: {
      ...(useDevelopmentHttps ? { https: await getHttpsServerOptions() } : {}),
      port: 3000,
      strictPort: true,
    },
    build: {
      rollupOptions: {
        input: {
          commands: resolve(import.meta.dirname, "commands.html"),
          feedback: resolve(import.meta.dirname, "feedback.html"),
          taskpane: resolve(import.meta.dirname, "taskpane.html"),
        },
      },
    },
  };
});
