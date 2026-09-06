import { resolve } from "node:path";
import { defineConfig } from "vite";
import { getHttpsServerOptions } from "office-addin-dev-certs";

export default defineConfig(async ({ command }) => {
  const useDevelopmentHttps = command === "serve" && process.env.VITEST === undefined;

  return {
    server: {
      ...(useDevelopmentHttps ? { https: await getHttpsServerOptions() } : {}),
      // Word 加载项保留 3000；Excel 固定使用独立端口，避免 manifest
      // 误连到另一加载项的开发服务器而导致 Ribbon/任务窗格失效。
      port: 3001,
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
