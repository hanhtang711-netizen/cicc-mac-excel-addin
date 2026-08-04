import { resolve } from "node:path";
import { defineConfig } from "vite";
import basicSsl from "@vitejs/plugin-basic-ssl";

export default defineConfig({
  plugins: [basicSsl()],
  server: {
    https: {},
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
});
