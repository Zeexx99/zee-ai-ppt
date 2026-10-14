import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { readFileSync } from "node:fs";

const pkg = JSON.parse(readFileSync(new URL("./package.json", import.meta.url), "utf-8"));
const buildDate = new Date().toISOString().split("T")[0];

export default defineConfig({
  base: "./",
  plugins: [react()],
  define: {
    __APP_VERSION__: JSON.stringify(pkg.version || "1.2.0"),
    __BUILD_DATE__: JSON.stringify(buildDate),
  },
  server: {
    port: 3300,
    strictPort: true,
    cors: true,
  },
  preview: {
    port: 3300,
    strictPort: true,
    cors: true,
  },
  build: {
    outDir: "dist",
    sourcemap: false,
    assetsDir: "assets",
    rollupOptions: {
      output: {
        manualChunks: undefined,
      },
    },
  },
});
