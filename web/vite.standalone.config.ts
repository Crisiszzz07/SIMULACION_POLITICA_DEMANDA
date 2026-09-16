import { defineConfig } from "vite";
import { resolve } from "node:path";

// Bundles all browser code in one classic script for direct file:// use.
export default defineConfig({
  build: {
    emptyOutDir: false,
    lib: {
      entry: resolve(__dirname, "src/main.ts"),
      name: "SimuladorInventario",
      formats: ["iife"],
      fileName: () => "simulador.js",
    },
    outDir: "standalone",
  },
});
