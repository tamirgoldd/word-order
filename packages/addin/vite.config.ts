import { defineConfig } from "vite";

export default defineConfig({
  base: process.env.BASE_PATH ? `${process.env.BASE_PATH.replace(/\/$/, "")}/addin/` : "/",
  build: {
    target: "es2022",
    sourcemap: true,
    outDir: "dist"
  }
});
