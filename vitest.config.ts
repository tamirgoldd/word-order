import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["packages/*/test/**/*.test.ts"],
    coverage: {
      include: ["packages/core/src/**/*.ts"],
      reporter: ["text", "html"]
    }
  }
});
