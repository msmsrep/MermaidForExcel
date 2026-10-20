import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["test/**/*.test.ts"],
    // Task pane code runs in a browser; build-script tests opt into "node".
    environment: "jsdom",
    restoreMocks: true,
    unstubGlobals: true,
  },
});
