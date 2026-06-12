const { defineConfig } = require("@playwright/test");

module.exports = defineConfig({
  testDir: "./tests/e2e-google-docs",
  timeout: 45000,
  expect: {
    timeout: 15000
  },
  fullyParallel: false,
  outputDir: "google-docs-smoke-results",
  reporter: [["list"]],
  use: {
    screenshot: "off",
    trace: "off",
    video: "off"
  }
});
