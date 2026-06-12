const { defineConfig } = require("@playwright/test");

module.exports = defineConfig({
  testDir: "./tests/e2e",
  timeout: 30000,
  expect: {
    timeout: 7000
  },
  fullyParallel: false,
  use: {
    trace: "retain-on-failure"
  }
});
