const { test, expect, chromium } = require("@playwright/test");
const fs = require("node:fs");
const path = require("node:path");

const extensionPath = path.resolve(__dirname, "..", "..");
const enabled = process.env.ENABLE_GOOGLE_DOCS_SMOKE === "true";
const smokeUrl = process.env.GOOGLE_DOCS_SMOKE_URL || "";

function isGoogleDocsDocumentUrl(value) {
  try {
    const url = new URL(value);
    return url.hostname === "docs.google.com" && url.pathname.startsWith("/document/d/");
  } catch (_error) {
    return false;
  }
}

async function launchExtensionContext(testInfo) {
  const userDataDir = testInfo.outputPath("chromium-profile");
  fs.mkdirSync(userDataDir, { recursive: true });
  return chromium.launchPersistentContext(userDataDir, {
    headless: process.env.PLAYWRIGHT_HEADLESS === "1",
    viewport: { width: 1280, height: 900 },
    args: [
      "--disable-crash-reporter",
      "--disable-crashpad",
      `--disable-extensions-except=${extensionPath}`,
      `--load-extension=${extensionPath}`
    ]
  });
}

async function openSmokeDoc(context) {
  const page = await context.newPage();
  await page.goto(smokeUrl, { waitUntil: "domcontentloaded", timeout: 45000 });
  await expect
    .poll(() => page.url(), { message: "Google Docs smoke URL should remain on a document page" })
    .toContain("docs.google.com/document/d/");
  return page;
}

test.describe("optional real Google Docs smoke tests", () => {
  test.skip(!enabled, "Set ENABLE_GOOGLE_DOCS_SMOKE=true to run optional real Google Docs smoke tests.");
  test.skip(enabled && !smokeUrl, "Set GOOGLE_DOCS_SMOKE_URL to a dedicated non-private test Google Doc URL.");

  test("extension loads on a real Google Doc", async ({}, testInfo) => {
    expect(isGoogleDocsDocumentUrl(smokeUrl), "GOOGLE_DOCS_SMOKE_URL must be a Google Docs document URL").toBe(true);
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openSmokeDoc(context);
      const widget = page.locator("#ace-widget");

      await expect(widget).toBeVisible();
      await expect(widget).toHaveAttribute("data-ace-extension-ready", "true");
      await expect(widget).toHaveAttribute("data-ace-supported-surface", "google-docs");
      await expect(widget).not.toContainText(/extension context is not available|refresh the google doc/i);
    } finally {
      await context.close();
    }
  });

  test("extension detects real document context", async ({}, testInfo) => {
    expect(isGoogleDocsDocumentUrl(smokeUrl), "GOOGLE_DOCS_SMOKE_URL must be a Google Docs document URL").toBe(true);
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openSmokeDoc(context);
      const widget = page.locator("#ace-widget");

      await expect(widget).toHaveAttribute("data-ace-supported-surface", "google-docs");
      await expect(widget).toHaveAttribute("data-ace-document-id", /.+/);
      await expect(widget).toHaveAttribute("data-ace-manuscript-surface-id", /.+/);
      await expect(widget).toHaveAttribute("data-ace-widget-state", /idle|prompt|loading|picker|confirm|active|completed|catch-up|recovery/);
    } finally {
      await context.close();
    }
  });
});
