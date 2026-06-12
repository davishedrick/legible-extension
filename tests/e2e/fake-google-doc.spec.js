const { test, expect, chromium } = require("@playwright/test");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const extensionPath = path.resolve(__dirname, "..", "..");
const fakeDocPath = path.resolve(__dirname, "..", "fixtures", "fake-google-doc.html");

async function launchExtensionContext(testInfo) {
  const userDataDir = testInfo.outputPath("chromium-profile");
  fs.mkdirSync(userDataDir, { recursive: true });
  const context = await chromium.launchPersistentContext(userDataDir, {
    headless: process.env.PLAYWRIGHT_HEADLESS === "1",
    viewport: { width: 1280, height: 900 },
    args: [
      "--disable-crash-reporter",
      "--disable-crashpad",
      `--disable-extensions-except=${extensionPath}`,
      `--load-extension=${extensionPath}`
    ]
  });
  await context.route("https://docs.google.com/document/d/**", async (route) => {
    await route.fulfill({
      status: 200,
      contentType: "text/html",
      body: fs.readFileSync(fakeDocPath, "utf8")
    });
  });
  return context;
}

async function openFakeDoc(context, {
  documentId = "doc-a",
  title = "Document A",
  tabTitle = "Tab A",
  words = 10305,
  projectStartingWordCount = null,
  projectBaselineEstablished = null,
  projectStartingWordCountSource = ""
} = {}) {
  const page = await context.newPage();
  const params = new URLSearchParams({
    scriptorFakeDocs: "1",
    title,
    tabTitle,
    words: String(words)
  });
  if (projectStartingWordCount !== null && projectStartingWordCount !== undefined) {
    params.set("projectStartingWordCount", String(projectStartingWordCount));
  }
  if (projectBaselineEstablished !== null && projectBaselineEstablished !== undefined) {
    params.set("projectBaselineEstablished", String(projectBaselineEstablished));
  }
  if (projectStartingWordCountSource) {
    params.set("projectStartingWordCountSource", projectStartingWordCountSource);
  }
  await page.goto(
    `https://docs.google.com/document/d/${encodeURIComponent(documentId)}/edit?${params.toString()}`
  );
  await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-extension-detected", "true");
  await expect(page.locator("#ace-widget")).toBeVisible();
  return page;
}

async function openControls(page) {
  const widget = page.locator("#ace-widget");
  const openButton = widget.getByLabel(/Open Scriptor session controls|Restore timer/);
  if (await openButton.isVisible()) {
    await openButton.click();
  }
  return widget;
}

async function bindFakeProject(page, expectedWordCountText = "10,305 words") {
  const widget = await openControls(page);
  await widget.getByRole("button", { name: "Bind project" }).click();
  await widget.getByRole("button", { name: /Fake Project A/ }).click();
  await expect(widget).toContainText(`Verified manuscript size: ${expectedWordCountText}`);
  return widget;
}

async function setFakeWordCount(page, count, dispatchInput = false) {
  await page.evaluate(
    ({ value, dispatch }) => window.__fakeDocsSetWordCount(value, dispatch),
    { value: count, dispatch: dispatchInput }
  );
  await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-word-count", String(count));
}

async function setFakeVisibleWordCount(page, count) {
  await page.evaluate((value) => window.__fakeDocsSetVisibleWordCount(value), count);
  await expect(page.getByTestId("word-count")).toContainText(`${count.toLocaleString("en-US")} words`);
}

async function setFakeWordCountUnavailable(page, mode) {
  await page.evaluate((nextMode) => window.__fakeDocsSetWordCountUnavailable(nextMode), mode);
  await expect(page.getByTestId("word-count")).toContainText("Word count unavailable");
}

async function setFakeDocument(page, documentId, title, count) {
  await page.evaluate(
    ({ id, nextTitle, words }) => window.__fakeDocsSetDocument(id, nextTitle, words),
    { id: documentId, nextTitle: title, words: count }
  );
  await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-document-id", documentId);
}

async function startWritingSession(page) {
  const widget = await openControls(page);
  await widget.getByRole("button", { name: "Start writing" }).click();
  await expect(widget).toContainText("Writing", { timeout: 10000 });
  return widget;
}

test.describe("fake Google Docs extension flows", () => {
  test("extension loads and detects fake Document A", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context);
      const widget = await openControls(page);

      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-document-id", "doc-a");
      await expect(widget).toContainText("Not bound");
      await expect(widget).toContainText("Document A");
    } finally {
      await context.close();
    }
  });

  test("extension reads the fake document word count", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 10305 });

      await bindFakeProject(page);

      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-starting-word-count", "10305");
    } finally {
      await context.close();
    }
  });

  test("bound prompt lists project document and tab titles", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, {
        documentId: "doc-hollowfield",
        title: "Hollowfield v7",
        tabTitle: "Chapter 7",
        words: 60240
      });
      const widget = await bindFakeProject(page, "60,240 words");
      await widget.getByLabel("Close Scriptor controls").click();

      await openControls(page);

      await expect(widget.locator(".ace-field-readout").filter({ hasText: "Project" })).toContainText("Fake Project A");
      await expect(widget.locator(".ace-field-readout").filter({ hasText: "Document" })).toContainText("Hollowfield v7");
      await expect(widget.locator(".ace-field-readout").filter({ hasText: "Tab" })).toContainText("Chapter 7");
    } finally {
      await context.close();
    }
  });

  test("first bind uses actual fake document count without logging catch-up", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 10305 });

      await bindFakeProject(page);

      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-current-word-count", "10305");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");
      await expect(page.locator("#ace-widget")).not.toContainText("+305");
    } finally {
      await context.close();
    }
  });

  test("verified bind replaces provisional zero project baseline", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, {
        words: 320,
        projectStartingWordCount: 0,
        projectBaselineEstablished: true,
        projectStartingWordCountSource: "provisional"
      });

      await bindFakeProject(page, "320 words");

      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-current-word-count", "320");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-starting-word-count", "320");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");
    } finally {
      await context.close();
    }
  });

  test("catch-up is detected before a new writing session starts", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 10305 });
      const widget = await bindFakeProject(page);
      await setFakeWordCount(page, 10800);

      await widget.getByRole("button", { name: "Start writing" }).click();

      await expect(widget).toContainText("Catch-up");
      await expect(widget).toContainText("+495 words");
      await expect(widget).not.toContainText("Writing 00");
    } finally {
      await context.close();
    }
  });

  test("manual sync after binding at 1,114 reports no change when the document is unchanged", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 1114 });
      const widget = await bindFakeProject(page, "1,114 words");

      await widget.getByRole("button", { name: "Sync document changes" }).click();

      await expect(widget).toContainText("Document changes are already synced.");
      await expect(widget).not.toContainText("-1,114");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-current-word-count", "1114");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");
    } finally {
      await context.close();
    }
  });

  test("manual sync rejects a false visible zero when the bound document API still reports 1,114", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 1114 });
      const widget = await bindFakeProject(page, "1,114 words");
      await setFakeVisibleWordCount(page, 0);

      await widget.getByRole("button", { name: "Sync document changes" }).click();

      await expect(widget).toContainText("Document changes are already synced.");
      await expect(widget).not.toContainText("-1,114");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-current-word-count", "1114");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");
    } finally {
      await context.close();
    }
  });

  test("manual sync suppresses catch-up when false visible zero conflicts with changed API count", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 60265 });
      const widget = await bindFakeProject(page, "60,265 words");
      await setFakeWordCount(page, 60251, false);
      await setFakeVisibleWordCount(page, 0);

      await widget.getByRole("button", { name: "Sync document changes" }).click();

      await expect(widget).toContainText("Could not verify document changes right now.");
      await expect(widget).not.toContainText("Catch-up");
      await expect(widget).not.toContainText("-14 words");
      await expect(widget).not.toContainText("-60,265");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-current-word-count", "60265");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");
    } finally {
      await context.close();
    }
  });

  test("manual sync fails safely when the word-count read fails after binding", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 1114 });
      const widget = await bindFakeProject(page, "1,114 words");
      await setFakeWordCountUnavailable(page, "failure");

      await widget.getByRole("button", { name: "Sync document changes" }).click();

      await expect(widget).toContainText("Could not verify document changes right now.");
      await expect(widget).not.toContainText("-1,114");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-current-word-count", "1114");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");
    } finally {
      await context.close();
    }
  });

  test("manual sync does not coerce missing word count to zero after binding", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 1114 });
      const widget = await bindFakeProject(page, "1,114 words");
      await setFakeWordCountUnavailable(page, "missing");

      await widget.getByRole("button", { name: "Sync document changes" }).click();

      await expect(widget).toContainText("Could not verify document changes right now.");
      await expect(widget).not.toContainText("-1,114");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-current-word-count", "1114");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");
    } finally {
      await context.close();
    }
  });

  test("manual sync accepts a verified zero on the same bound document", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 1114 });
      const widget = await bindFakeProject(page, "1,114 words");
      await setFakeWordCount(page, 0);

      await widget.getByRole("button", { name: "Sync document changes" }).click();

      await expect(widget).toContainText("Catch-up");
      await expect(widget).toContainText("-1114 words");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-current-word-count", "1114");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");
    } finally {
      await context.close();
    }
  });

  test("deleted bound document reopens as rebind required instead of active", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { documentId: "doc-deleted", title: "Deleted Document", words: 1114 });
      let widget = await bindFakeProject(page, "1,114 words");
      await setFakeWordCountUnavailable(page, "deleted");
      await widget.getByLabel("Close Scriptor controls").click();

      widget = await openControls(page);

      await expect(widget).toContainText("Bound document unavailable. Rebind required.");
      await expect(widget).toContainText("Not bound");
      await expect(widget).not.toContainText("Start writing");
      await expect(widget).not.toContainText("Sync document changes");
    } finally {
      await context.close();
    }
  });

  test("manual sync from Document B does not apply a delta to bound Document A", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { documentId: "doc-a", title: "Document A", words: 1114 });
      const widget = await bindFakeProject(page, "1,114 words");
      await setFakeDocument(page, "doc-b", "Document B", 500);

      await widget.getByRole("button", { name: "Sync document changes" }).click();

      await expect(widget).toContainText("Could not verify document changes right now.");
      await expect(widget).not.toContainText("-614");
      await expect(widget).not.toContainText("-1,114");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-current-word-count", "1114");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");
    } finally {
      await context.close();
    }
  });

  test("active typing does not trigger a random catch-up prompt", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 10800 });
      const widget = await bindFakeProject(page, "10,800 words");

      await setFakeWordCount(page, 10805, true);

      await expect(widget).not.toContainText("Catch-up");
      await widget.getByRole("button", { name: "Start writing" }).click();
      await expect(widget).toContainText("+5 words");
    } finally {
      await context.close();
    }
  });

  test("empty-start writing session uses positive API count when visible counter falsely stays zero", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 0 });
      await bindFakeProject(page, "0 words");
      const widget = await startWritingSession(page);
      await setFakeWordCount(page, 2024, true);
      await setFakeVisibleWordCount(page, 0);

      await widget.getByRole("button", { name: "End" }).click();

      await expect(widget).toContainText("Session saved", { timeout: 10000 });
      await expect(widget).toContainText("Net: +2024 words");
      await expect(widget).not.toContainText("Net: 0 words");
      await expect(widget).toContainText("Synced.");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-current-word-count", "2024");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "1");
    } finally {
      await context.close();
    }
  });

  test("active session remains tied to Document A when page identity switches", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 60500 });
      await bindFakeProject(page, "60,500 words");
      const widget = await startWritingSession(page);
      await setFakeDocument(page, "doc-b", "Document B", 500);

      await widget.getByRole("button", { name: "End" }).click();

      await expect(widget).toContainText("This session belongs to another Google Docs tab");
      await expect(widget).not.toContainText("-60,000");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");
    } finally {
      await context.close();
    }
  });

  test("minimized timer keeps the active session running and restores without duplication", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 60500 });
      await bindFakeProject(page, "60,500 words");
      const widget = await startWritingSession(page);

      await expect(widget.getByRole("button", { name: "Minimize timer" })).toBeVisible();
      await widget.getByRole("button", { name: "Minimize", exact: true }).click();

      await expect(widget).toHaveClass(/ace-widget--active-minimized/);
      await expect(widget.getByLabel("Restore timer")).toBeVisible();
      await expect(widget.locator(".ace-minimized-indicator")).toBeVisible();
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");

      await page.waitForTimeout(1100);
      await openControls(page);

      await expect(widget).toContainText("Writing");
      await expect(widget.getByRole("button", { name: "End" })).toBeVisible();
      await expect(widget).not.toHaveClass(/ace-widget--active-minimized/);
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");

      await setFakeWordCount(page, 60620, true);
      await widget.getByRole("button", { name: "End" }).click();
      await expect(widget).toContainText("Net: +120 words", { timeout: 10000 });
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "1");
    } finally {
      await context.close();
    }
  });

  test("minimized indicator is cleared when a restored session ends", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 60500 });
      await bindFakeProject(page, "60,500 words");
      const widget = await startWritingSession(page);

      await widget.getByRole("button", { name: "Minimize", exact: true }).click();
      await expect(widget.locator(".ace-minimized-indicator")).toBeVisible();

      await openControls(page);
      await widget.getByRole("button", { name: "End" }).click();
      await expect(widget).toContainText("Session saved", { timeout: 10000 });
      await widget.getByRole("button", { name: "Close Scriptor controls" }).click();

      await expect(widget).toHaveClass(/ace-widget--idle/);
      await expect(widget).not.toHaveClass(/ace-widget--active-minimized/);
      await expect(widget.locator(".ace-minimized-indicator")).toHaveCount(0);
      await expect(widget.getByLabel("Open Scriptor session controls")).toBeVisible();
    } finally {
      await context.close();
    }
  });

  test("valid negative writing session is allowed on the same document", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 60500 });
      await bindFakeProject(page, "60,500 words");
      const widget = await startWritingSession(page);
      await setFakeWordCount(page, 60000, true);

      await widget.getByRole("button", { name: "End" }).click();

      await expect(widget).toContainText("Session saved", { timeout: 10000 });
      await expect(widget).toContainText("Net: -500 words");
      await expect(widget).toContainText("Synced.");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-project-current-word-count", "60000");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "1");
    } finally {
      await context.close();
    }
  });

  test("impossible wrong-document delta is blocked", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { words: 60500 });
      await bindFakeProject(page, "60,500 words");
      const widget = await startWritingSession(page);
      await setFakeDocument(page, "doc-b", "Document B", 500);

      await widget.getByRole("button", { name: "End" }).click();

      await expect(widget).toContainText("Return to");
      await expect(widget).not.toContainText("-60,000");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");
    } finally {
      await context.close();
    }
  });

  test("tab switching does not let Document B finalize Document A session", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const pageA = await openFakeDoc(context, { documentId: "doc-a", title: "Document A", words: 60500 });
      await bindFakeProject(pageA, "60,500 words");
      const widgetA = await startWritingSession(pageA);

      const pageB = await openFakeDoc(context, { documentId: "doc-b", title: "Document B", words: 500 });
      const widgetB = await openControls(pageB);
      await expect(widgetB).not.toContainText("Session saved");
      await expect(pageB.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");

      await pageA.bringToFront();
      await expect(widgetA).toContainText("Writing");
      await setFakeWordCount(pageA, 60000, true);
      await widgetA.getByRole("button", { name: "End" }).click();
      await expect(widgetA).toContainText("Net: -500 words");
    } finally {
      await context.close();
    }
  });

  test("minimized session opened after switching to Document B prompts return to Document A", async ({}, testInfo) => {
    const context = await launchExtensionContext(testInfo);
    try {
      const page = await openFakeDoc(context, { documentId: "doc-a", title: "Document A", words: 60500 });
      await bindFakeProject(page, "60,500 words");
      const widget = await startWritingSession(page);
      await widget.getByRole("button", { name: "Minimize", exact: true }).click();
      await expect(widget.getByLabel("Restore timer")).toBeVisible();
      await setFakeDocument(page, "doc-b", "Document B", 500);

      await openControls(page);

      await expect(widget).toContainText("Return to");
      await expect(widget).not.toContainText("Session saved");
      await expect(page.getByTestId("fake-doc")).toHaveAttribute("data-session-count", "0");
    } finally {
      await context.close();
    }
  });
});
