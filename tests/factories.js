function documentFactory(overrides = {}) {
  return {
    documentId: "doc-test",
    title: "Test Google Doc",
    url: "https://docs.google.com/document/d/doc-test/edit",
    ...overrides
  };
}

function tabFactory(overrides = {}) {
  return {
    tabId: "tab-a",
    tabTitle: "Tab A",
    ...overrides
  };
}

function surfaceFactory(overrides = {}) {
  const document = documentFactory(overrides);
  const tab = tabFactory(overrides);
  const tabId = tab.tabId || "default";
  const manuscriptSurfaceId = overrides.manuscriptSurfaceId || `${document.documentId}:${tabId}`;
  return {
    documentId: document.documentId,
    tabId,
    tabTitle: tab.tabTitle || "",
    manuscriptSurfaceId,
    manuscriptSurfaceLabel: tab.tabTitle || "Current manuscript",
    ...overrides
  };
}

function projectFactory(overrides = {}) {
  return {
    id: "project-a",
    bookTitle: "Project A",
    title: "Project A",
    currentWordCount: 1000,
    ...overrides
  };
}

function bindingFactory(overrides = {}) {
  const surface = surfaceFactory(overrides);
  const project = overrides.project || projectFactory(overrides.projectOverrides || {});
  return {
    ...surface,
    projectId: project.id,
    project,
    updatedAt: "2026-05-20T00:00:00.000Z",
    ...overrides
  };
}

function baselineFactory(count = 1000, overrides = {}) {
  const surface = surfaceFactory(overrides);
  const project = overrides.project || projectFactory(overrides.projectOverrides || {});
  return {
    ...surface,
    projectId: project.id,
    project,
    endDocumentWordCount: count,
    wordCountMethod: "stable-visible",
    revisionId: "baseline-revision",
    syncedAt: "2026-05-20T00:00:00.000Z",
    sessionId: "",
    ...overrides
  };
}

function sessionFactory(overrides = {}) {
  const surface = surfaceFactory(overrides);
  const project = overrides.project || projectFactory(overrides.projectOverrides || {});
  const startDocumentWordCount = Object.prototype.hasOwnProperty.call(overrides, "startDocumentWordCount")
    ? overrides.startDocumentWordCount
    : 1000;
  const endDocumentWordCount = Object.prototype.hasOwnProperty.call(overrides, "endDocumentWordCount")
    ? overrides.endDocumentWordCount
    : 1200;
  return {
    ...surface,
    projectId: project.id,
    project,
    sessionType: "writing",
    startedAt: "2026-05-20T00:00:00.000Z",
    endedAt: "2026-05-20T00:01:00.000Z",
    durationMinutes: 1,
    source: "chrome-extension",
    documentUrl: `https://docs.google.com/document/d/${surface.documentId}/edit`,
    extensionSessionId: "session-a",
    startDocumentWordCount,
    endDocumentWordCount,
    netWordsChanged: endDocumentWordCount - startDocumentWordCount,
    wordCountMethod: "stable-visible-count",
    measurementPending: false,
    ...overrides
  };
}

function pendingSessionFactory(overrides = {}) {
  return sessionFactory({
    extensionSessionId: "pending-session-a",
    source: "catch-up",
    measurementPending: false,
    ...overrides
  });
}

function issueFactory(overrides = {}) {
  const surface = surfaceFactory(overrides);
  const project = overrides.project || projectFactory(overrides.projectOverrides || {});
  return {
    ...surface,
    projectId: project.id,
    id: "issue-a",
    extensionIssueId: "issue-a",
    title: "Issue A",
    note: "Tighten this paragraph.",
    snippet: "Selected text",
    status: "open",
    priority: "medium",
    ...overrides
  };
}

function snapshotFactory(wordCount, overrides = {}) {
  return {
    ok: Number.isFinite(Number(wordCount)),
    wordCount: Number.isFinite(Number(wordCount)) ? Number(wordCount) : null,
    apiWordCount: Number.isFinite(Number(wordCount)) ? Number(wordCount) : null,
    visibleWordCount: Number.isFinite(Number(wordCount)) ? Number(wordCount) : null,
    currentCountSource: "stable-visible",
    currentCountTrusted: Number.isFinite(Number(wordCount)),
    revisionId: "current-revision",
    ...overrides
  };
}

function locationForSurface(surface) {
  const resolved = surfaceFactory(surface);
  const hash = `#tab=${encodeURIComponent(resolved.tabId)}&tabTitle=${encodeURIComponent(resolved.tabTitle)}`;
  return {
    href: `https://docs.google.com/document/d/${resolved.documentId}/edit${hash}`,
    pathname: `/document/d/${resolved.documentId}/edit`,
    search: "",
    hash
  };
}

function createBackendMock(routes = {}) {
  const calls = [];
  const sendMessage = function (message, callback) {
    calls.push(message);
    if (message.aceType !== "ace-api-fetch") {
      callback?.({
        ok: false,
        wordCount: null,
        error: "No mocked Google Docs API response queued."
      });
      return;
    }

    const method = String(message.options?.method || "GET").toUpperCase();
    const routeKey = `${method} ${message.path}`;
    const basePath = String(message.path || "").split("?")[0];
    const baseKey = `${method} ${basePath}`;
    const handler = routes[routeKey] || routes[baseKey];
    if (!handler) {
      callback?.({
        ok: false,
        error: `No mocked backend route for ${routeKey}.`
      });
      return;
    }

    Promise.resolve(typeof handler === "function" ? handler(message) : handler)
      .then((payload) => {
        callback?.({
          ok: payload?.ok !== false,
          payload: payload?.payload !== undefined ? payload.payload : payload
        });
      })
      .catch((error) => {
        callback?.({
          ok: false,
          error: error.message || String(error)
        });
      });
  };

  return { calls, sendMessage };
}

module.exports = {
  baselineFactory,
  bindingFactory,
  createBackendMock,
  documentFactory,
  issueFactory,
  locationForSurface,
  pendingSessionFactory,
  projectFactory,
  sessionFactory,
  snapshotFactory,
  surfaceFactory,
  tabFactory
};
