const assert = require("node:assert/strict");
const test = require("node:test");
const { loadBackground, loadContent } = require("./helpers");
const {
  baselineFactory,
  bindingFactory,
  createBackendMock,
  issueFactory,
  locationForSurface,
  pendingSessionFactory,
  projectFactory,
  sessionFactory,
  snapshotFactory,
  surfaceFactory
} = require("./factories");

function evaluateCatchUp(exports, overrides = {}) {
  const surface = overrides.surface || surfaceFactory();
  const baseline = Object.prototype.hasOwnProperty.call(overrides, "baseline")
    ? overrides.baseline
    : baselineFactory(1000, surface);
  return exports.aceEvaluateCatchUpCandidate({
    trigger: "qa-harness",
    surface,
    baseline,
    baselineKey: overrides.baselineKey || surface.manuscriptSurfaceId,
    baselineIsLegacy: Boolean(overrides.baselineIsLegacy),
    currentSnapshot: Object.prototype.hasOwnProperty.call(overrides, "currentSnapshot")
      ? overrides.currentSnapshot
      : snapshotFactory(1200),
    binding: Object.prototype.hasOwnProperty.call(overrides, "binding")
      ? overrides.binding
      : bindingFactory(surface),
    pendingSession: overrides.pendingSession || null,
    completedSession: overrides.completedSession || null
  });
}

function bindingProjectItem(surface, project, overrides = {}) {
  return {
    project,
    isBound: true,
    binding: bindingFactory({ ...surface, project }),
    bindingStatus: "active",
    staleReason: "",
    ...overrides
  };
}

test("QA harness factories build consistent manuscript surfaces", () => {
  const surface = surfaceFactory({
    documentId: "doc-1",
    tabId: "tab-7",
    tabTitle: "Chapter 7"
  });
  const binding = bindingFactory({ ...surface, project: projectFactory({ id: "project-7" }) });
  const baseline = baselineFactory(777, surface);
  const session = sessionFactory({ ...surface, startDocumentWordCount: 700, endDocumentWordCount: 777 });
  const issue = issueFactory(surface);

  assert.equal(surface.manuscriptSurfaceId, "doc-1:tab-7");
  assert.equal(binding.manuscriptSurfaceId, surface.manuscriptSurfaceId);
  assert.equal(binding.projectId, "project-7");
  assert.equal(baseline.endDocumentWordCount, 777);
  assert.equal(session.netWordsChanged, 77);
  assert.equal(issue.tabTitle, "Chapter 7");
});

test("background app bridge forwards Scriptor session cookie as extension auth header", async () => {
  const { exports, fetchCalls } = loadBackground([], {
    sessionCookie: "signed-session-cookie",
    defaultFetchPayload: { projects: [] }
  });

  const response = await exports.aceFetchFromApp({
    path: "/api/extension/projects",
    options: { method: "GET" }
  });

  assert.equal(response.ok, true);
  assert.equal(fetchCalls[0].options.credentials, "include");
  assert.equal(fetchCalls[0].options.headers["X-Scriptor-Session"], "signed-session-cookie");
});

test("background app bridge falls back to domain cookie lookup when URL lookup misses", async () => {
  const { exports, fetchCalls } = loadBackground([], {
    sessionCookies: [{ domain: "davishedrick.pythonanywhere.com", value: "domain-session-cookie" }],
    defaultFetchPayload: { projects: [] }
  });

  const response = await exports.aceFetchFromApp({
    path: "/api/extension/projects",
    options: { method: "GET" }
  });

  assert.equal(response.ok, true);
  assert.equal(fetchCalls[0].options.headers["X-Scriptor-Session"], "domain-session-cookie");
});

test("project fetch recovers when runtime bridge auth misses but direct app session works", async () => {
  const project = projectFactory({ id: "project-auth", bookTitle: "Auth Project" });
  const { exports, fetchCalls } = loadContent({
    sendMessage(message, callback) {
      if (message.aceType === "ace-api-fetch" && message.path === "/api/extension/projects") {
        callback({
          ok: false,
          status: 401,
          payload: { error: "Authentication required." },
          error: "Authentication required."
        });
        return;
      }
      callback({ ok: false, error: "Unexpected message." });
    },
    fetch: async () => ({
      ok: true,
      status: 200,
      text: async () => JSON.stringify({ projects: [project] })
    })
  });

  const projects = await exports.aceGetExtensionProjects();

  assert.equal(projects.length, 1);
  assert.equal(projects[0].id, "project-auth");
  assert.equal(fetchCalls[0].url, "https://davishedrick.pythonanywhere.com/api/extension/projects");
  assert.equal(fetchCalls[0].options.credentials, "include");
});

test("background can focus owning Chrome tab and request session end", async () => {
  const tabMessages = [];
  const tab = { id: 101, windowId: 7, active: false };
  const windowRecord = { id: 7, focused: false };
  const { exports } = loadBackground([], {
    tabs: { 101: tab },
    windows: { 7: windowRecord },
    sendTabMessage(tabId, message, callback) {
      tabMessages.push({ tabId, message });
      callback({ ok: true });
    }
  });

  const response = await exports.aceEndSessionInChromeTab({
    chromeTabId: "101",
    extensionSessionId: "session-a"
  });

  assert.equal(response.ok, true);
  assert.equal(tab.active, true);
  assert.equal(windowRecord.focused, true);
  assert.equal(tabMessages[0].tabId, 101);
  assert.equal(tabMessages[0].message.aceType, "ace-end-active-session");
  assert.equal(tabMessages[0].message.extensionSessionId, "session-a");
});

test("background reports missing owning Chrome tab instead of silently failing", async () => {
  const { exports } = loadBackground([], { tabs: {} });

  await assert.rejects(
    () => exports.aceFocusChromeTab({ chromeTabId: "404" }),
    /Original writing tab is no longer available/
  );
});

test("one manuscript tab stores only one local project binding", async () => {
  const { exports, storage } = loadContent();
  const surface = surfaceFactory({ documentId: "doc-1", tabId: "tab-a" });

  await exports.aceSaveLocalDocumentBinding(surface, projectFactory({ id: "project-a" }));
  await exports.aceSaveLocalDocumentBinding(surface, projectFactory({ id: "project-b" }));

  assert.equal(storage.aceDocumentBindings["doc-1:tab-a"].projectId, "project-b");
  assert.equal(Object.keys(storage.aceDocumentBindings).length, 1);
});

test("one project cannot remain locally bound to multiple manuscript tabs", async () => {
  const { exports, storage } = loadContent();
  const tabA = surfaceFactory({ documentId: "doc-1", tabId: "tab-a", tabTitle: "Tab A" });
  const tabB = surfaceFactory({ documentId: "doc-1", tabId: "tab-b", tabTitle: "Tab B" });
  const project = projectFactory({ id: "project-a" });

  await exports.aceSaveLocalDocumentBinding(tabA, project);
  await exports.aceSaveLocalDocumentBinding(tabB, project);

  assert.equal(storage.aceDocumentBindings["doc-1:tab-a"], undefined);
  assert.equal(storage.aceDocumentBindings["doc-1:tab-b"].projectId, "project-a");
});

test("server-deleted binding clears stale local binding during reconciliation", async () => {
  const surface = surfaceFactory({ documentId: "doc-1", tabId: "tab-a" });
  const backend = createBackendMock({
    "GET /api/extension/document-binding": { project: null }
  });
  const { exports, storage } = loadContent({
    sendMessage: backend.sendMessage,
    initialStorage: {
      aceDocumentBindings: {
        [surface.manuscriptSurfaceId]: bindingFactory(surface)
      }
    }
  });

  const binding = await exports.aceGetBoundProjectForDocument(surface);

  assert.equal(binding, null);
  assert.equal(storage.aceDocumentBindings[surface.manuscriptSurfaceId], undefined);
});

test("server binding reconciliation replaces stale local project metadata", async () => {
  const surface = surfaceFactory({ documentId: "doc-1", tabId: "tab-a" });
  const serverProject = projectFactory({ id: "project-server", bookTitle: "Server Project" });
  const backend = createBackendMock({
    "GET /api/extension/document-binding": { project: serverProject }
  });
  const { exports, storage } = loadContent({
    sendMessage: backend.sendMessage,
    initialStorage: {
      aceDocumentBindings: {
        [surface.manuscriptSurfaceId]: bindingFactory({
          ...surface,
          project: projectFactory({ id: "project-stale", bookTitle: "Stale Project" })
        })
      }
    }
  });

  const binding = await exports.aceGetBoundProjectForDocument(surface);

  assert.equal(binding.projectId, "project-server");
  assert.equal(storage.aceDocumentBindings[surface.manuscriptSurfaceId].projectId, "project-server");
});

test("project picker marks deleted document binding as unbound with deleted marker", async () => {
  const surface = surfaceFactory({ documentId: "deleted-doc", tabId: "tab-a", tabTitle: "Deleted Tab" });
  const project = projectFactory({ id: "project-deleted", bookTitle: "Deleted Project" });
  const routeCalls = [];
  const { exports } = loadContent({
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({
          ok: false,
          status: 404,
          wordCount: null,
          error: "E-GOOGLE-API-404: Requested entity was not found."
        });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        routeCalls.push(message);
        callback({ ok: true, payload: { binding: { status: "stale_missing_doc" } } });
      }
    }
  });

  const projects = await exports.aceReconcileProjectPickerBindings([
    bindingProjectItem(surface, project)
  ]);

  assert.equal(projects[0].bindingStatus, "stale_missing_doc");
  assert.equal(projects[0].isBound, false);
  assert.equal(projects[0].binding, null);
  assert.equal(projects[0].deletedBinding.documentId, "deleted-doc");
  assert.match(exports.aceProjectPickerStatusLabel(projects[0]), /missing doc/);
  assert.equal(routeCalls[0].options.method, "PATCH");
});

test("project picker does not mark binding stale on validation network failure", async () => {
  const surface = surfaceFactory({ documentId: "network-doc", tabId: "tab-a" });
  const project = projectFactory({ id: "project-network" });
  const { exports } = loadContent({
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({
          ok: false,
          status: 0,
          wordCount: null,
          error: "Network unavailable."
        });
      }
    }
  });

  const projects = await exports.aceReconcileProjectPickerBindings([
    bindingProjectItem(surface, project)
  ]);

  assert.equal(projects[0].bindingStatus, "active");
  assert.equal(projects[0].staleReason, "");
});

test("project picker classifies missing tab as stale_missing_tab", async () => {
  const surface = surfaceFactory({ documentId: "doc-1", tabId: "missing-tab" });
  const project = projectFactory({ id: "project-tab" });
  const { exports } = loadContent({
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({
          ok: false,
          status: 404,
          wordCount: null,
          error: "E-GOOGLE-DOC-TAB-NOT-FOUND: Active Google Docs tab was not found in the API payload."
        });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        callback({ ok: true, payload: { binding: { status: "stale_missing_tab" } } });
      }
    }
  });

  const projects = await exports.aceReconcileProjectPickerBindings([
    bindingProjectItem(surface, project)
  ]);

  assert.equal(projects[0].bindingStatus, "stale_missing_tab");
  assert.match(exports.aceProjectPickerStatusLabel(projects[0]), /missing tab/);
});

test("project picker classifies forbidden doc as stale_inaccessible", async () => {
  const surface = surfaceFactory({ documentId: "forbidden-doc", tabId: "tab-a" });
  const project = projectFactory({ id: "project-forbidden" });
  const { exports } = loadContent({
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({
          ok: false,
          status: 403,
          wordCount: null,
          error: "E-GOOGLE-API-403: The caller does not have permission."
        });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        callback({ ok: true, payload: { binding: { status: "stale_inaccessible" } } });
      }
    }
  });

  const projects = await exports.aceReconcileProjectPickerBindings([
    bindingProjectItem(surface, project)
  ]);

  assert.equal(projects[0].bindingStatus, "stale_inaccessible");
  assert.match(exports.aceProjectPickerStatusLabel(projects[0]), /inaccessible/);
});

test("clearing stale binding removes backend and local binding without touching history", async () => {
  const surface = surfaceFactory({ documentId: "deleted-doc", tabId: "tab-a" });
  const project = projectFactory({ id: "project-history" });
  const staleItem = bindingProjectItem(surface, project, { bindingStatus: "stale_missing_doc" });
  const backend = createBackendMock({
    "DELETE /api/extension/document-binding": { removed: true }
  });
  const { exports, storage } = loadContent({
    sendMessage: backend.sendMessage,
    initialStorage: {
      aceDocumentBindings: {
        [surface.manuscriptSurfaceId]: bindingFactory({ ...surface, project })
      },
      acePendingSessions: [
        pendingSessionFactory({ ...surface, extensionSessionId: "history-session" })
      ]
    }
  });

  await exports.aceClearStaleProjectBinding(staleItem);

  assert.equal(storage.aceDocumentBindings[surface.manuscriptSurfaceId], undefined);
  assert.equal(storage.acePendingSessions.length, 1);
  assert.equal(backend.calls[0].options.method, "DELETE");
});

test("deleted binding project asks before rebinding and yes refreshes baseline", async () => {
  const deletedSurface = surfaceFactory({ documentId: "deleted-doc", tabId: "tab-a", tabTitle: "Deleted Tab" });
  const currentSurface = surfaceFactory({ documentId: "new-doc", tabId: "tab-b", tabTitle: "New Tab" });
  const project = projectFactory({ id: "project-rebind", bookTitle: "Rebind Project" });
  const backend = createBackendMock({
    "PUT /api/extension/document-binding": { project }
  });
  const { exports, storage } = loadContent({
    topFrame: true,
    visibleWordCount: 397,
    location: locationForSurface(currentSurface),
    sendMessage: backend.sendMessage
  });
  await new Promise((resolve) => setImmediate(resolve));
  exports.aceSetTestState({
    state: "project-picker",
    projectPickerMode: "bind",
    currentSurface,
    projects: [
      {
        project,
        isBound: false,
        bindingStatus: "stale_missing_doc",
        binding: null,
        deletedBinding: {
          ...deletedSurface,
          projectId: project.id,
          title: "Deleted Tab",
          url: "https://docs.google.com/document/d/deleted-doc/edit",
          status: "stale_missing_doc",
          staleReason: "404"
        }
      }
    ]
  });

  await exports.aceChooseProject(project.id);

  let state = exports.aceTestState();
  assert.equal(state.state, "deleted-binding-rebind-confirm");
  assert.match(
    state.widgetHtml,
    /This project was bound to a now-deleted file\. Update this project to your current tab\?/
  );
  assert.match(state.widgetHtml, />Yes</);
  assert.match(state.widgetHtml, />No</);

  await exports.aceConfirmDeletedBindingRebind();

  state = exports.aceTestState();
  const putCall = backend.calls.find((call) => call.options?.method === "PUT");
  const payload = JSON.parse(putCall.options.body);
  assert.equal(payload.documentId, "new-doc");
  assert.equal(payload.manuscriptSurfaceId, "new-doc:tab-b");
  assert.equal(storage.aceDocumentBaselines[currentSurface.manuscriptSurfaceId].endDocumentWordCount, 397);
  assert.equal(state.state, "prompt");
  assert.equal(state.catchUpCandidate, null);
  assert.doesNotMatch(state.widgetHtml, /Catch-up/);
});

test("declining deleted binding rebind leaves project unbound without baseline", async () => {
  const deletedSurface = surfaceFactory({ documentId: "deleted-doc", tabId: "tab-a" });
  const currentSurface = surfaceFactory({ documentId: "new-doc", tabId: "tab-b" });
  const project = projectFactory({ id: "project-no-rebind", bookTitle: "No Rebind" });
  const backend = createBackendMock({
    "PUT /api/extension/document-binding": { project }
  });
  const { exports, storage } = loadContent({
    topFrame: true,
    visibleWordCount: 397,
    location: locationForSurface(currentSurface),
    sendMessage: backend.sendMessage
  });
  await new Promise((resolve) => setImmediate(resolve));
  exports.aceSetTestState({
    state: "project-picker",
    projectPickerMode: "bind",
    currentSurface,
    projects: [
      {
        project,
        isBound: false,
        bindingStatus: "stale_missing_doc",
        deletedBinding: {
          ...deletedSurface,
          projectId: project.id,
          status: "stale_missing_doc"
        }
      }
    ]
  });

  await exports.aceChooseProject(project.id);
  exports.aceCancelDeletedBindingRebind();

  const state = exports.aceTestState();
  assert.equal(state.state, "project-picker");
  assert.equal(storage.aceDocumentBaselines?.[currentSurface.manuscriptSurfaceId], undefined);
  assert.equal(backend.calls.some((call) => call.options?.method === "PUT"), false);
  assert.equal(state.catchUpCandidate, null);
});

test("old documentId-only binding fallback applies only to default tab", async () => {
  const { exports, storage } = loadContent();
  storage.aceDocumentBindings = {
    "doc-legacy": { projectId: "project-legacy", project: projectFactory({ id: "project-legacy" }) }
  };

  const defaultBinding = await exports.aceGetLocalDocumentBinding(surfaceFactory({
    documentId: "doc-legacy",
    tabId: "default",
    manuscriptSurfaceId: "doc-legacy:default"
  }));
  const tabBinding = await exports.aceGetLocalDocumentBinding(surfaceFactory({
    documentId: "doc-legacy",
    tabId: "tab-b",
    manuscriptSurfaceId: "doc-legacy:tab-b"
  }));

  assert.equal(defaultBinding.projectId, "project-legacy");
  assert.equal(tabBinding, null);
});

test("catch-up contract covers positive negative full deletion zero and missing counts", () => {
  const { exports } = loadContent({ topFrame: true });

  assert.equal(evaluateCatchUp(exports, {
    baseline: baselineFactory(1000),
    currentSnapshot: snapshotFactory(1200)
  }).candidate.netWordsChanged, 200);
  assert.equal(evaluateCatchUp(exports, {
    baseline: baselineFactory(1000),
    currentSnapshot: snapshotFactory(900)
  }).candidate.netWordsChanged, -100);
  assert.equal(evaluateCatchUp(exports, {
    baseline: baselineFactory(1000),
    currentSnapshot: snapshotFactory(0)
  }).candidate.netWordsChanged, -1000);
  assert.equal(evaluateCatchUp(exports, {
    baseline: baselineFactory(0),
    currentSnapshot: snapshotFactory(0)
  }).reason, "zero-net");
  assert.equal(evaluateCatchUp(exports, {
    baseline: baselineFactory(1000),
    currentSnapshot: snapshotFactory(null, {
      ok: false,
      wordCount: null,
      currentCountTrusted: false,
      skipCatchUp: true
    })
  }).reason, "current-count-diagnostic-only");
});

test("manual session start shows positive catch-up before starting", async () => {
  const surface = surfaceFactory({ documentId: "doc-catchup", tabId: "tab-a", tabTitle: "Tab A" });
  const project = projectFactory({ id: "project-catchup", bookTitle: "Catchup Project" });
  const { exports } = loadContent({
    topFrame: true,
    visibleWordCount: 10800,
    location: locationForSurface(surface),
    initialStorage: {
      aceDocumentBaselines: {
        [surface.manuscriptSurfaceId]: baselineFactory(10305, { ...surface, project })
      }
    },
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 10800, revisionId: "current" });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        callback({ ok: true, payload: { project } });
      }
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  exports.aceSetTestState({
    state: "prompt",
    currentSurface: surface,
    currentBinding: bindingFactory({ ...surface, project }),
    selectedProject: project
  });

  await exports.aceStartSession("writing");

  const state = exports.aceTestState();
  assert.equal(state.state, "catch-up");
  assert.equal(state.activeSession, null);
  assert.equal(state.catchUpCandidate.netWordsChanged, 495);
  assert.match(state.widgetHtml, /Net: \+495 words since last session\./);
  assert.match(state.widgetHtml, /Detected change: .*>\+495 words</);
});

test("page load does not show catch-up for existing bound surface", async () => {
  const surface = surfaceFactory({ documentId: "doc-page-load", tabId: "tab-a", tabTitle: "Tab A" });
  const project = projectFactory({ id: "project-page-load", bookTitle: "Page Load Project" });
  const { exports } = loadContent({
    topFrame: true,
    visibleWordCount: 500,
    location: locationForSurface(surface),
    initialStorage: {
      aceDocumentBindings: {
        [surface.manuscriptSurfaceId]: bindingFactory({ ...surface, project })
      },
      aceDocumentBaselines: {
        [surface.manuscriptSurfaceId]: baselineFactory(0, { ...surface, project })
      }
    },
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 500, revisionId: "current" });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        callback({ ok: true, payload: { project } });
      }
    }
  });

  await new Promise((resolve) => setTimeout(resolve, 900));

  const state = exports.aceTestState();
  assert.notEqual(state.state, "catch-up");
  assert.equal(state.catchUpCandidate, null);
});

test("auto-start activity does not show catch-up while writing", async () => {
  const surface = surfaceFactory({ documentId: "doc-delete", tabId: "tab-a", tabTitle: "Tab A" });
  const project = projectFactory({ id: "project-delete", bookTitle: "Delete Project" });
  const { exports } = loadContent({
    topFrame: true,
    visibleWordCount: 0,
    location: locationForSurface(surface),
    initialStorage: {
      aceDocumentBindings: {
        [surface.manuscriptSurfaceId]: bindingFactory({ ...surface, project })
      },
      aceDocumentBaselines: {
        [surface.manuscriptSurfaceId]: baselineFactory(2300, { ...surface, project })
      }
    },
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 0, revisionId: "empty" });
        return;
      }
      if (message.aceType === "ace-google-doc-start-snapshot") {
        callback({ ok: true, status: 200, wordCount: 0, revisionId: "empty" });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        callback({ ok: true, payload: { project } });
      }
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  exports.aceSetTestState({
    state: "idle",
    activeSession: null,
    completedSession: null,
    catchUpCandidate: null
  });

  await exports.aceAutoStartBoundSessionFromActivity();

  const state = exports.aceTestState();
  assert.notEqual(state.state, "catch-up");
  assert.equal(state.activeSession, null);
  assert.equal(state.catchUpCandidate, null);
});

test("typing suppression blocks ambient catch-up but explicit start still reconciles", () => {
  const surface = surfaceFactory({ documentId: "doc-typing", tabId: "tab-a", tabTitle: "Tab A" });
  const project = projectFactory({ id: "project-typing", bookTitle: "Typing Project" });
  const { exports } = loadContent({ topFrame: true });

  exports.aceRegisterWritingActivity();
  assert.equal(exports.aceIsTypingSuppressionActive(), true);

  const ambient = exports.aceEvaluateCatchUpCandidate({
    trigger: "auto-start-attempt",
    surface,
    baseline: baselineFactory(10800, { ...surface, project }),
    baselineKey: surface.manuscriptSurfaceId,
    currentSnapshot: snapshotFactory(10805),
    binding: bindingFactory({ ...surface, project }),
    promptSuppressionReason: "typing-suppressed"
  });
  const explicit = exports.aceEvaluateCatchUpCandidate({
    trigger: "pre-session",
    surface,
    baseline: baselineFactory(10800, { ...surface, project }),
    baselineKey: surface.manuscriptSurfaceId,
    currentSnapshot: snapshotFactory(10805),
    binding: bindingFactory({ ...surface, project })
  });

  assert.equal(ambient.candidate, null);
  assert.equal(ambient.reason, "typing-suppressed");
  assert.equal(explicit.candidate.netWordsChanged, 5);
});

test("manual sync shows catch-up reconciliation on demand", async () => {
  const surface = surfaceFactory({ documentId: "doc-manual-sync", tabId: "tab-a", tabTitle: "Tab A" });
  const project = projectFactory({ id: "project-manual-sync", bookTitle: "Manual Sync Project" });
  const { exports } = loadContent({
    topFrame: true,
    visibleWordCount: 600,
    location: locationForSurface(surface),
    initialStorage: {
      aceDocumentBindings: {
        [surface.manuscriptSurfaceId]: bindingFactory({ ...surface, project })
      },
      aceDocumentBaselines: {
        [surface.manuscriptSurfaceId]: baselineFactory(100, { ...surface, project })
      }
    },
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 600, revisionId: "manual" });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        callback({ ok: true, payload: { project } });
      }
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  exports.aceSetTestState({
    state: "prompt",
    currentSurface: surface,
    currentBinding: bindingFactory({ ...surface, project }),
    selectedProject: project
  });

  await exports.aceManualSyncDocumentChanges();

  const state = exports.aceTestState();
  assert.equal(state.state, "catch-up");
  assert.equal(state.catchUpCandidate.netWordsChanged, 500);
  assert.match(state.widgetHtml, /Net: \+500 words since last session/);
});

test("catch-up handled before session keeps tracked session net separate", async () => {
  const surface = surfaceFactory({ documentId: "doc-separate", tabId: "tab-a", tabTitle: "Tab A" });
  const project = projectFactory({ id: "project-separate", bookTitle: "Separate Project" });
  const { exports, storage } = loadContent({
    topFrame: true,
    visibleWordCount: 500,
    location: locationForSurface(surface),
    initialStorage: {
      aceDocumentBaselines: {
        [surface.manuscriptSurfaceId]: baselineFactory(0, { ...surface, project })
      }
    },
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 500, revisionId: "five-hundred" });
        return;
      }
      if (message.aceType === "ace-google-doc-start-snapshot") {
        callback({ ok: true, status: 200, wordCount: 500, revisionId: "five-hundred" });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        const method = message.options?.method || "GET";
        callback({ ok: true, payload: { project: method === "PUT" ? project : null } });
      }
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  const catchUp = evaluateCatchUp(exports, {
    surface,
    baseline: baselineFactory(0, { ...surface, project }),
    currentSnapshot: snapshotFactory(500),
    binding: bindingFactory({ ...surface, project })
  });
  const catchUpSession = exports.aceBuildCatchUpSession(catchUp.candidate, project.id);
  assert.equal(catchUpSession.netWordsChanged, 500);

  await exports.aceSaveSkippedCatchUpBaseline(catchUp.candidate);
  exports.aceSetTestState({
    state: "prompt",
    currentSurface: surface,
    currentBinding: bindingFactory({ ...surface, project }),
    selectedProject: project,
    catchUpCandidate: null
  });
  await exports.aceStartSession("writing");

  assert.equal(exports.aceTestState().activeSession.startDocumentWordCount, 500);
  await exports.aceSaveDocumentBaseline(sessionFactory({
    ...surface,
    project,
    startDocumentWordCount: 500,
    endDocumentWordCount: 700
  }), project);
  assert.equal(storage.aceDocumentBaselines[surface.manuscriptSurfaceId].endDocumentWordCount, 700);
});

test("logged full-deletion catch-up advances baseline and cannot repeat", async () => {
  const surface = surfaceFactory({ documentId: "doc-repeat", tabId: "tab-a", tabTitle: "Tab A" });
  const project = projectFactory({ id: "project-repeat", bookTitle: "Repeat Project" });
  const { exports, storage } = loadContent({ location: locationForSurface(surface) });
  const result = evaluateCatchUp(exports, {
    surface,
    baseline: baselineFactory(2300, { ...surface, project }),
    currentSnapshot: snapshotFactory(0),
    binding: bindingFactory({ ...surface, project })
  });
  const catchUpSession = exports.aceBuildCatchUpSession(result.candidate, project.id);

  await exports.aceSaveDocumentBaseline(catchUpSession, project);
  const next = evaluateCatchUp(exports, {
    surface,
    baseline: storage.aceDocumentBaselines[surface.manuscriptSurfaceId],
    currentSnapshot: snapshotFactory(0),
    binding: bindingFactory({ ...surface, project })
  });

  assert.equal(storage.aceDocumentBaselines[surface.manuscriptSurfaceId].endDocumentWordCount, 0);
  assert.equal(next.candidate, null);
  assert.equal(next.reason, "zero-net");
});

test("initial bind with existing text establishes baseline without catch-up", async () => {
  const surface = surfaceFactory({ documentId: "doc-bind-text", tabId: "tab-a", tabTitle: "Tab A" });
  const project = projectFactory({ id: "project-bind-text", bookTitle: "Bind Text", currentWordCount: 10000 });
  const calls = [];
  const { exports, storage } = loadContent({
    topFrame: true,
    visibleWordCount: 10305,
    location: locationForSurface(surface),
    sendMessage(message, callback) {
      calls.push(message);
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 10305, revisionId: "ten-three-oh-five" });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        const method = message.options?.method || "GET";
        callback({ ok: true, payload: { project: method === "PUT" ? project : null } });
      }
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  exports.aceSetTestState({
    state: "project-picker",
    currentSurface: surface,
    projects: [project]
  });

  await exports.aceBindCurrentSurfaceToProject(project.id);

  let state = exports.aceTestState();
  assert.equal(state.state, "prompt");
  assert.equal(state.catchUpCandidate, null);
  assert.equal(storage.aceDocumentBaselines[surface.manuscriptSurfaceId].endDocumentWordCount, 10305);
  assert.equal(storage.acePendingSessions, undefined);
  assert.match(state.widgetHtml, /Verified manuscript size: 10,305 words/);
  const bindingCall = calls.find((message) => (
    message.aceType === "ace-api-fetch"
    && message.path === "/api/extension/document-binding"
    && message.options?.method === "PUT"
  ));
  const bindingPayload = JSON.parse(bindingCall.options.body);
  assert.equal(bindingPayload.verifiedWordCount, 10305);
});

test("active project binding mismatch requires confirmation instead of overwriting", async () => {
  const originalSurface = surfaceFactory({ documentId: "doc-bound-a", tabId: "tab-a", tabTitle: "Document A" });
  const currentSurface = surfaceFactory({ documentId: "doc-bound-b", tabId: "tab-a", tabTitle: "Document B" });
  const project = projectFactory({ id: "project-bound", bookTitle: "Bound Project" });
  const calls = [];
  const { exports, storage } = loadContent({
    topFrame: true,
    visibleWordCount: 500,
    location: locationForSurface(currentSurface),
    initialStorage: {
      aceDocumentBindings: {
        [originalSurface.manuscriptSurfaceId]: bindingFactory({ ...originalSurface, project })
      }
    },
    sendMessage(message, callback) {
      calls.push(message);
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 500, revisionId: "current-doc-b" });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        callback({
          ok: false,
          status: 409,
          error: "Project is already bound."
        });
      }
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  exports.aceSetTestState({
    state: "project-picker",
    currentSurface,
    projects: [project]
  });

  await exports.aceBindCurrentSurfaceToProject(project.id);

  const state = exports.aceTestState();
  const bindingCall = calls.find((message) => (
    message.aceType === "ace-api-fetch"
    && message.path === "/api/extension/document-binding"
    && message.options?.method === "PUT"
  ));
  assert.equal(state.state, "prompt");
  assert.match(state.widgetHtml, /Project is already bound/);
  assert.equal(Boolean(bindingCall), true);
  assert.equal(storage.aceDocumentBindings[originalSurface.manuscriptSurfaceId].projectId, project.id);
  assert.equal(storage.aceDocumentBindings[currentSurface.manuscriptSurfaceId], undefined);
  assert.equal(storage.aceDocumentBaselines, undefined);
});

test("initial bind does not trust false visible zero when API sees document words", async () => {
  const surface = surfaceFactory({ documentId: "doc-bind-visible-zero", tabId: "tab-a", tabTitle: "Tab A" });
  const project = projectFactory({ id: "project-bind-zero-api", bookTitle: "Bind Zero API" });
  const calls = [];
  const { exports, storage } = loadContent({
    topFrame: true,
    visibleWordCount: 0,
    location: locationForSurface(surface),
    sendMessage(message, callback) {
      calls.push(message);
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 522, revisionId: "api-522" });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        const method = message.options?.method || "GET";
        callback({ ok: true, payload: { project: method === "PUT" ? project : null } });
      }
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  exports.aceSetTestState({
    state: "project-picker",
    currentSurface: surface,
    projects: [project]
  });

  await exports.aceBindCurrentSurfaceToProject(project.id);

  const state = exports.aceTestState();
  assert.equal(state.state, "prompt");
  assert.equal(state.catchUpCandidate, null);
  assert.equal(storage.aceDocumentBaselines[surface.manuscriptSurfaceId].endDocumentWordCount, 522);
  assert.match(state.widgetHtml, /Verified manuscript size: 522 words/);
  assert.doesNotMatch(state.widgetHtml, /Verified manuscript size: 0 words/);
  const bindingCall = calls.find((message) => (
    message.aceType === "ace-api-fetch"
    && message.path === "/api/extension/document-binding"
    && message.options?.method === "PUT"
  ));
  const bindingPayload = JSON.parse(bindingCall.options.body);
  assert.equal(bindingPayload.verifiedWordCount, 522);
});

test("initial bind with stale local zero baseline resets to verified document count", async () => {
  const surface = surfaceFactory({ documentId: "doc-bind-refresh", tabId: "tab-a", tabTitle: "Tab A" });
  const project = projectFactory({ id: "project-bind-refresh", bookTitle: "Bind Refresh" });
  const { exports, storage } = loadContent({
    topFrame: true,
    visibleWordCount: 397,
    location: locationForSurface(surface),
    initialStorage: {
      aceDocumentBaselines: {
        [surface.manuscriptSurfaceId]: baselineFactory(0, { ...surface, project })
      }
    },
    sendMessage(message, callback) {
      if (message.aceType === "ace-api-fetch") {
        const method = message.options?.method || "GET";
        callback({ ok: true, payload: { project: method === "PUT" ? project : null } });
      }
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  exports.aceSetTestState({
    state: "project-picker",
    currentSurface: surface,
    projects: [project]
  });

  await exports.aceBindCurrentSurfaceToProject(project.id);

  const state = exports.aceTestState();
  assert.equal(state.state, "prompt");
  assert.equal(state.catchUpCandidate, null);
  assert.equal(storage.aceDocumentBaselines[surface.manuscriptSurfaceId].endDocumentWordCount, 397);
  assert.match(state.widgetHtml, /Verified manuscript size: 397 words/);
});

test("binding an empty tab initializes a zero baseline", async () => {
  const surface = surfaceFactory({ documentId: "doc-bind-empty", tabId: "tab-a", tabTitle: "Tab A" });
  const project = projectFactory({ id: "project-bind-empty", bookTitle: "Bind Empty" });
  const { exports, storage } = loadContent({
    topFrame: true,
    visibleWordCount: 0,
    location: locationForSurface(surface),
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 0, revisionId: "empty" });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        callback({ ok: true, payload: { project } });
      }
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  exports.aceSetTestState({
    state: "project-picker",
    currentSurface: surface,
    projects: [project]
  });

  await exports.aceBindCurrentSurfaceToProject(project.id);

  assert.equal(storage.aceDocumentBaselines[surface.manuscriptSurfaceId].endDocumentWordCount, 0);
});

test("visible zero is valid but null visible/API count is unavailable", async () => {
  const { exports } = loadContent();
  const zero = await exports.aceGoogleDocWordCountAfterSettle("doc-test", baselineFactory(500), {
    settleDelayMs: 0,
    visibleDelayMs: 0,
    visibleTimeoutMs: 100,
    readVisibleCount: async () => ({ count: 0, candidates: [{ count: 0 }], selectedCandidate: { count: 0 } }),
    apiCall: async () => ({ ok: true, wordCount: 0 })
  });
  const missing = await exports.aceGoogleDocWordCountAfterSettle("doc-test", baselineFactory(500), {
    settleDelayMs: 0,
    visibleDelayMs: 0,
    visibleTimeoutMs: 0,
    readVisibleCount: async () => ({ count: null, candidates: [], selectedCandidate: null }),
    apiCall: async () => ({ ok: false, wordCount: null })
  });

  assert.equal(zero.ok, true);
  assert.equal(zero.wordCount, 0);
  assert.equal(zero.netWordsChanged, -500);
  assert.equal(missing.ok, false);
  assert.equal(missing.wordCount, null);
});

test("stable visible count wins when API is unavailable or mismatched", async () => {
  const { exports } = loadContent();
  const unavailable = await exports.aceGoogleDocNetAfterSave("doc-test", "session-a", 1000, "", true, {
    settleDelayMs: 0,
    visibleDelayMs: 0,
    visibleTimeoutMs: 100,
    readVisibleCount: async () => ({ count: 1200, candidates: [{ count: 1200 }], selectedCandidate: { count: 1200 } }),
    apiCall: async () => ({ ok: false, wordCount: null })
  });
  const mismatch = await exports.aceGoogleDocNetAfterSave("doc-test", "session-a", 1000, "", true, {
    settleDelayMs: 0,
    visibleDelayMs: 0,
    visibleTimeoutMs: 100,
    readVisibleCount: async () => ({ count: 1200, candidates: [{ count: 1200 }], selectedCandidate: { count: 1200 } }),
    apiCall: async () => ({ ok: true, wordCount: 1199 })
  });

  assert.equal(unavailable.ok, true);
  assert.equal(unavailable.wordCount, 1200);
  assert.equal(mismatch.ok, true);
  assert.equal(mismatch.wordCount, 1200);
  assert.match(mismatch.wordCountDiagnostic, /W-API-VISIBLE-MISMATCH|D-STABLE-VISIBLE-NET/);
});

test("pending sessions retry state is isolated by manuscript surface", async () => {
  const tabA = surfaceFactory({ documentId: "doc-1", tabId: "tab-a" });
  const tabB = surfaceFactory({ documentId: "doc-1", tabId: "tab-b" });
  const { exports, storage } = loadContent();

  await exports.aceStorePendingSession(pendingSessionFactory({ ...tabA, extensionSessionId: "pending-a" }));
  await exports.aceStorePendingSession(pendingSessionFactory({ ...tabB, extensionSessionId: "pending-b" }));

  const pending = await exports.acePendingSessions();
  assert.equal(pending.length, 2);
  assert.equal(pending.find((session) => session.extensionSessionId === "pending-a").manuscriptSurfaceId, "doc-1:tab-a");
  assert.equal(pending.find((session) => session.extensionSessionId === "pending-b").manuscriptSurfaceId, "doc-1:tab-b");
  assert.equal(storage.acePendingSessions.length, 2);
});

test("tab switching refreshes project baseline and catch-up state for the new tab", async () => {
  const tabA = surfaceFactory({ documentId: "doc-1", tabId: "tab-a", tabTitle: "Tab A" });
  const tabB = surfaceFactory({ documentId: "doc-1", tabId: "tab-b", tabTitle: "Tab B" });
  const projectB = projectFactory({ id: "project-b", bookTitle: "Project B" });
  const { exports, storage, context } = loadContent({
    topFrame: true,
    location: locationForSurface(tabA)
  });
  await new Promise((resolve) => setImmediate(resolve));
  storage.aceDocumentBindings = {
    [tabB.manuscriptSurfaceId]: bindingFactory({ ...tabB, project: projectB })
  };
  storage.aceDocumentBaselines = {
    [tabA.manuscriptSurfaceId]: baselineFactory(1000, tabA),
    [tabB.manuscriptSurfaceId]: baselineFactory(3000, { ...tabB, project: projectB })
  };

  context.location.hash = locationForSurface(tabB).hash;
  context.location.href = locationForSurface(tabB).href;
  await exports.aceHandleSurfaceLifecycleChange("qa-tab-switch");
  const baseline = await exports.aceGetDocumentBaselineForCatchUp(tabB);

  assert.equal(exports.aceTestState().currentBinding.projectId, "project-b");
  assert.equal(baseline.baseline.endDocumentWordCount, 3000);
  assert.equal(exports.aceTestState().catchUpCandidate, null);
});

test("active session blocks on tab switch and resumes only on original tab", async () => {
  const tabA = surfaceFactory({ documentId: "doc-1", tabId: "tab-a", tabTitle: "Tab A" });
  const tabB = surfaceFactory({ documentId: "doc-1", tabId: "tab-b", tabTitle: "Tab B" });
  const { exports, context } = loadContent({
    topFrame: true,
    location: locationForSurface(tabA)
  });
  await new Promise((resolve) => setImmediate(resolve));
  await exports.aceHandleSurfaceLifecycleChange("qa-init");
  exports.aceSetTestState({
    state: "active",
    activeSession: sessionFactory({
      ...tabA,
      startedAt: new Date(Date.now() - 30000).toISOString()
    })
  });

  context.location.hash = locationForSurface(tabB).hash;
  context.location.href = locationForSurface(tabB).href;
  await exports.aceHandleSurfaceLifecycleChange("qa-switch");
  assert.equal(exports.aceTestState().state, "tab-blocked");

  context.location.hash = locationForSurface(tabA).hash;
  context.location.href = locationForSurface(tabA).href;
  await exports.aceHandleSurfaceLifecycleChange("qa-return");
  assert.equal(exports.aceTestState().state, "active");
});

test("minimizing an active timer keeps the session running as UI-only state", async () => {
  const surface = surfaceFactory({ documentId: "doc-minimize", tabId: "tab-a", tabTitle: "Draft" });
  const { exports, storage } = loadContent({
    topFrame: true,
    location: locationForSurface(surface)
  });
  await new Promise((resolve) => setImmediate(resolve));
  const startedAt = new Date(Date.now() - 61000).toISOString();
  exports.aceSetTestState({
    state: "active",
    activeSession: sessionFactory({
      ...surface,
      startedAt,
      endedAt: "",
      startDocumentWordCount: 1000,
      endDocumentWordCount: null
    })
  });

  exports.aceRenderActive();
  const before = exports.aceTestState().activeSession;
  exports.aceMinimizeWidget();
  await new Promise((resolve) => setImmediate(resolve));

  const state = exports.aceTestState();
  assert.equal(state.state, "active-minimized");
  assert.equal(state.activeSession.extensionSessionId, before.extensionSessionId);
  assert.equal(state.activeSession.startedAt, startedAt);
  assert.equal(state.activeSession.timerDisplayMode, "minimized");
  assert.equal(state.activeSession.showMinimizedIndicator, true);
  assert.match(state.widgetClassName, /ace-widget--active-minimized/);
  assert.match(state.widgetHtml, /Restore timer/);
  assert.equal(storage.aceActiveSession.extensionSessionId, before.extensionSessionId);
  assert.equal(storage.aceActiveSession.timerDisplayMode, "minimized");
});

test("restoring a minimized timer reuses the active session without duplication", async () => {
  const surface = surfaceFactory({ documentId: "doc-restore", tabId: "tab-a", tabTitle: "Draft" });
  const project = projectFactory({ id: "project-restore" });
  const { exports, storage } = loadContent({
    topFrame: true,
    location: locationForSurface(surface)
  });
  await new Promise((resolve) => setImmediate(resolve));
  const activeSession = sessionFactory({
    ...surface,
    project,
    projectId: project.id,
    extensionSessionId: "session-restore",
    startedAt: new Date(Date.now() - 65000).toISOString(),
    endedAt: "",
    startDocumentWordCount: 1000,
    endDocumentWordCount: null,
    timerDisplayMode: "minimized",
    showMinimizedIndicator: true,
    sessionScope: exports.aceCreateSessionScope(surface, {
      projectId: project.id,
      chromeTabId: "1"
    }),
    chromeTabId: "1"
  });
  exports.aceSetTestState({
    state: "active-minimized",
    activeSession
  });
  exports.aceRenderIdle();

  await exports.aceShowControls();

  const state = exports.aceTestState();
  assert.equal(state.state, "active");
  assert.equal(state.activeSession.extensionSessionId, "session-restore");
  assert.equal(state.activeSession.startedAt, activeSession.startedAt);
  assert.equal(state.activeSession.timerDisplayMode, "expanded");
  assert.equal(state.activeSession.showMinimizedIndicator, false);
  assert.equal(state.completedSession, null);
  assert.equal(storage.acePendingSessions, undefined);
  assert.match(state.widgetHtml, /End/);
  assert.match(state.widgetHtml, /Minimize/);
});

test("minimized active indicator appears only for active minimized sessions", async () => {
  const surface = surfaceFactory({ documentId: "doc-indicator", tabId: "tab-a", tabTitle: "Draft" });
  const { exports } = loadContent({
    topFrame: true,
    location: locationForSurface(surface)
  });
  await new Promise((resolve) => setImmediate(resolve));

  exports.aceSetTestState({ state: "idle", activeSession: null, completedSession: null });
  exports.aceRenderIdle();
  assert.doesNotMatch(exports.aceTestState().widgetClassName, /ace-widget--active-minimized/);
  assert.doesNotMatch(exports.aceTestState().widgetHtml, /ace-minimized-indicator/);

  exports.aceSetTestState({
    state: "active",
    activeSession: sessionFactory({
      ...surface,
      endedAt: "",
      endDocumentWordCount: null,
      timerDisplayMode: "expanded",
      showMinimizedIndicator: false
    })
  });
  exports.aceRenderActive();
  assert.doesNotMatch(exports.aceTestState().widgetHtml, /ace-minimized-indicator/);

  exports.aceMinimizeWidget();
  assert.match(exports.aceTestState().widgetHtml, /ace-minimized-indicator/);

  exports.aceSetTestState({ state: "idle", activeSession: null });
  exports.aceRenderIdle();
  assert.doesNotMatch(exports.aceTestState().widgetHtml, /ace-minimized-indicator/);
});

test("restoring minimized timer on the wrong tab keeps the session blocked", async () => {
  const tabA = surfaceFactory({ documentId: "doc-min-wrong", tabId: "tab-a", tabTitle: "Tab A" });
  const tabB = surfaceFactory({ documentId: "doc-min-wrong", tabId: "tab-b", tabTitle: "Tab B" });
  const project = projectFactory({ id: "project-min-wrong" });
  const { exports, context } = loadContent({
    topFrame: true,
    location: locationForSurface(tabA)
  });
  await new Promise((resolve) => setImmediate(resolve));
  exports.aceSetTestState({
    state: "active-minimized",
    activeSession: sessionFactory({
      ...tabA,
      project,
      projectId: project.id,
      extensionSessionId: "session-min-wrong",
      startedAt: new Date(Date.now() - 60000).toISOString(),
      endedAt: "",
      startDocumentWordCount: 1000,
      endDocumentWordCount: null,
      timerDisplayMode: "minimized",
      showMinimizedIndicator: true,
      sessionScope: exports.aceCreateSessionScope(tabA, {
        projectId: project.id,
        chromeTabId: "1"
      }),
      chromeTabId: "1"
    })
  });

  context.location.hash = locationForSurface(tabB).hash;
  context.location.href = locationForSurface(tabB).href;
  await exports.aceShowControls();

  const state = exports.aceTestState();
  assert.equal(state.state, "tab-blocked");
  assert.equal(state.activeSession.extensionSessionId, "session-min-wrong");
  assert.equal(state.completedSession, null);
  assert.match(state.widgetHtml, /Return to "Tab A"|belongs to another Google Docs tab/);
});

test("same document and same Chrome tab can complete a negative writing session", async () => {
  const surface = surfaceFactory({ documentId: "doc-negative", tabId: "tab-a", tabTitle: "Draft" });
  const project = projectFactory({ id: "project-negative", bookTitle: "Negative Project" });
  const calls = [];
  const { exports } = loadContent({
    topFrame: true,
    chromeTabId: 101,
    visibleWordCount: 60000,
    location: locationForSurface(surface),
    sendMessage(message, callback) {
      calls.push(message);
      if (message.aceType === "ace-google-doc-net-count") {
        callback({
          ok: true,
          status: 200,
          wordCount: 60000,
          netWordsChanged: -500,
          revisionId: "end-revision",
          wordCountMethod: "google-docs-api"
        });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        callback({ ok: true, payload: { project, session: { netWordsChanged: -500 } } });
        return;
      }
      callback({ ok: false, wordCount: null, error: "unexpected message" });
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  const sessionScope = exports.aceCreateSessionScope(surface, {
    projectId: project.id,
    chromeTabId: "101"
  });
  exports.aceSetTestState({
    state: "active",
    activeSession: sessionFactory({
      ...surface,
      project,
      projectId: project.id,
      extensionSessionId: "session-negative",
      sessionType: "writing",
      startedAt: new Date(Date.now() - 60000).toISOString(),
      endedAt: "",
      startDocumentWordCount: 60500,
      endDocumentWordCount: null,
      sessionScope,
      chromeTabId: "101"
    })
  });

  await exports.aceEndSession();

  const state = exports.aceTestState();
  assert.equal(state.state, "completed");
  assert.equal(state.completedSession.netWordsChanged, -500);
  assert.equal(state.completedSession.endDocumentWordCount, 60000);
  assert.equal(state.activeSession, null);
  assert.equal(calls.some((message) => message.aceType === "ace-google-doc-net-count"), true);
});

test("different Google document blocks session completion before word count", async () => {
  const docA = surfaceFactory({ documentId: "doc-a", tabId: "tab-a", tabTitle: "Doc A" });
  const docB = surfaceFactory({ documentId: "doc-b", tabId: "tab-b", tabTitle: "Doc B" });
  const project = projectFactory({ id: "project-scope", bookTitle: "Scoped Project" });
  const calls = [];
  const { exports, storage } = loadContent({
    topFrame: true,
    chromeTabId: 202,
    visibleWordCount: 500,
    location: locationForSurface(docB),
    sendMessage(message, callback) {
      calls.push(message);
      callback({ ok: false, wordCount: null, error: "wrong scope should not measure" });
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  const activeSession = sessionFactory({
    ...docA,
    project,
    projectId: project.id,
    extensionSessionId: "session-doc-a",
    sessionType: "writing",
    startedAt: new Date(Date.now() - 60000).toISOString(),
    endedAt: "",
    startDocumentWordCount: 60500,
    endDocumentWordCount: null,
    sessionScope: exports.aceCreateSessionScope(docA, {
      projectId: project.id,
      chromeTabId: "101"
    }),
    chromeTabId: "101"
  });
  exports.aceSetTestState({
    state: "active",
    activeSession
  });

  const callsBeforeEnd = calls.length;
  await exports.aceEndSession();
  const completionCalls = calls.slice(callsBeforeEnd);

  const state = exports.aceTestState();
  assert.equal(state.state, "tab-blocked");
  assert.equal(state.activeSession.extensionSessionId, "session-doc-a");
  assert.equal(state.completedSession, null);
  assert.equal(storage.acePendingSessions, undefined);
  assert.equal(completionCalls.some((message) => message.aceType === "ace-google-doc-net-count"), false);
  assert.equal(completionCalls.some((message) => message.aceType === "ace-api-fetch"), false);
  assert.match(state.widgetHtml, /belongs to another Google Docs tab|Return to/);
});

test("same Google document in a different Chrome tab cannot complete the session", async () => {
  const surface = surfaceFactory({ documentId: "doc-shared", tabId: "tab-a", tabTitle: "Shared Doc" });
  const project = projectFactory({ id: "project-shared", bookTitle: "Shared Project" });
  const calls = [];
  const { exports } = loadContent({
    topFrame: true,
    chromeTabId: 303,
    visibleWordCount: 0,
    location: locationForSurface(surface),
    sendMessage(message, callback) {
      calls.push(message);
      callback({ ok: false, wordCount: null, error: "wrong Chrome tab should not measure" });
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  exports.aceSetTestState({
    state: "active",
    activeSession: sessionFactory({
      ...surface,
      project,
      projectId: project.id,
      extensionSessionId: "session-shared",
      startedAt: new Date(Date.now() - 60000).toISOString(),
      endedAt: "",
      startDocumentWordCount: 2300,
      endDocumentWordCount: null,
      sessionScope: exports.aceCreateSessionScope(surface, {
        projectId: project.id,
        chromeTabId: "101"
      }),
      chromeTabId: "101"
    })
  });

  const callsBeforeEnd = calls.length;
  await exports.aceEndSession();
  const completionCalls = calls.slice(callsBeforeEnd);

  const state = exports.aceTestState();
  assert.equal(state.state, "tab-blocked");
  assert.equal(state.completedSession, null);
  assert.equal(completionCalls.some((message) => message.aceType === "ace-google-doc-net-count"), false);
  assert.equal(completionCalls.some((message) => message.aceType === "ace-api-fetch"), false);
});

test("session scope validation reports document and Chrome tab mismatches", () => {
  const surface = surfaceFactory({ documentId: "doc-scope", tabId: "tab-a", tabTitle: "Scope Tab" });
  const project = projectFactory({ id: "project-scope-validation" });
  const { exports } = loadContent({ location: locationForSurface(surface) });
  const session = sessionFactory({
    ...surface,
    project,
    projectId: project.id,
    sessionScope: exports.aceCreateSessionScope(surface, {
      projectId: project.id,
      chromeTabId: "101"
    }),
    chromeTabId: "101"
  });

  const sameScope = exports.aceCreateSessionScope(surface, {
    projectId: project.id,
    chromeTabId: "101"
  });
  const otherDocScope = exports.aceCreateSessionScope(surfaceFactory({
    documentId: "doc-other",
    tabId: "tab-a",
    tabTitle: "Other"
  }), {
    projectId: project.id,
    chromeTabId: "101"
  });
  const otherChromeTabScope = exports.aceCreateSessionScope(surface, {
    projectId: project.id,
    chromeTabId: "202"
  });

  assert.equal(exports.aceValidateSessionScope(session, sameScope).ok, true);
  assert.equal(exports.aceValidateSessionScope(session, otherDocScope).reason, "document-mismatch");
  assert.equal(exports.aceValidateSessionScope(session, otherChromeTabScope).reason, "chrome-tab-mismatch");
});

test("reload with stored active session shows abandoned-session recovery modal", async () => {
  const surface = surfaceFactory({ documentId: "doc-reload", tabId: "tab-a", tabTitle: "Reload Tab" });
  const project = projectFactory({ id: "project-reload", bookTitle: "The Black Harbor", currentWordCount: 60000 });
  const activeSession = sessionFactory({
    ...surface,
    project,
    projectId: project.id,
    extensionSessionId: "session-reload",
    startedAt: new Date(Date.now() - 43 * 60000).toISOString(),
    endedAt: "",
    startDocumentWordCount: 60000,
    endDocumentWordCount: null
  });
  const { exports, storage } = loadContent({
    topFrame: true,
    visibleWordCount: 60812,
    location: locationForSurface(surface),
    initialStorage: {
      aceActiveSession: activeSession
    },
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 60812, revisionId: "recovery" });
        return;
      }
      callback({ ok: true, payload: { project } });
    }
  });

  await new Promise((resolve) => setTimeout(resolve, 900));

  const state = exports.aceTestState();
  assert.equal(state.state, "recovery");
  assert.equal(state.activeSession, null);
  assert.equal(state.recoveryCandidate.netWordsChanged, 812);
  assert.equal(storage.aceAbandonedSessions[0].extensionSessionId, "session-reload");
  assert.match(state.widgetHtml, /You closed your last session without saving\./);
  assert.match(state.widgetHtml, /The Black Harbor/);
  assert.match(state.widgetHtml, /Word change/);
  assert.match(state.widgetHtml, /\+812 words/);
  assert.match(state.widgetHtml, /Recover Session/);
  assert.match(state.widgetHtml, /Discard/);
});

test("recovery blocks stale start count that conflicts with saved baseline", async () => {
  const surface = surfaceFactory({ documentId: "doc-stale-recovery", tabId: "tab-a", tabTitle: "Version 7" });
  const project = projectFactory({ id: "project-stale-recovery", bookTitle: "Hollowfield v7", currentWordCount: 63000 });
  const staleActiveSession = sessionFactory({
    ...surface,
    project,
    projectId: project.id,
    extensionSessionId: "session-stale-recovery",
    startedAt: new Date(Date.now() - 1 * 60000).toISOString(),
    endedAt: "",
    startDocumentWordCount: 1081,
    endDocumentWordCount: null
  });
  const { exports } = loadContent({
    topFrame: true,
    visibleWordCount: 60240,
    location: locationForSurface(surface),
    initialStorage: {
      aceActiveSession: staleActiveSession,
      aceDocumentBaselines: {
        [surface.manuscriptSurfaceId]: baselineFactory(63000, { ...surface, project })
      }
    },
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 60240, revisionId: "current-60240" });
        return;
      }
      callback({ ok: true, payload: { project } });
    }
  });
  await new Promise((resolve) => setTimeout(resolve, 900));

  const state = exports.aceTestState();
  assert.equal(state.state, "recovery");
  assert.equal(state.recoveryCandidate.measurementPending, true);
  assert.notEqual(state.recoveryCandidate.netWordsChanged, 59159);
  assert.match(state.recoveryCandidate.currentSnapshot.wordCountDiagnostic, /E-RECOVERY-START-BASELINE-MISMATCH/);
  assert.doesNotMatch(state.widgetHtml, /\+59159 words/);

  await exports.aceRecoverAbandonedSession();

  const recoveredState = exports.aceTestState();
  assert.equal(recoveredState.state, "recovery");
  assert.equal(recoveredState.completedSession, null);
  assert.match(recoveredState.widgetHtml, /Stored session start count conflicts/);
});

test("recovering abandoned session syncs once and advances baseline", async () => {
  const surface = surfaceFactory({ documentId: "doc-recover", tabId: "tab-a", tabTitle: "Recover Tab" });
  const project = projectFactory({ id: "project-recover", bookTitle: "Recover Project", currentWordCount: 60000 });
  const abandonedSession = sessionFactory({
    ...surface,
    project,
    projectId: project.id,
    extensionSessionId: "session-recover",
    startedAt: new Date(Date.now() - 20 * 60000).toISOString(),
    endedAt: "",
    startDocumentWordCount: 60000,
    endDocumentWordCount: null
  });
  const calls = [];
  const { exports, storage } = loadContent({
    topFrame: true,
    visibleWordCount: 60800,
    location: locationForSurface(surface),
    initialStorage: {
      aceActiveSession: abandonedSession
    },
    sendMessage(message, callback) {
      calls.push(message);
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 60800, revisionId: "recover-end" });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        callback({
          ok: true,
          payload: {
            project,
            session: {
              extensionSessionId: "session-recover",
              netWordsChanged: 800
            }
          }
        });
      }
    }
  });
  await new Promise((resolve) => setTimeout(resolve, 900));

  await exports.aceRecoverAbandonedSession();

  const postCalls = calls.filter((message) => (
    message.aceType === "ace-api-fetch"
    && message.path === "/api/extension/sessions"
    && message.options?.method === "POST"
  ));
  const payload = JSON.parse(postCalls[0].options.body);
  assert.equal(postCalls.length, 1);
  assert.equal(payload.extensionSessionId, "session-recover");
  assert.equal(payload.startDocumentWordCount, 60000);
  assert.equal(payload.endDocumentWordCount, 60800);
  assert.equal(payload.netWordsChanged, 800);
  assert.equal(storage.aceAbandonedSessions.length, 0);
  assert.equal(storage.aceActiveSession, undefined);
  assert.equal(storage.aceDocumentBaselines[surface.manuscriptSurfaceId].endDocumentWordCount, 60800);
});

test("discarding abandoned session does not advance baseline and catch-up remains possible", async () => {
  const surface = surfaceFactory({ documentId: "doc-discard", tabId: "tab-a", tabTitle: "Discard Tab" });
  const project = projectFactory({ id: "project-discard", bookTitle: "Discard Project" });
  const activeSession = sessionFactory({
    ...surface,
    project,
    projectId: project.id,
    extensionSessionId: "session-discard",
    startedAt: new Date(Date.now() - 10 * 60000).toISOString(),
    endedAt: "",
    startDocumentWordCount: 60000,
    endDocumentWordCount: null
  });
  const { exports, storage } = loadContent({
    topFrame: true,
    visibleWordCount: 60800,
    location: locationForSurface(surface),
    initialStorage: {
      aceActiveSession: activeSession,
      aceDocumentBindings: {
        [surface.manuscriptSurfaceId]: bindingFactory({ ...surface, project })
      },
      aceDocumentBaselines: {
        [surface.manuscriptSurfaceId]: baselineFactory(60000, { ...surface, project })
      }
    },
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 60800, revisionId: "discard-current" });
        return;
      }
      if (message.aceType === "ace-api-fetch") {
        callback({ ok: true, payload: { project } });
      }
    }
  });
  await new Promise((resolve) => setTimeout(resolve, 900));

  await exports.aceDiscardAbandonedSession();

  assert.equal(storage.aceAbandonedSessions.length, 0);
  assert.equal(storage.aceActiveSession, undefined);
  assert.equal(storage.aceDocumentBaselines[surface.manuscriptSurfaceId].endDocumentWordCount, 60000);

  const catchUpResult = await exports.aceBuildCatchUpCandidate(surface.documentId, "pre-session");
  assert.equal(catchUpResult.candidate.netWordsChanged, 800);
  assert.equal(catchUpResult.candidate.startDocumentWordCount, 60000);
  assert.equal(catchUpResult.candidate.endDocumentWordCount, 60800);
});

test("pending session suppresses duplicate abandoned-session recovery", async () => {
  const surface = surfaceFactory({ documentId: "doc-duplicate", tabId: "tab-a", tabTitle: "Duplicate Tab" });
  const project = projectFactory({ id: "project-duplicate", bookTitle: "Duplicate Project" });
  const activeSession = sessionFactory({
    ...surface,
    project,
    projectId: project.id,
    extensionSessionId: "session-duplicate",
    startedAt: new Date(Date.now() - 10 * 60000).toISOString(),
    endedAt: "",
    startDocumentWordCount: 1000,
    endDocumentWordCount: null
  });
  const pendingSession = pendingSessionFactory({
    ...surface,
    project,
    projectId: project.id,
    extensionSessionId: "session-duplicate",
    startDocumentWordCount: 1000,
    endDocumentWordCount: 1200,
    netWordsChanged: 200
  });
  const { exports } = loadContent({
    topFrame: true,
    visibleWordCount: 1200,
    location: locationForSurface(surface),
    initialStorage: {
      aceActiveSession: activeSession,
      acePendingSessions: [pendingSession],
      aceAbandonedSessions: [activeSession]
    },
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 1200, revisionId: "duplicate" });
        return;
      }
      callback({ ok: true, payload: { project } });
    }
  });
  await new Promise((resolve) => setImmediate(resolve));

  const state = exports.aceTestState();
  assert.equal(state.state, "completed");
  assert.equal(state.recoveryCandidate, null);
  assert.equal(state.completedSession.extensionSessionId, "session-duplicate");
});
