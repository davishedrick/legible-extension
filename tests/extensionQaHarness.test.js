const assert = require("node:assert/strict");
const test = require("node:test");
const { loadContent } = require("./helpers");
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

test("project picker marks deleted document binding as stale and shows clear action", async () => {
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
  assert.equal(projects[0].isBound, true);
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
  const { exports } = loadContent();

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
    visibleWordCount: 500,
    location: locationForSurface(surface),
    initialStorage: {
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
  assert.equal(state.catchUpCandidate.netWordsChanged, 500);
});

test("page load evaluates catch-up for existing bound surface", async () => {
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
  assert.equal(state.state, "catch-up");
  assert.equal(state.catchUpCandidate.netWordsChanged, 500);
});

test("auto-start is blocked by required negative catch-up", async () => {
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
  assert.equal(state.state, "catch-up");
  assert.equal(state.activeSession, null);
  assert.equal(state.catchUpCandidate.netWordsChanged, -2300);
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
        callback({ ok: true, payload: { project } });
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

test("binding an existing text tab initializes baseline to current count", async () => {
  const surface = surfaceFactory({ documentId: "doc-bind-text", tabId: "tab-a", tabTitle: "Tab A" });
  const project = projectFactory({ id: "project-bind-text", bookTitle: "Bind Text" });
  const { exports, storage } = loadContent({
    topFrame: true,
    visibleWordCount: 500,
    location: locationForSurface(surface),
    sendMessage(message, callback) {
      if (message.aceType === "ace-google-doc-word-count") {
        callback({ ok: true, status: 200, wordCount: 500, revisionId: "five-hundred" });
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

  assert.equal(storage.aceDocumentBaselines[surface.manuscriptSurfaceId].endDocumentWordCount, 500);
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
    initialStorage: {
      aceActiveSession: sessionFactory({
        ...tabA,
        startedAt: new Date(Date.now() - 30000).toISOString()
      })
    },
    location: locationForSurface(tabA)
  });
  await new Promise((resolve) => setImmediate(resolve));
  await exports.aceHandleSurfaceLifecycleChange("qa-init");

  context.location.hash = locationForSurface(tabB).hash;
  context.location.href = locationForSurface(tabB).href;
  await exports.aceHandleSurfaceLifecycleChange("qa-switch");
  assert.equal(exports.aceTestState().state, "tab-blocked");

  context.location.hash = locationForSurface(tabA).hash;
  context.location.href = locationForSurface(tabA).href;
  await exports.aceHandleSurfaceLifecycleChange("qa-return");
  assert.equal(exports.aceTestState().state, "active");
});
