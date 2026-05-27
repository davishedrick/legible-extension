const assert = require("node:assert/strict");
const { spawnSync } = require("node:child_process");
const path = require("node:path");
const test = require("node:test");
const { loadContent } = require("./helpers");

function session(overrides = {}) {
  return {
    documentId: "doc-test",
    projectId: "project-test",
    sessionType: "editing",
    startedAt: "2026-05-20T00:00:00.000Z",
    endedAt: "2026-05-20T00:02:00.000Z",
    durationMinutes: 2,
    source: "chrome-extension",
    documentUrl: "https://docs.google.com/document/d/doc-test/edit",
    extensionSessionId: "session-test",
    wordsWritten: 0,
    wordsEdited: 999,
    wordsAdded: 999,
    wordsRemoved: 999,
    netWordsChanged: 1,
    startDocumentWordCount: 4,
    endDocumentWordCount: 5,
    wordCountMethod: "google-docs-api",
    measurementPending: false,
    ...overrides
  };
}

function catchUpSurface(overrides = {}) {
  return {
    documentId: "doc-test",
    tabId: "tab-1",
    tabTitle: "Tab 1",
    manuscriptSurfaceId: "doc-test:tab-1",
    manuscriptSurfaceLabel: "Tab 1",
    ...overrides
  };
}

function catchUpBaseline(count = 1000, overrides = {}) {
  return {
    ...catchUpSurface(),
    projectId: "project-test",
    endDocumentWordCount: count,
    revisionId: "baseline-revision",
    syncedAt: "2026-05-20T00:00:00.000Z",
    ...overrides
  };
}

function catchUpSnapshot(count, overrides = {}) {
  return {
    ok: true,
    wordCount: count,
    apiWordCount: count,
    visibleWordCount: count,
    currentCountSource: "stable-visible",
    currentCountTrusted: true,
    revisionId: "current-revision",
    ...overrides
  };
}

function evaluateCatchUp(exports, overrides = {}) {
  const surface = overrides.surface || catchUpSurface();
  return exports.aceEvaluateCatchUpCandidate({
    trigger: "test",
    surface,
    baseline: Object.prototype.hasOwnProperty.call(overrides, "baseline")
      ? overrides.baseline
      : catchUpBaseline(1000, surface),
    baselineKey: overrides.baselineKey || surface.manuscriptSurfaceId,
    baselineIsLegacy: Boolean(overrides.baselineIsLegacy),
    currentSnapshot: overrides.currentSnapshot || catchUpSnapshot(1200),
    pendingSession: overrides.pendingSession || null,
    completedSession: overrides.completedSession || null,
    binding: Object.prototype.hasOwnProperty.call(overrides, "binding")
      ? overrides.binding
      : { projectId: "project-test", project: { id: "project-test", bookTitle: "Test Book" } }
  });
}

test("sync payload uses net words and zeroes legacy added/removed fields", () => {
  const { exports } = loadContent();
  const payload = exports.aceSessionSyncPayload(session({
    tabId: "tabABC",
    tabTitle: "Draft V2",
    manuscriptSurfaceId: "doc-test:tabABC",
    manuscriptSurfaceLabel: "Draft V2"
  }));

  assert.equal(payload.sessionType, "editing");
  assert.equal(payload.documentId, "doc-test");
  assert.equal(payload.tabId, "tabABC");
  assert.equal(payload.tabTitle, "Draft V2");
  assert.equal(payload.manuscriptSurfaceId, "doc-test:tabABC");
  assert.equal(payload.wordsAdded, 0);
  assert.equal(payload.wordsRemoved, 0);
  assert.equal(payload.wordsEdited, 0);
  assert.equal(payload.netWordsChanged, 1);
  assert.equal(payload.startDocumentWordCount, 4);
  assert.equal(payload.endDocumentWordCount, 5);
  assert.equal(exports.aceMeasurementPathForSession(payload), "google-docs-net-count");
});

test("manuscript surface id helpers normalize tab identity", () => {
  const { exports } = loadContent();

  assert.equal(exports.aceCreateManuscriptSurfaceId("doc123", "tabABC"), "doc123:tabABC");
  assert.equal(exports.aceCreateManuscriptSurfaceId("doc123", ""), "doc123:default");
  assert.equal(exports.aceNormalizeTabId(null), "default");
  assert.equal(exports.aceNormalizeTabTitle("  Draft   V2  "), "Draft V2");
});

test("current manuscript surface falls back to default tab when unavailable", () => {
  const { exports } = loadContent();
  const surface = exports.aceCurrentManuscriptSurface("doc123");

  assert.deepEqual({ ...surface }, {
    documentId: "doc123",
    tabId: "default",
    tabTitle: "",
    manuscriptSurfaceId: "doc123:default",
    manuscriptSurfaceLabel: "Current manuscript"
  });
});

test("local binding lookup prefers manuscript surface id", async () => {
  const { exports, storage } = loadContent();
  storage.aceDocumentBindings = {
    doc123: { projectId: "old-project", project: { id: "old-project" } },
    "doc123:tabABC": {
      documentId: "doc123",
      tabId: "tabABC",
      manuscriptSurfaceId: "doc123:tabABC",
      projectId: "surface-project",
      project: { id: "surface-project" }
    }
  };

  const binding = await exports.aceGetLocalDocumentBinding({
    documentId: "doc123",
    tabId: "tabABC"
  });

  assert.equal(binding.projectId, "surface-project");
  assert.equal(binding.manuscriptSurfaceId, "doc123:tabABC");
});

test("documentId binding fallback only applies to default tab", async () => {
  const { exports, storage } = loadContent();
  storage.aceDocumentBindings = {
    doc123: { projectId: "old-project", project: { id: "old-project" } }
  };

  const defaultBinding = await exports.aceGetLocalDocumentBinding({
    documentId: "doc123",
    tabId: "default"
  });
  const tabBinding = await exports.aceGetLocalDocumentBinding({
    documentId: "doc123",
    tabId: "tabB"
  });

  assert.equal(defaultBinding.projectId, "old-project");
  assert.equal(defaultBinding.manuscriptSurfaceId, "doc123:default");
  assert.equal(tabBinding, null);
});

test("saving local bindings stores by manuscript surface id", async () => {
  const { exports, storage } = loadContent();

  await exports.aceSaveLocalDocumentBinding({
    documentId: "doc123",
    tabId: "tabABC",
    tabTitle: "Draft V2",
    manuscriptSurfaceId: "doc123:tabABC",
    manuscriptSurfaceLabel: "Draft V2"
  }, { id: "project-a", bookTitle: "A" });

  assert.equal(storage.aceDocumentBindings["doc123:tabABC"].projectId, "project-a");
  assert.equal(storage.aceDocumentBindings["doc123:tabABC"].tabTitle, "Draft V2");
  assert.equal(storage.aceDocumentBindings.doc123, undefined);
});

test("removing a local binding affects only the current manuscript surface", async () => {
  const { exports, storage } = loadContent();
  storage.aceDocumentBindings = {
    "doc123:tabA": { projectId: "project-a", project: { id: "project-a" } },
    "doc123:tabB": { projectId: "project-b", project: { id: "project-b" } }
  };

  await exports.aceRemoveLocalDocumentBinding({
    documentId: "doc123",
    tabId: "tabA",
    manuscriptSurfaceId: "doc123:tabA"
  });

  assert.equal(storage.aceDocumentBindings["doc123:tabA"], undefined);
  assert.equal(storage.aceDocumentBindings["doc123:tabB"].projectId, "project-b");
});

test("create project validation keeps required fields compact and strict", () => {
  const { exports } = loadContent();
  const draft = exports.aceDefaultCreateProjectDraft(528);

  assert.equal(draft.manuscriptType, "Novel");
  assert.equal(draft.structureUnit, "Chapter");
  assert.equal(draft.targetWordCount, 80000);
  assert.equal(draft.wordsWrittenSoFar, 528);
  assert.equal(exports.aceValidateCreateProjectDraft(draft), "Title required.");
  assert.equal(exports.aceValidateCreateProjectDraft({
    ...draft,
    title: "The Hollow Orchard"
  }), "");
  assert.equal(exports.aceValidateCreateProjectDraft({
    ...draft,
    title: "The Hollow Orchard",
    targetWordCount: 0
  }), "Target must be positive.");
});

test("baseline lookup prefers manuscript surface and falls back only for default", async () => {
  const { exports, storage } = loadContent();
  storage.aceDocumentBaselines = {
    doc123: { endDocumentWordCount: 100 },
    "doc123:tabABC": { endDocumentWordCount: 200 }
  };

  const surfaceBaseline = await exports.aceGetDocumentBaseline({
    documentId: "doc123",
    tabId: "tabABC"
  });
  const defaultBaseline = await exports.aceGetDocumentBaseline({
    documentId: "doc123",
    tabId: "default"
  });
  const missingBaseline = await exports.aceGetDocumentBaseline({
    documentId: "doc123",
    tabId: "tabB"
  });

  assert.equal(surfaceBaseline.endDocumentWordCount, 200);
  assert.equal(defaultBaseline.endDocumentWordCount, 100);
  assert.equal(missingBaseline, null);
});

test("saving document baseline stores surface metadata", async () => {
  const { exports, storage } = loadContent();

  await exports.aceSaveDocumentBaseline(session({
    tabId: "tabABC",
    tabTitle: "Draft V2",
    manuscriptSurfaceId: "doc-test:tabABC",
    manuscriptSurfaceLabel: "Draft V2",
    endDocumentWordCount: 814
  }), { id: "project-test" });

  const baseline = storage.aceDocumentBaselines["doc-test:tabABC"];
  assert.equal(baseline.documentId, "doc-test");
  assert.equal(baseline.tabId, "tabABC");
  assert.equal(baseline.manuscriptSurfaceId, "doc-test:tabABC");
  assert.equal(baseline.endDocumentWordCount, 814);
});

test("catch-up baseline lookup requires current manuscript surface", async () => {
  const { exports, storage } = loadContent();
  storage.aceDocumentBaselines = {
    "doc-test:tabA": catchUpBaseline(1000, {
      tabId: "tabA",
      manuscriptSurfaceId: "doc-test:tabA"
    })
  };

  const tabA = await exports.aceGetDocumentBaselineForCatchUp({
    documentId: "doc-test",
    tabId: "tabA",
    manuscriptSurfaceId: "doc-test:tabA"
  });
  const tabB = await exports.aceGetDocumentBaselineForCatchUp({
    documentId: "doc-test",
    tabId: "tabB",
    manuscriptSurfaceId: "doc-test:tabB"
  });

  assert.equal(tabA.baseline.endDocumentWordCount, 1000);
  assert.equal(tabA.key, "doc-test:tabA");
  assert.equal(tabB.baseline, null);
  assert.equal(tabB.key, "doc-test:tabB");
});

test("catch-up decision allows trustworthy positive net", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1000),
    currentSnapshot: catchUpSnapshot(1200)
  });

  assert.equal(result.candidate.netWordsChanged, 200);
  assert.equal(result.candidate.sessionType, "writing");
  assert.equal(result.trace.decision.action, "show-catch-up");
  assert.equal(result.trace.decision.code, "D-CATCHUP-SHOW-POSITIVE");
});

test("catch-up decision allows trustworthy negative net", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1000),
    currentSnapshot: catchUpSnapshot(900)
  });

  assert.equal(result.candidate.netWordsChanged, -100);
  assert.equal(result.candidate.sessionType, "editing");
  assert.equal(result.trace.decision.action, "show-catch-up");
  assert.equal(result.trace.decision.code, "D-CATCHUP-SHOW-NEGATIVE");
});

test("catch-up decision allows full deletion to zero words", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1000),
    currentSnapshot: catchUpSnapshot(0)
  });

  assert.equal(result.candidate.netWordsChanged, -1000);
  assert.equal(result.candidate.sessionType, "editing");
  assert.equal(result.candidate.startDocumentWordCount, 1000);
  assert.equal(result.candidate.endDocumentWordCount, 0);
  assert.equal(result.trace.decision.code, "D-CATCHUP-SHOW-NEGATIVE");
});

test("catch-up decision skips zero change", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1000),
    currentSnapshot: catchUpSnapshot(1000)
  });

  assert.equal(result.candidate, null);
  assert.equal(result.reason, "zero-net");
  assert.equal(result.trace.decision.code, "D-CATCHUP-NET-ZERO");
});

test("catch-up decision skips zero-to-zero change", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(0),
    currentSnapshot: catchUpSnapshot(0)
  });

  assert.equal(result.candidate, null);
  assert.equal(result.reason, "zero-net");
  assert.equal(result.trace.decision.code, "D-CATCHUP-NET-ZERO");
});

test("catch-up decision skips missing baseline", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: null,
    currentSnapshot: catchUpSnapshot(1200)
  });

  assert.equal(result.candidate, null);
  assert.equal(result.reason, "missing-baseline");
  assert.equal(result.trace.decision.code, "D-CATCHUP-NO-BASELINE");
});

test("catch-up decision skips unbound surface", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1000),
    currentSnapshot: catchUpSnapshot(1200),
    binding: null
  });

  assert.equal(result.candidate, null);
  assert.equal(result.reason, "no-binding");
  assert.equal(result.trace.decision.code, "D-CATCHUP-NO-BINDING");
});

test("catch-up decision skips visible-unavailable API negative result", () => {
  const { exports } = loadContent();
  const diagnostic = exports.aceGoogleDocNetDiagnostic({
    code: "D-CATCHUP-NO-STABLE-VISIBLE",
    startWordCount: 1332,
    apiEndWordCount: 1213,
    visibleEndWordCount: null,
    netWordsChanged: -119,
    startSource: "saved-total-baseline",
    endSource: "google-docs-api"
  });
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1332),
    currentSnapshot: catchUpSnapshot(null, {
      ok: false,
      skipCatchUp: true,
      wordCount: null,
      apiWordCount: 1213,
      visibleWordCount: null,
      currentCountSource: "none",
      currentCountTrusted: false,
      wordCountDiagnostic: diagnostic
    })
  });

  assert.equal(result.candidate, null);
  assert.equal(result.reason, "current-count-diagnostic-only");
  assert.match(diagnostic, /visible end unavailable/);
});

test("catch-up decision trusts stable visible count over API mismatch", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1000),
    currentSnapshot: catchUpSnapshot(1200, {
      apiWordCount: 1213,
      visibleWordCount: 1200,
      currentCountSource: "stable-visible",
      currentCountTrusted: true,
      apiVisibleMismatch: true
    })
  });

  assert.equal(result.candidate.netWordsChanged, 200);
  assert.equal(result.trace.currentCounts.apiCount, 1213);
  assert.equal(result.trace.currentCounts.selectedVisibleCount, 1200);
  assert.equal(result.trace.decision.action, "show-catch-up");
});

test("stable visible count ignores early zero reads for catch-up", async () => {
  const { exports } = loadContent();
  const reads = [0, 0, 1332, 1332];
  const result = await exports.aceStableVisibleGoogleDocWordCount({
    delayMs: 0,
    timeoutMs: 100,
    ignoreZero: true,
    readVisibleCount: async () => {
      const count = reads.shift();
      return {
        count,
        candidates: [{ count, snippet: `${count} words`, score: count }],
        selectedCandidate: { count, snippet: `${count} words`, score: count }
      };
    }
  });

  assert.equal(result.count, 1332);
  assert.equal(result.stable, true);
});

test("catch-up current count treats stable visible zero as valid", async () => {
  const { exports } = loadContent();
  const reads = [0, 0];
  const result = await exports.aceGoogleDocWordCountAfterSettle("doc-test", catchUpBaseline(1000), {
    settleDelayMs: 0,
    visibleDelayMs: 0,
    visibleTimeoutMs: 100,
    readVisibleCount: async () => {
      const count = reads.shift();
      return {
        count,
        candidates: [{ count, snippet: `${count} words` }],
        selectedCandidate: { count, snippet: `${count} words` }
      };
    },
    apiCall: async () => ({ ok: true, wordCount: 0, revisionId: "empty-doc" })
  });

  assert.equal(result.ok, true);
  assert.equal(result.wordCount, 0);
  assert.equal(result.visibleWordCount, 0);
  assert.equal(result.netWordsChanged, -1000);
  assert.equal(result.currentCountTrusted, true);
});

test("catch-up decision skips positive API without stable visible count", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1000),
    currentSnapshot: catchUpSnapshot(null, {
      ok: false,
      skipCatchUp: true,
      wordCount: null,
      apiWordCount: 1200,
      visibleWordCount: null,
      currentCountSource: "none",
      currentCountTrusted: false
    })
  });

  assert.equal(result.candidate, null);
  assert.equal(result.reason, "current-count-diagnostic-only");
});

test("catch-up decision skips baseline surface mismatch", () => {
  const { exports } = loadContent();
  const surface = catchUpSurface({
    tabId: "tab-2",
    tabTitle: "Tab 2",
    manuscriptSurfaceId: "doc-test:tab-2"
  });
  const result = evaluateCatchUp(exports, {
    surface,
    baseline: catchUpBaseline(1000, {
      tabId: "tab-1",
      tabTitle: "Tab 1",
      manuscriptSurfaceId: "doc-test:tab-1"
    }),
    baselineKey: "doc-test:tab-1",
    currentSnapshot: catchUpSnapshot(1200)
  });

  assert.equal(result.candidate, null);
  assert.equal(result.reason, "baseline-surface-mismatch");
});

test("catch-up decision skips when pending session covers surface", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    pendingSession: session({
      tabId: "tab-1",
      manuscriptSurfaceId: "doc-test:tab-1",
      startDocumentWordCount: 1000,
      endDocumentWordCount: 1200,
      source: "catch-up"
    })
  });

  assert.equal(result.candidate, null);
  assert.equal(result.reason, "pending-or-completed-session-exists");
});

test("negative catch-up creates editing session payload", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1000),
    currentSnapshot: catchUpSnapshot(900)
  });
  const payload = exports.aceBuildCatchUpSession(result.candidate, "project-test", Date.parse("2026-05-20T00:02:00.000Z"));

  assert.equal(payload.sessionType, "editing");
  assert.equal(payload.source, "catch-up");
  assert.equal(payload.netWordsChanged, -100);
  assert.equal(payload.wordsWritten, 0);
  assert.equal(payload.startDocumentWordCount, 1000);
  assert.equal(payload.endDocumentWordCount, 900);
  assert.equal(payload.durationMinutes, 1);
});

test("full deletion catch-up creates editing session payload", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1000),
    currentSnapshot: catchUpSnapshot(0)
  });
  const payload = exports.aceBuildCatchUpSession(result.candidate, "project-test", Date.parse("2026-05-20T00:02:00.000Z"));

  assert.equal(payload.sessionType, "editing");
  assert.equal(payload.source, "catch-up");
  assert.equal(payload.netWordsChanged, -1000);
  assert.equal(payload.startDocumentWordCount, 1000);
  assert.equal(payload.endDocumentWordCount, 0);
});

test("positive catch-up creates writing session payload", () => {
  const { exports } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1000),
    currentSnapshot: catchUpSnapshot(1200)
  });
  const payload = exports.aceBuildCatchUpSession(result.candidate, "project-test", Date.parse("2026-05-20T00:02:00.000Z"));

  assert.equal(payload.sessionType, "writing");
  assert.equal(payload.source, "catch-up");
  assert.equal(payload.netWordsChanged, 200);
  assert.equal(payload.wordsWritten, 200);
  assert.equal(payload.startDocumentWordCount, 1000);
  assert.equal(payload.endDocumentWordCount, 1200);
});

test("skipping catch-up updates baseline to current count", async () => {
  const { exports, storage } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1000),
    currentSnapshot: catchUpSnapshot(1200)
  });

  await exports.aceSaveSkippedCatchUpBaseline(result.candidate);

  assert.equal(storage.aceDocumentBaselines["doc-test:tab-1"].endDocumentWordCount, 1200);
  assert.equal(storage.aceDocumentBaselines["doc-test:tab-1"].wordCountMethod, "stable-visible");
});

test("skipping full deletion updates baseline to zero", async () => {
  const { exports, storage } = loadContent();
  const result = evaluateCatchUp(exports, {
    baseline: catchUpBaseline(1000),
    currentSnapshot: catchUpSnapshot(0)
  });

  await exports.aceSaveSkippedCatchUpBaseline(result.candidate);

  assert.equal(storage.aceDocumentBaselines["doc-test:tab-1"].endDocumentWordCount, 0);
  assert.equal(storage.aceDocumentBaselines["doc-test:tab-1"].wordCountMethod, "stable-visible");
});

test("normal session baseline sync updates baseline to session end", async () => {
  const { exports, storage } = loadContent();

  await exports.aceSaveDocumentBaseline(session({
    tabId: "tab-1",
    tabTitle: "Tab 1",
    manuscriptSurfaceId: "doc-test:tab-1",
    endDocumentWordCount: 1300
  }), { id: "project-test" });

  assert.equal(storage.aceDocumentBaselines["doc-test:tab-1"].endDocumentWordCount, 1300);
});

test("start snapshot accepts API word count zero", async () => {
  const messages = [];
  const { exports } = loadContent({
    sendMessage(message, callback) {
      messages.push(message);
      callback({
        ok: true,
        status: 200,
        method: "google-docs-api",
        revisionId: "empty-doc",
        wordCount: 0,
        wordCountTokenizerVersion: "test-tokenizer"
      });
    }
  });

  const snapshot = await exports.aceStartSnapshotWithVisibleFallback(
    "doc-test",
    "session-zero",
    true,
    "test start",
    { allowVisibleFallback: true }
  );

  assert.equal(snapshot.ok, true);
  assert.equal(snapshot.wordCount, 0);
  assert.equal(messages[0].extensionSessionId, "session-zero");
});

test("baseline start snapshot accepts zero and stores under session id", async () => {
  const { exports, storage } = loadContent();

  const snapshot = await exports.aceSeedGoogleDocStartSnapshotFromBaseline(
    "doc-test",
    "session-zero",
    catchUpBaseline(0)
  );

  assert.equal(snapshot.ok, true);
  assert.equal(snapshot.wordCount, 0);
  assert.equal(storage["aceWordSnapshot:session-zero"].wordCount, 0);
  assert.equal(storage["aceWordSnapshot:session-zero"].documentId, "doc-test");
});

test("restoring a pending session uses the pending session project", async () => {
  const { exports, storage } = loadContent({ topFrame: true });
  await new Promise((resolve) => setImmediate(resolve));
  storage.acePendingSessions = [
    session({
      project: { id: "project-test", bookTitle: "Test Book" }
    })
  ];

  await assert.doesNotReject(exports.aceRestorePendingSessionForCurrentDocument());
});

test("issue payload includes manuscript surface identity", () => {
  const { exports } = loadContent();
  const payload = exports.aceIssueSyncPayload({
    documentId: "doc123",
    tabId: "tabABC",
    tabTitle: "Draft V2",
    manuscriptSurfaceId: "doc123:tabABC",
    manuscriptSurfaceLabel: "Draft V2",
    extensionIssueId: "issue-1",
    documentUrl: "https://docs.google.com/document/d/doc123/edit"
  }, "project-a", "Fix tone", "quoted words");

  assert.equal(payload.documentId, "doc123");
  assert.equal(payload.tabId, "tabABC");
  assert.equal(payload.tabTitle, "Draft V2");
  assert.equal(payload.manuscriptSurfaceId, "doc123:tabABC");
  assert.equal(payload.projectId, "project-a");
});

test("active session surface mismatch is detected", () => {
  const { exports } = loadContent();

  assert.equal(exports.aceSessionSurfaceMismatch({
    documentId: "doc123",
    tabId: "tabA",
    manuscriptSurfaceId: "doc123:tabA"
  }, {
    documentId: "doc123",
    tabId: "tabB",
    manuscriptSurfaceId: "doc123:tabB"
  }), true);
  assert.equal(exports.aceSessionSurfaceMismatch({
    documentId: "doc123",
    tabId: "tabA",
    manuscriptSurfaceId: "doc123:tabA"
  }, {
    documentId: "doc123",
    tabId: "tabA",
    manuscriptSurfaceId: "doc123:tabA"
  }), false);
});

test("tab switch refresh clears previous bound prompt state", async () => {
  const { exports, storage, context } = loadContent({
    topFrame: true,
    location: {
      href: "https://docs.google.com/document/d/doc-test/edit#tab=tabA&tabTitle=Tab%20A",
      hash: "#tab=tabA&tabTitle=Tab%20A"
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  storage.aceDocumentBindings = {
    "doc-test:tabA": {
      documentId: "doc-test",
      tabId: "tabA",
      tabTitle: "Tab A",
      manuscriptSurfaceId: "doc-test:tabA",
      projectId: "project-a",
      project: { id: "project-a", bookTitle: "Project A" }
    }
  };

  await exports.aceRefreshStateForSurfaceSwitch(exports.aceCurrentManuscriptSurface(), "test");
  assert.equal(exports.aceTestState().currentBinding.projectId, "project-a");

  context.location.hash = "#tab=tabB&tabTitle=Tab%20B";
  context.location.href = "https://docs.google.com/document/d/doc-test/edit#tab=tabB&tabTitle=Tab%20B";
  await exports.aceHandleSurfaceLifecycleChange("test-tab-switch");

  const state = exports.aceTestState();
  assert.equal(state.currentSurface.manuscriptSurfaceId, "doc-test:tabB");
  assert.equal(state.currentBinding, null);
  assert.equal(state.catchUpCandidate, null);
  assert.match(state.widgetHtml, /Not bound/);
  assert.match(state.widgetHtml, /Tab B/);
  assert.doesNotMatch(state.widgetHtml, /Project A/);
});

test("active session blocks on tab switch and resumes on original tab", async () => {
  const activeSession = session({
    tabId: "tabA",
    tabTitle: "Tab A",
    manuscriptSurfaceId: "doc-test:tabA",
    manuscriptSurfaceLabel: "Tab A",
    startedAt: new Date(Date.now() - 60000).toISOString(),
    project: { id: "project-a", bookTitle: "Project A" },
    projectId: "project-a"
  });
  const { exports, context } = loadContent({
    topFrame: true,
    initialStorage: { aceActiveSession: activeSession },
    location: {
      href: "https://docs.google.com/document/d/doc-test/edit#tab=tabA&tabTitle=Tab%20A",
      hash: "#tab=tabA&tabTitle=Tab%20A"
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  await exports.aceHandleSurfaceLifecycleChange("test-init");

  context.location.hash = "#tab=tabB&tabTitle=Tab%20B";
  context.location.href = "https://docs.google.com/document/d/doc-test/edit#tab=tabB&tabTitle=Tab%20B";
  await exports.aceHandleSurfaceLifecycleChange("test-tab-switch");

  let state = exports.aceTestState();
  assert.equal(state.state, "tab-blocked");
  assert.equal(Boolean(state.activeSession.tabBlockedAt), true);
  assert.match(state.widgetHtml, /Tab changed/);
  assert.match(state.widgetHtml, /Tab A/);

  context.location.hash = "#tab=tabA&tabTitle=Tab%20A";
  context.location.href = "https://docs.google.com/document/d/doc-test/edit#tab=tabA&tabTitle=Tab%20A";
  await exports.aceHandleSurfaceLifecycleChange("test-tab-return");

  state = exports.aceTestState();
  assert.equal(state.state, "active");
  assert.equal(state.activeSession.tabBlockedAt, "");
  assert(state.activeSession.pausedDurationMs >= 0);
});

test("tab title rename keeps surface id and updates active session metadata", async () => {
  const activeSession = session({
    tabId: "tabA",
    tabTitle: "Old Title",
    manuscriptSurfaceId: "doc-test:tabA",
    manuscriptSurfaceLabel: "Old Title",
    startedAt: new Date(Date.now() - 60000).toISOString()
  });
  const { exports, context } = loadContent({
    topFrame: true,
    initialStorage: { aceActiveSession: activeSession },
    location: {
      href: "https://docs.google.com/document/d/doc-test/edit#tab=tabA&tabTitle=Old%20Title",
      hash: "#tab=tabA&tabTitle=Old%20Title"
    }
  });
  await new Promise((resolve) => setImmediate(resolve));
  await exports.aceHandleSurfaceLifecycleChange("test-init");

  context.location.hash = "#tab=tabA&tabTitle=New%20Title";
  context.location.href = "https://docs.google.com/document/d/doc-test/edit#tab=tabA&tabTitle=New%20Title";
  await exports.aceHandleSurfaceLifecycleChange("test-title-rename");

  const state = exports.aceTestState();
  assert.equal(state.activeSession.manuscriptSurfaceId, "doc-test:tabA");
  assert.equal(state.activeSession.tabTitle, "New Title");
});

test("net words are computed from start and end counts", () => {
  const { exports } = loadContent();

  assert.equal(exports.aceSessionSyncPayload(session({
    startDocumentWordCount: 0,
    endDocumentWordCount: 600
  })).netWordsChanged, 600);
  assert.equal(exports.aceSessionSyncPayload(session({
    startDocumentWordCount: 1000,
    endDocumentWordCount: 1000
  })).netWordsChanged, 0);
  assert.equal(exports.aceSessionSyncPayload(session({
    startDocumentWordCount: 1000,
    endDocumentWordCount: 800
  })).netWordsChanged, -200);
});

test("old pending sessions derive net before retry payload sync", () => {
  const { exports } = loadContent();

  assert.equal(exports.aceSessionSyncPayload(session({
    sessionType: "writing",
    startDocumentWordCount: null,
    endDocumentWordCount: null,
    netWordsChanged: undefined,
    wordsWritten: 250
  })).netWordsChanged, 250);
  assert.equal(exports.aceSessionSyncPayload(session({
    sessionType: "editing",
    startDocumentWordCount: null,
    endDocumentWordCount: null,
    netWordsChanged: undefined,
    wordsAdded: 40,
    wordsRemoved: 60
  })).netWordsChanged, -20);
});

test("editing session UI copy displays net only", () => {
  const { exports } = loadContent();
  const copy = exports.aceSessionWordsCopy(session({
    sessionType: "editing",
    startDocumentWordCount: 714,
    endDocumentWordCount: 814,
    netWordsChanged: 100
  }));

  assert.equal(copy, " · Net: +100 words");
  assert(!copy.includes("- 40"));
});

test("writing session UI copy displays net only", () => {
  const { exports } = loadContent();
  const copy = exports.aceSessionWordsCopy(session({
    sessionType: "writing",
    wordsWritten: 600,
    startDocumentWordCount: 0,
    endDocumentWordCount: 600,
    netWordsChanged: 600
  }));

  assert.equal(copy, " · Net: +600 words");
});

test("negative and zero net copy is explicit", () => {
  const { exports } = loadContent();

  assert.equal(exports.aceSessionWordsCopy(session({
    startDocumentWordCount: 1000,
    endDocumentWordCount: 1000,
    netWordsChanged: 0
  })), " · Net: 0 words");
  assert.equal(exports.aceSessionWordsCopy(session({
    startDocumentWordCount: 1000,
    endDocumentWordCount: 800,
    netWordsChanged: -200
  })), " · Net: -200 words");
});

test("visible word count reader probes the bottom-left Google Docs counter", () => {
  const { exports } = loadContent();
  const makeElement = ({ textContent = "", innerText = "", rect, parentElement = null }) => ({
    textContent,
    innerText,
    parentElement,
    closest() { return null; },
    getAttribute() { return ""; },
    getBoundingClientRect() { return rect; }
  });
  const falseZero = makeElement({
    textContent: "0 words",
    rect: { left: 24, top: 1030, right: 180, bottom: 1080, width: 156, height: 50 }
  });
  const visibleCounter = makeElement({
    innerText: "1,918 words",
    rect: { left: 20, top: 980, right: 280, bottom: 1054, width: 260, height: 74 }
  });
  const targetDocument = {
    documentElement: { clientWidth: 2048, clientHeight: 1152 },
    querySelectorAll() {
      return [falseZero];
    },
    elementsFromPoint() {
      return [visibleCounter];
    }
  };
  const candidates = exports.aceVisibleWordCountCandidatesInDocument(
    targetDocument,
    { innerWidth: 2048, innerHeight: 1152 }
  );
  const best = candidates.sort((a, b) => b.score - a.score)[0];

  assert.equal(best.count, 1918);
});

test("visible word count parsing accepts exact Google Docs counter text only", () => {
  const { exports } = loadContent();
  const element = (text) => ({
    textContent: text,
    innerText: "",
    getAttribute() { return ""; }
  });

  assert.equal(exports.aceVisibleWordCountFromElement(element("2,209 words")), 2209);
  assert.equal(exports.aceVisibleWordCountFromElement(element("Diagnostic: visible end 2208 words")), null);
});

test("stable visible word count uses the latest stable value", async () => {
  const { exports } = loadContent();
  const reads = [2208, 2209, 2209];

  const result = await exports.aceStableVisibleGoogleDocWordCount({
    delayMs: 0,
    timeoutMs: 100,
    readVisibleCount: async () => {
      const count = reads.shift();
      return {
        count,
        candidates: [{ count, snippet: `${count} words`, score: count }],
        selectedCandidate: { count, snippet: `${count} words`, score: count }
      };
    }
  });

  assert.equal(result.count, 2209);
  assert.equal(result.stable, true);
  assert.match(result.diagnostic, /W-VISIBLE-COUNT-CHANGED/);
});

test("stable visible word count finishes after two matching reads", async () => {
  const { exports } = loadContent();
  const reads = [2209, 2209, 2210, 2210];

  const result = await exports.aceStableVisibleGoogleDocWordCount({
    delayMs: 0,
    timeoutMs: 100,
    readVisibleCount: async () => {
      const count = reads.shift();
      return {
        count,
        candidates: [{ count, snippet: `${count} words` }],
        selectedCandidate: { count, snippet: `${count} words` }
      };
    }
  });

  assert.equal(result.count, 2209);
  assert.equal(result.stable, true);
  assert.equal(result.readCount, 2);
});

test("session end uses stable visible count without waiting for slow API", async () => {
  const { exports } = loadContent();
  const reads = [2209, 2209];
  let apiFinished = false;

  const result = await exports.aceGoogleDocNetAfterSave(
    "doc-test",
    "session-test",
    1000,
    "start-revision",
    true,
    {
      settleDelayMs: 0,
      visibleDelayMs: 0,
      visibleTimeoutMs: 100,
      readVisibleCount: async () => {
        const count = reads.shift();
        return {
          count,
          candidates: [{ count, snippet: `${count} words` }],
          selectedCandidate: { count, snippet: `${count} words` }
        };
      },
      apiCall: () => new Promise((resolve) => {
        setTimeout(() => {
          apiFinished = true;
          resolve({ ok: true, wordCount: 2208, revisionId: "api-revision" });
        }, 50);
      })
    }
  );

  assert.equal(result.ok, true);
  assert.equal(result.wordCount, 2209);
  assert.equal(result.netWordsChanged, 1209);
  assert.equal(result.method, "stable-visible-count");
  assert.equal(result.timing.apiAttempts, 1);
  assert.equal(apiFinished, false);
});

test("session end records API mismatch but keeps stable visible count", async () => {
  const { exports } = loadContent();
  const reads = [2209, 2209];

  const result = await exports.aceGoogleDocNetAfterSave(
    "doc-test",
    "session-test",
    1000,
    "start-revision",
    true,
    {
      settleDelayMs: 0,
      visibleDelayMs: 0,
      visibleTimeoutMs: 100,
      readVisibleCount: async () => {
        const count = reads.shift();
        return {
          count,
          candidates: [{ count, snippet: `${count} words` }],
          selectedCandidate: { count, snippet: `${count} words` }
        };
      },
      apiCall: async () => ({ ok: true, wordCount: 2208, revisionId: "api-revision" })
    }
  );

  assert.equal(result.ok, true);
  assert.equal(result.wordCount, 2209);
  assert.match(result.wordCountDiagnostic, /W-API-VISIBLE-MISMATCH|D-STABLE-VISIBLE-NET/);
});

test("session end can use API only when surface is explicitly trusted", async () => {
  const { exports } = loadContent();

  const result = await exports.aceGoogleDocNetAfterSave(
    "doc-test",
    "session-test",
    1000,
    "start-revision",
    true,
    {
      settleDelayMs: 0,
      visibleDelayMs: 0,
      visibleTimeoutMs: 0,
      apiSurfaceTrusted: true,
      readVisibleCount: async () => ({ count: null, candidates: [], selectedCandidate: null }),
      apiCall: async () => ({ ok: true, wordCount: 1200, revisionId: "api-revision" })
    }
  );

  assert.equal(result.ok, true);
  assert.equal(result.wordCount, 1200);
  assert.equal(result.netWordsChanged, 200);
  assert.match(result.wordCountDiagnostic, /D-API-FALLBACK-NET/);
});

test("session end does not use untrusted API when visible is unavailable", async () => {
  const { exports } = loadContent();

  const result = await exports.aceGoogleDocNetAfterSave(
    "doc-test",
    "session-test",
    1000,
    "start-revision",
    true,
    {
      settleDelayMs: 0,
      visibleDelayMs: 0,
      visibleTimeoutMs: 0,
      readVisibleCount: async () => ({ count: null, candidates: [], selectedCandidate: null }),
      apiCall: async () => ({ ok: true, wordCount: 1200, revisionId: "api-revision" })
    }
  );

  assert.equal(result.ok, false);
  assert.equal(result.wordCount, null);
  assert.match(result.wordCountDiagnostic, /E-NO-TRUSTED-END-COUNT/);
});

test("catch-up check uses stable visible count for positive and negative changes", async () => {
  const { exports } = loadContent();
  const positiveReads = [1200, 1200];
  const positive = await exports.aceGoogleDocWordCountAfterSettle("doc-test", catchUpBaseline(1000), {
    settleDelayMs: 0,
    visibleDelayMs: 0,
    visibleTimeoutMs: 100,
    readVisibleCount: async () => {
      const count = positiveReads.shift();
      return { count, candidates: [{ count }], selectedCandidate: { count } };
    },
    apiCall: async () => ({ ok: true, wordCount: 1200 })
  });

  const negativeReads = [900, 900];
  const negative = await exports.aceGoogleDocWordCountAfterSettle("doc-test", catchUpBaseline(1000), {
    settleDelayMs: 0,
    visibleDelayMs: 0,
    visibleTimeoutMs: 100,
    readVisibleCount: async () => {
      const count = negativeReads.shift();
      return { count, candidates: [{ count }], selectedCandidate: { count } };
    },
    apiCall: async () => ({ ok: true, wordCount: 900 })
  });

  assert.equal(positive.ok, true);
  assert.equal(positive.netWordsChanged, 200);
  assert.equal(negative.ok, true);
  assert.equal(negative.netWordsChanged, -100);
});

test("conflicting visible candidates prefer the bottom-left current counter", () => {
  const { exports } = loadContent();
  const best = exports.aceBestVisibleWordCountCandidate([
    { count: 2208, score: 1, bottom: 900, left: 20, snippet: "2,208 words" },
    { count: 2209, score: 3, bottom: 1080, left: 20, snippet: "2,209 words" }
  ]);
  const diagnostic = exports.aceVisibleCountDiagnostic({
    count: best.count,
    stable: true,
    reads: [{ count: best.count }],
    candidates: [
      { count: 2208, snippet: "2,208 words" },
      { count: 2209, snippet: "2,209 words" }
    ],
    selectedCandidate: best
  });

  assert.equal(best.count, 2209);
  assert.match(diagnostic, /W-VISIBLE-CANDIDATE-CONFLICT/);
  assert.match(diagnostic, /selected 2209/);
});

test("visible fallback net uses stabilized count", async () => {
  const { exports } = loadContent();
  const result = await exports.aceStableVisibleGoogleDocWordCount({
    delayMs: 0,
    timeoutMs: 100,
    context: {
      apiWordCount: 0,
      startWordCount: 1000,
      startSource: "stored-start-count",
      endSource: "visible-total-fallback"
    },
    readVisibleCount: async () => ({
      count: 2209,
      candidates: [{ count: 2209, snippet: "2,209 words" }],
      selectedCandidate: { count: 2209, snippet: "2,209 words" }
    })
  });
  const netWordsChanged = result.count - 1000;
  const classes = exports.aceVisibleCountClasses(result, {
    apiWordCount: 0,
    startWordCount: 1000,
    startSource: "stored-start-count",
    endSource: "visible-total-fallback"
  });

  assert.equal(netWordsChanged, 1209);
  assert(classes.includes("W-API-ZERO"));
  assert(classes.includes("W-SOURCE-MISMATCH"));
});

test("backend compatibility accepts zeroed legacy fields and applies net words", () => {
  const appPath = path.resolve(__dirname, "..", "..", "Author-companion", "writing_app");
  const script = `
import json, sys
sys.path.insert(0, ${JSON.stringify(appPath)})
from extension_bridge import append_extension_session
state = {
    "projects": [{"id": "project-test", "project": {"bookTitle": "Test", "currentWordCount": 714}, "sessions": []}],
    "extensionDocumentBindings": {"doc-test": "project-test"},
}
payload = {
    "documentId": "doc-test",
    "projectId": "project-test",
    "sessionType": "editing",
    "startedAt": "2026-05-20T00:00:00+00:00",
    "endedAt": "2026-05-20T00:02:00+00:00",
    "durationMinutes": 2,
    "source": "chrome-extension",
    "documentUrl": "https://docs.google.com/document/d/doc-test/edit",
    "extensionSessionId": "session-test",
    "wordsEdited": 0,
    "wordsAdded": 0,
    "wordsRemoved": 0,
    "netWordsChanged": 100,
    "startDocumentWordCount": 714,
    "endDocumentWordCount": 814,
    "wordCountMethod": "google-docs-api",
    "measurementPending": False,
}
session, project, duplicate = append_extension_session(state, payload)
print(json.dumps({
    "duplicate": duplicate,
    "wordsAdded": session["wordsAdded"],
    "wordsRemoved": session["wordsRemoved"],
    "wordsEdited": session["wordsEdited"],
    "netWordsChanged": session["netWordsChanged"],
    "currentWordCount": project["currentWordCount"],
}))
`;
  const result = spawnSync("python3", ["-c", script], { encoding: "utf8" });

  assert.equal(result.status, 0, result.stderr);
  assert.deepEqual(JSON.parse(result.stdout), {
    duplicate: false,
    wordsAdded: 0,
    wordsRemoved: 0,
    wordsEdited: 0,
    netWordsChanged: 100,
    currentWordCount: 814
  });
});
