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
