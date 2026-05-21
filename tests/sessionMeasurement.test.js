const assert = require("node:assert/strict");
const { spawnSync } = require("node:child_process");
const path = require("node:path");
const test = require("node:test");
const { loadContent } = require("./helpers");

function exactEditingSession(overrides = {}) {
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
    wordsEdited: 5,
    wordsAdded: 3,
    wordsRemoved: 2,
    netWordsChanged: 1,
    startDocumentWordCount: 4,
    endDocumentWordCount: 5,
    wordCountMethod: "google-docs-api",
    wordDiffMethod: "google-api-token-sequence",
    measurementPending: false,
    ...overrides
  };
}

test("completed editing session sync payload preserves exact added and removed counts", () => {
  const { exports } = loadContent();
  const payload = exports.aceSessionSyncPayload(exactEditingSession());

  assert.equal(payload.sessionType, "editing");
  assert.equal(payload.wordsAdded, 3);
  assert.equal(payload.wordsRemoved, 2);
  assert.equal(payload.wordsEdited, 5);
  assert.equal(payload.netWordsChanged, 1);
  assert.equal(exports.aceMeasurementPathForSession(payload), "exact-api-sequence-diff");
});

test("net-zero editing payload does not collapse exact +100 / -100 to +0 / -0", () => {
  const { exports } = loadContent();
  const payload = exports.aceSessionSyncPayload(exactEditingSession({
    wordsEdited: 200,
    wordsAdded: 100,
    wordsRemoved: 100,
    netWordsChanged: 0,
    startDocumentWordCount: 1000,
    endDocumentWordCount: 1000
  }));

  assert.equal(payload.wordsAdded, 100);
  assert.equal(payload.wordsRemoved, 100);
  assert.equal(payload.wordsEdited, 200);
  assert.equal(payload.netWordsChanged, 0);
});

test("extension UI copy displays exact editing breakdown", () => {
  const { exports } = loadContent();
  const copy = exports.aceSessionWordsCopy(exactEditingSession());

  assert.equal(copy, " · (+3 words - 2)");
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

test("backend persistence preserves exact extension breakdown", () => {
  const appPath = path.resolve(__dirname, "..", "..", "Author-companion", "writing_app");
  const script = `
import json, sys
sys.path.insert(0, ${JSON.stringify(appPath)})
from extension_bridge import append_extension_session
state = {
    "projects": [{"id": "project-test", "project": {"bookTitle": "Test", "currentWordCount": 4}, "sessions": []}],
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
    "wordsEdited": 5,
    "wordsAdded": 3,
    "wordsRemoved": 2,
    "netWordsChanged": 1,
    "startDocumentWordCount": 4,
    "endDocumentWordCount": 5,
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
    wordsAdded: 3,
    wordsRemoved: 2,
    wordsEdited: 5,
    netWordsChanged: 1,
    currentWordCount: 5
  });
});
