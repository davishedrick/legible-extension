const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const { loadBackground, makeGoogleDoc } = require("./helpers");

const startFixture = JSON.parse(fs.readFileSync(path.resolve(__dirname, "fixtures", "googleDocStart.json"), "utf8"));
const endFixture = JSON.parse(fs.readFileSync(path.resolve(__dirname, "fixtures", "googleDocEnd.json"), "utf8"));

function localArray(value) {
  return Array.from(value || []);
}

test("canonical case #1 computes +3 / -2 / net +1 from Google Docs token snapshots", async () => {
  const { exports } = loadBackground([startFixture, endFixture]);

  const start = await exports.aceStoreGoogleDocStartSnapshot({
    documentId: "doc-test",
    extensionSessionId: "session-test",
    interactive: false
  });
  const diff = await exports.aceFetchGoogleDocWordDiff({
    documentId: "doc-test",
    extensionSessionId: "session-test",
    interactive: false,
    clearSnapshot: false
  });

  assert.equal(start.wordCount, 4);
  assert.deepEqual(localArray(start.wordTokens), ["apple", "banana", "dog", "friend"]);
  assert.equal(diff.wordDiffMethod, "google-api-token-sequence");
  assert.equal(diff.wordsAdded, 3);
  assert.equal(diff.wordsRemoved, 2);
  assert.equal(diff.wordsAdded + diff.wordsRemoved, 5);
  assert.equal(diff.netWordsChanged, 1);
});

test("canonical case #2 computes +100 / -100 / net 0", async () => {
  const stablePrefix = Array.from({ length: 450 }, (_, index) => `keep${index}`);
  const stableSuffix = Array.from({ length: 450 }, (_, index) => `tail${index}`);
  const removed = Array.from({ length: 100 }, (_, index) => `removed${index}`);
  const added = Array.from({ length: 100 }, (_, index) => `added${index}`);
  const startText = [...stablePrefix, ...removed, ...stableSuffix].join(" ");
  const endText = [...stablePrefix, ...added, ...stableSuffix].join(" ");
  const { exports } = loadBackground([
    makeGoogleDoc(startText, "start-1000"),
    makeGoogleDoc(endText, "end-1000")
  ]);

  await exports.aceStoreGoogleDocStartSnapshot({
    documentId: "doc-test",
    extensionSessionId: "session-test",
    interactive: false
  });
  const diff = await exports.aceFetchGoogleDocWordDiff({
    documentId: "doc-test",
    extensionSessionId: "session-test",
    interactive: false,
    clearSnapshot: false
  });

  assert.equal(diff.wordDiffMethod, "google-api-token-sequence");
  assert.equal(diff.wordCount, 1000);
  assert.equal(diff.startWordCount, 1000);
  assert.equal(diff.wordsAdded, 100);
  assert.equal(diff.wordsRemoved, 100);
  assert.equal(diff.wordsAdded + diff.wordsRemoved, 200);
  assert.equal(diff.netWordsChanged, 0);
});

test("draft-only words deleted before session end do not count as removals", async () => {
  const finalTokens = Array.from({ length: 600 }, (_, index) => `survivor${index}`);
  const { exports } = loadBackground([
    makeGoogleDoc("", "start-empty"),
    makeGoogleDoc(finalTokens.join(" "), "end-600")
  ]);

  await exports.aceStoreGoogleDocStartSnapshot({
    documentId: "doc-test",
    extensionSessionId: "session-test",
    interactive: false
  });
  const diff = await exports.aceFetchGoogleDocWordDiff({
    documentId: "doc-test",
    extensionSessionId: "session-test",
    interactive: false,
    clearSnapshot: false
  });

  assert.equal(diff.wordDiffMethod, "google-api-token-sequence");
  assert.equal(diff.startWordCount, 0);
  assert.equal(diff.wordCount, 600);
  assert.equal(diff.wordsAdded, 600);
  assert.equal(diff.wordsRemoved, 0);
  assert.equal(diff.wordsAdded + diff.wordsRemoved, 600);
  assert.equal(diff.netWordsChanged, 600);
});

test("baseline deletions count, deleted draft words do not", async () => {
  const stablePrefix = Array.from({ length: 450 }, (_, index) => `keep${index}`);
  const stableSuffix = Array.from({ length: 450 }, (_, index) => `tail${index}`);
  const removedBaseline = Array.from({ length: 100 }, (_, index) => `baseline${index}`);
  const survivingDraft = Array.from({ length: 150 }, (_, index) => `draft${index}`);
  const startText = [...stablePrefix, ...removedBaseline, ...stableSuffix].join(" ");
  const endText = [...stablePrefix, ...survivingDraft, ...stableSuffix].join(" ");
  const { exports } = loadBackground([
    makeGoogleDoc(startText, "start-1000"),
    makeGoogleDoc(endText, "end-1050")
  ]);

  await exports.aceStoreGoogleDocStartSnapshot({
    documentId: "doc-test",
    extensionSessionId: "session-test",
    interactive: false
  });
  const diff = await exports.aceFetchGoogleDocWordDiff({
    documentId: "doc-test",
    extensionSessionId: "session-test",
    interactive: false,
    clearSnapshot: false
  });

  assert.equal(diff.wordDiffMethod, "google-api-token-sequence");
  assert.equal(diff.startWordCount, 1000);
  assert.equal(diff.wordCount, 1050);
  assert.equal(diff.wordsAdded, 150);
  assert.equal(diff.wordsRemoved, 100);
  assert.equal(diff.wordsAdded + diff.wordsRemoved, 250);
  assert.equal(diff.netWordsChanged, 50);
});

test("case-only changes normalize away", () => {
  const { exports } = loadBackground();
  const diff = exports.aceCompareWordTokens(
    exports.aceWordTokensInText("apple"),
    exports.aceWordTokensInText("Apple")
  );

  assert.deepEqual(localArray(exports.aceWordTokensInText("Apple")), ["apple"]);
  assert.equal(diff.wordsAdded, 0);
  assert.equal(diff.wordsRemoved, 0);
});

test("punctuation and symbols split into Google Docs-like tokens", () => {
  const { exports } = loadBackground();
  assert.deepEqual(
    localArray(exports.aceWordTokensInText("test_app.py author-companion davishedrick_pythonanywhere_com_wsgi.py")),
    [
      "test",
      "app",
      "py",
      "author",
      "companion",
      "davishedrick",
      "pythonanywhere",
      "com",
      "wsgi",
      "py"
    ]
  );
});
