const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const { loadBackground, makeGoogleDoc } = require("./helpers");

function localArray(value) {
  return Array.from(value || []);
}

test("Google Docs snapshot flow returns net word change only", async () => {
  const { exports } = loadBackground([
    makeGoogleDoc("apple banana dog friend", "start"),
    makeGoogleDoc("Apple banana pie face token", "end")
  ]);

  const start = await exports.aceStoreGoogleDocStartSnapshot({
    documentId: "doc-test",
    extensionSessionId: "session-test",
    interactive: false
  });
  const result = await exports.aceFetchGoogleDocNetCount({
    documentId: "doc-test",
    extensionSessionId: "session-test",
    interactive: false,
    clearSnapshot: false
  });

  assert.equal(start.wordCount, 4);
  assert.equal(result.startWordCount, 4);
  assert.equal(result.wordCount, 5);
  assert.equal(result.netWordsChanged, 1);
  assert.equal(result.wordsAdded, undefined);
  assert.equal(result.wordsRemoved, undefined);
});

test("net count handles positive, zero, and negative changes", async () => {
  for (const [startCount, endCount, expectedNet] of [
    [0, 600, 600],
    [1000, 1000, 0],
    [1000, 800, -200]
  ]) {
    const startText = Array.from({ length: startCount }, (_, index) => `start${index}`).join(" ");
    const endText = Array.from({ length: endCount }, (_, index) => `end${index}`).join(" ");
    const { exports } = loadBackground([
      makeGoogleDoc(startText, `start-${startCount}`),
      makeGoogleDoc(endText, `end-${endCount}`)
    ]);

    await exports.aceStoreGoogleDocStartSnapshot({
      documentId: "doc-test",
      extensionSessionId: "session-test",
      interactive: false
    });
    const result = await exports.aceFetchGoogleDocNetCount({
      documentId: "doc-test",
      extensionSessionId: "session-test",
      interactive: false,
      clearSnapshot: false
    });

    assert.equal(result.startWordCount, startCount);
    assert.equal(result.wordCount, endCount);
    assert.equal(result.netWordsChanged, expectedNet);
  }
});

test("case normalization keeps visible word count stable", () => {
  const { exports } = loadBackground();

  assert.deepEqual(localArray(exports.aceWordTokensInText("apple Apple APPLE")), [
    "apple",
    "apple",
    "apple"
  ]);
});

test("punctuation and symbols split into Google Docs-like count tokens", () => {
  const { exports } = loadBackground();
  assert.deepEqual(
    localArray(exports.aceWordTokensInText("test_app.py author-companion davishedrick_pythonanywhere_com_wsgi.py")),
    [
      "test",
      "app",
      "py",
      "author-companion",
      "davishedrick",
      "pythonanywhere",
      "com",
      "wsgi",
      "py"
    ]
  );
});

test("Unicode fixture matches Google Docs visible word count", () => {
  const { exports } = loadBackground();
  const text = fs.readFileSync(path.resolve(__dirname, "fixtures", "mismatch_192_visible_220_api_unicode.txt"), "utf8");
  const tokens = localArray(exports.aceWordTokensInText(text));

  assert.equal(tokens.length, 192);
  assert.equal(tokens[5], "consectetur");
  assert(tokens.includes("countcount"));
  assert(tokens.includes("wordword"));
  assert(tokens.includes("gammadelta"));
  assert(tokens.includes("endofstresstestdocument"));
  assert(!tokens.includes("123456"));
  assert(!tokens.includes("789"));
});

test("zero-width format controls do not split visible words", () => {
  const { exports } = loadBackground();

  assert.deepEqual(localArray(exports.aceWordTokensInText("word\u200bword")), ["wordword"]);
  assert.deepEqual(localArray(exports.aceWordTokensInText("word\u200cword")), ["wordword"]);
  assert.deepEqual(localArray(exports.aceWordTokensInText("word\u200dword")), ["wordword"]);
  assert.deepEqual(localArray(exports.aceWordTokensInText("word\u2060word")), ["wordword"]);
  assert.deepEqual(localArray(exports.aceWordTokensInText("word\ufeffword")), ["wordword"]);
});

test("non-breaking and Unicode spaces behave as word separators", () => {
  const { exports } = loadBackground();

  assert.deepEqual(localArray(exports.aceWordTokensInText("word\u00a0word")), ["word", "word"]);
  assert.deepEqual(
    localArray(exports.aceWordTokensInText("figure\u2007thin\u2009narrow\u202fideographic\u3000space")),
    ["figure", "thin", "narrow", "ideographic", "space"]
  );
});

test("soft hyphen and combining marks remain inside visible words", () => {
  const { exports } = loadBackground();

  assert.deepEqual(localArray(exports.aceWordTokensInText("soft\u00adhyphen")), ["softhyphen"]);
  assert.deepEqual(localArray(exports.aceWordTokensInText("café cafe\u0301")), ["café", "café"]);
  assert.deepEqual(localArray(exports.aceWordTokensInText("a\u0338rcu e\u0337get Bullet\u0300")), ["a̸rcu", "e̷get", "bullet̀"]);
});

test("numbers alone do not count as manuscript words, but alphanumeric words do", () => {
  const { exports } = loadBackground();

  assert.deepEqual(localArray(exports.aceWordTokensInText("123456 789\u200c012 345\u200d678")), []);
  assert.deepEqual(localArray(exports.aceWordTokensInText("keep0 draft42")), ["keep0", "draft42"]);
});

test("tabs/body duplicate fixture counts tab content once", async () => {
  const duplicateText = "apple banana dog friend";
  const payload = {
    revisionId: "tabs-1",
    body: {
      content: [
        {
          paragraph: {
            elements: [{ textRun: { content: `${duplicateText}\n` } }]
          }
        }
      ]
    },
    tabs: [
      {
        documentTab: {
          body: {
            content: [
              {
                paragraph: {
                  elements: [{ textRun: { content: `${duplicateText}\n` } }]
                }
              }
            ]
          }
        }
      }
    ]
  };
  const { exports } = loadBackground([payload]);
  const snapshot = await exports.aceFetchGoogleDocSnapshot("doc-test", false);

  assert.equal(snapshot.wordCount, 4);
});

test("suggested insertions are not counted in the final visible text", async () => {
  const payload = {
    revisionId: "suggestions-1",
    body: {
      content: [
        {
          paragraph: {
            elements: [
              { textRun: { content: "visible words\n" } },
              { textRun: { content: "suggested hidden\n", suggestedInsertionIds: ["suggestion-1"] } }
            ]
          }
        }
      ]
    }
  };
  const { exports } = loadBackground([payload]);
  const snapshot = await exports.aceFetchGoogleDocSnapshot("doc-test", false);

  assert.equal(snapshot.wordCount, 2);
});
