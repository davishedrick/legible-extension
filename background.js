(function () {
  "use strict";

  const ACE_API_BASE_URL = "https://davishedrick.pythonanywhere.com";
  const ACE_API_MESSAGE = "ace-api-fetch";
  const ACE_GOOGLE_DOC_WORD_COUNT_MESSAGE = "ace-google-doc-word-count";
  const ACE_GOOGLE_DOC_START_SNAPSHOT_MESSAGE = "ace-google-doc-start-snapshot";
  const ACE_GOOGLE_DOC_DIFF_MESSAGE = "ace-google-doc-diff";
  const ACE_GOOGLE_DOCS_SCOPE = "https://www.googleapis.com/auth/documents.readonly";
  const ACE_WORD_SNAPSHOT_STORAGE_PREFIX = "aceWordSnapshot:";
  const ACE_WORD_TOKENIZER_VERSION = "google-docs-like-v3";
  const ACE_GOOGLE_DOCS_SUGGESTIONS_VIEW_MODE = "PREVIEW_WITHOUT_SUGGESTIONS";
  const ACE_SCRIPTOR_SESSION_COOKIE = "session";
  const ACE_SCRIPTOR_SESSION_HEADER = "X-Scriptor-Session";

  chrome.runtime.onMessage.addListener(function (message, _sender, sendResponse) {
    if (!message) {
      return false;
    }

    if (message.aceType === ACE_API_MESSAGE) {
      aceFetchFromApp(message)
        .then(sendResponse)
        .catch(function (error) {
          sendResponse({
            ok: false,
            status: 0,
            error: error.message || "Scriptor API request failed."
          });
        });

      return true;
    }

    if (message.aceType === ACE_GOOGLE_DOC_WORD_COUNT_MESSAGE) {
      aceFetchGoogleDocWordCount(message)
        .then(sendResponse)
        .catch(function (error) {
          console.warn("[ACE] Google Docs word count failed", error);
          sendResponse({
            ok: false,
            status: 0,
            wordCount: null,
            error: error.message || "Google Docs word count failed."
          });
        });

      return true;
    }

    if (message.aceType === ACE_GOOGLE_DOC_START_SNAPSHOT_MESSAGE) {
      aceStoreGoogleDocStartSnapshot(message)
        .then(sendResponse)
        .catch(function (error) {
          console.warn("[ACE] Google Docs start snapshot failed", error);
          sendResponse({
            ok: false,
            status: 0,
            wordCount: null,
            error: error.message || "Google Docs start snapshot failed."
          });
        });

      return true;
    }

    if (message.aceType === ACE_GOOGLE_DOC_DIFF_MESSAGE) {
      aceFetchGoogleDocWordDiff(message)
        .then(sendResponse)
        .catch(function (error) {
          console.warn("[ACE] Google Docs word diff failed", error);
          sendResponse({
            ok: false,
            status: 0,
            wordCount: null,
            wordsAdded: 0,
            wordsRemoved: 0,
            error: error.message || "Google Docs word diff failed."
          });
        });

      return true;
    }

    return false;
  });

  async function aceFetchFromApp(message) {
    const path = String(message.path || "");
    if (!path.startsWith("/api/")) {
      return {
        ok: false,
        status: 400,
        error: "Invalid Scriptor API path."
      };
    }

    const options = message.options || {};
    const headers = {
      "Content-Type": "application/json",
      ...(options.headers || {})
    };
    const sessionCookie = await aceGetScriptorSessionCookie();
    if (sessionCookie && !headers[ACE_SCRIPTOR_SESSION_HEADER]) {
      headers[ACE_SCRIPTOR_SESSION_HEADER] = sessionCookie;
    }
    const response = await fetch(`${ACE_API_BASE_URL}${path}`, {
      method: options.method || "GET",
      credentials: "include",
      headers,
      body: options.body || undefined
    });

    const text = await response.text();
    let payload = {};

    if (text) {
      try {
        payload = JSON.parse(text);
      } catch (_error) {
        payload = { raw: text };
      }
    }

    return {
      ok: response.ok,
      status: response.status,
      payload,
      error: payload.error || response.statusText || "Scriptor API request failed."
    };
  }

  function aceGetScriptorSessionCookie() {
    return new Promise(function (resolve) {
      if (!chrome.cookies?.get) {
        resolve("");
        return;
      }
      chrome.cookies.get(
        {
          url: ACE_API_BASE_URL,
          name: ACE_SCRIPTOR_SESSION_COOKIE
        },
        function (cookie) {
          resolve(cookie?.value || "");
        }
      );
    });
  }

  async function aceFetchGoogleDocWordCount(message) {
    const documentId = String(message.documentId || "").trim();
    if (!documentId) {
      return {
        ok: false,
        status: 400,
        wordCount: null,
        error: "E-DOC-ID-MISSING: Google Docs document ID is required."
      };
    }

    if (!aceGoogleOAuthConfigured()) {
      return {
        ok: false,
        status: 400,
        wordCount: null,
        error: "E-GOOGLE-OAUTH-CONFIG: Google OAuth client ID is not configured."
      };
    }

    const snapshot = await aceFetchGoogleDocSnapshot(documentId, Boolean(message.interactive));
    console.info("[ACE] SESSION END", {
      measurementPath: "exact-api-sequence-diff",
      documentId,
      revisionId: snapshot.revisionId || "",
      apiWordCount: snapshot.wordCount,
      tokenCount: snapshot.wordTokens.length
    });

    return {
      ok: true,
      status: snapshot.status,
      method: "google-docs-api",
      revisionId: snapshot.revisionId || "",
      wordCount: snapshot.wordCount,
      wordCountTokenizerVersion: snapshot.wordCountTokenizerVersion,
      wordCounts: snapshot.wordCounts,
      wordTokens: snapshot.wordTokens
    };
  }

  async function aceStoreGoogleDocStartSnapshot(message) {
    const documentId = String(message.documentId || "").trim();
    const extensionSessionId = String(message.extensionSessionId || "").trim();
    if (!documentId || !extensionSessionId) {
      return {
        ok: false,
        status: 400,
        wordCount: null,
        error: "E-SNAPSHOT-INPUT: Google Docs document ID and session ID are required."
      };
    }

    if (!aceGoogleOAuthConfigured()) {
      return {
        ok: false,
        status: 400,
        wordCount: null,
        error: "E-GOOGLE-OAUTH-CONFIG: Google OAuth client ID is not configured."
      };
    }

    const snapshot = await aceFetchGoogleDocSnapshot(documentId, Boolean(message.interactive));
    await aceStorageSet({
      [aceSnapshotStorageKey(extensionSessionId)]: {
        documentId,
        revisionId: snapshot.revisionId,
        wordCount: snapshot.wordCount,
        wordCounts: snapshot.wordCounts,
        wordTokens: snapshot.wordTokens,
        wordCountTokenizerVersion: snapshot.wordCountTokenizerVersion,
        createdAt: new Date().toISOString()
      }
    });

    console.info("[ACE] SESSION START", {
      measurementPath: "exact-api-sequence-diff",
      documentId,
      extensionSessionId,
      revisionId: snapshot.revisionId || "",
      apiWordCount: snapshot.wordCount,
      tokenCount: snapshot.wordTokens.length,
      wordCountTokenizerVersion: snapshot.wordCountTokenizerVersion
    });

    return {
      ok: true,
      status: snapshot.status,
      method: "google-docs-api",
      revisionId: snapshot.revisionId || "",
      wordCount: snapshot.wordCount,
      wordCountTokenizerVersion: snapshot.wordCountTokenizerVersion,
      wordCounts: snapshot.wordCounts,
      wordTokens: snapshot.wordTokens
    };
  }

  async function aceFetchGoogleDocWordDiff(message) {
    const documentId = String(message.documentId || "").trim();
    const extensionSessionId = String(message.extensionSessionId || "").trim();
    if (!documentId || !extensionSessionId) {
      return {
        ok: false,
        status: 400,
        wordCount: null,
        wordsAdded: 0,
        wordsRemoved: 0,
        error: "E-DIFF-INPUT: Google Docs document ID and session ID are required."
      };
    }

    if (!aceGoogleOAuthConfigured()) {
      return {
        ok: false,
        status: 400,
        wordCount: null,
        wordsAdded: 0,
        wordsRemoved: 0,
        error: "E-GOOGLE-OAUTH-CONFIG: Google OAuth client ID is not configured."
      };
    }

    const stored = await aceStorageGet(aceSnapshotStorageKey(extensionSessionId));
    const startSnapshot = stored[aceSnapshotStorageKey(extensionSessionId)];
    if (startSnapshot?.source === "visible-total-baseline") {
      return {
        ok: false,
        status: 409,
        wordCount: null,
        wordsAdded: 0,
        wordsRemoved: 0,
        error: "E-VISIBLE-BASELINE: Session started from the visible Google Docs count because the extension context/API was unavailable. End measurement will use the visible count fallback if the API diff is unavailable."
      };
    }

    if (!startSnapshot?.wordTokens || startSnapshot.documentId !== documentId) {
      return {
        ok: false,
        status: 404,
        wordCount: null,
        wordsAdded: 0,
        wordsRemoved: 0,
        error: "E-SNAPSHOT-MISSING: No Google API before token snapshot exists for this session. Reload the Google Doc after choosing a project, wait two seconds, then start a new session."
      };
    }

    if (startSnapshot.wordCountTokenizerVersion !== ACE_WORD_TOKENIZER_VERSION) {
      return {
        ok: false,
        status: 409,
        wordCount: null,
        wordsAdded: 0,
        wordsRemoved: 0,
        error: `E-TOKENIZER-VERSION: The before snapshot used ${startSnapshot.wordCountTokenizerVersion || "an old tokenizer"}, but this extension uses ${ACE_WORD_TOKENIZER_VERSION}. Start a new session so added/removed words can be compared accurately.`
      };
    }

    const endSnapshot = await aceFetchGoogleDocSnapshot(documentId, Boolean(message.interactive));
    const diff = aceCompareWordTokens(startSnapshot.wordTokens, endSnapshot.wordTokens);

    if (message.clearSnapshot) {
      await aceStorageRemove(aceSnapshotStorageKey(extensionSessionId));
    }

    console.info("[ACE] DIFF RESULT", {
      measurementPath: diff.method || "exact-api-sequence-diff",
      documentId,
      extensionSessionId,
      startRevisionId: startSnapshot.revisionId || "",
      endRevisionId: endSnapshot.revisionId || "",
      wordsAdded: diff.wordsAdded,
      wordsRemoved: diff.wordsRemoved,
      wordsEdited: diff.wordsAdded + diff.wordsRemoved,
      netWordsChanged: endSnapshot.wordCount - startSnapshot.wordCount
    });

    return {
      ok: true,
      status: endSnapshot.status,
      method: "google-docs-api",
      revisionId: endSnapshot.revisionId || "",
      startRevisionId: startSnapshot.revisionId || "",
      wordCount: endSnapshot.wordCount,
      startWordCount: startSnapshot.wordCount,
      endWordCounts: endSnapshot.wordCounts,
      endWordTokens: endSnapshot.wordTokens,
      wordCountTokenizerVersion: endSnapshot.wordCountTokenizerVersion,
      wordsAdded: diff.wordsAdded,
      wordsRemoved: diff.wordsRemoved,
      netWordsChanged: endSnapshot.wordCount - startSnapshot.wordCount,
      wordDiffMethod: diff.method
    };
  }

  async function aceFetchGoogleDocSnapshot(documentId, interactive) {
    const token = await aceGetGoogleAuthToken(interactive);
    const params = new URLSearchParams({
      includeTabsContent: "true",
      suggestionsViewMode: ACE_GOOGLE_DOCS_SUGGESTIONS_VIEW_MODE
    });
    const response = await fetch(
      `https://docs.googleapis.com/v1/documents/${encodeURIComponent(documentId)}?${params.toString()}`,
      {
        headers: {
          Authorization: `Bearer ${token}`
        }
      }
    );

    if (response.status === 401 || response.status === 403) {
      await aceRemoveCachedGoogleAuthToken(token);
    }

    if (!response.ok) {
      let error = response.statusText || "Google Docs word count failed.";
      try {
        const payload = await response.json();
        error = payload.error?.message || error;
      } catch (_error) {
        // Keep the HTTP status text when Google does not return JSON.
      }

      const failure = new Error(`E-GOOGLE-API-${response.status}: ${error}`);
      failure.status = response.status;
      throw failure;
    }

    const documentPayload = await response.json();
    const extraction = aceExtractGoogleDocTextBySource(documentPayload);
    const wordTokens = aceWordTokensInText(extraction.text);
    const wordCounts = aceWordCountsFromTokens(wordTokens);
    console.info("[ACE] API TEXT DIAGNOSTIC", {
      documentId,
      revisionId: documentPayload.revisionId || "",
      suggestionsViewMode: ACE_GOOGLE_DOCS_SUGGESTIONS_VIEW_MODE,
      extractedText: extraction.text,
      tokenCount: wordTokens.length,
      fullTokens: wordTokens,
      first150Tokens: wordTokens.slice(0, 150),
      final100Tokens: wordTokens.slice(-100),
      last150Tokens: wordTokens.slice(-150),
      suspiciousUnicodeRanges: aceSuspiciousUnicodeRanges(extraction.text),
      sourceTokenCounts: aceSourceTokenCounts(extraction.sources),
      sourceTextSamples: aceSourceTextSamples(extraction.sources),
      duplicateTextHashes: aceDuplicateTextHashes(extraction.sources)
    });
    return {
      status: response.status,
      revisionId: documentPayload.revisionId || "",
      wordCountTokenizerVersion: ACE_WORD_TOKENIZER_VERSION,
      wordTokens,
      wordCounts,
      wordCount: wordTokens.length
    };
  }

  function aceGoogleOAuthConfigured() {
    const clientId = chrome.runtime.getManifest().oauth2?.client_id || "";
    return Boolean(clientId) && !clientId.includes("REPLACE_WITH_GOOGLE_OAUTH_CLIENT_ID");
  }

  function aceGetGoogleAuthToken(interactive) {
    return new Promise(function (resolve, reject) {
      chrome.identity.getAuthToken(
        {
          interactive,
          scopes: [ACE_GOOGLE_DOCS_SCOPE]
        },
        function (token) {
          if (chrome.runtime.lastError || !token) {
            reject(new Error(`E-GOOGLE-OAUTH: ${chrome.runtime.lastError?.message || "Google sign-in was not completed."}`));
            return;
          }

          resolve(token);
        }
      );
    });
  }

  function aceRemoveCachedGoogleAuthToken(token) {
    return new Promise(function (resolve) {
      chrome.identity.removeCachedAuthToken({ token }, resolve);
    });
  }

  function aceStorageGet(keys) {
    return new Promise(function (resolve) {
      chrome.storage.local.get(keys, function (value) {
        resolve(chrome.runtime.lastError ? {} : value);
      });
    });
  }

  function aceStorageSet(values) {
    return new Promise(function (resolve, reject) {
      chrome.storage.local.set(values, function () {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
          return;
        }

        resolve();
      });
    });
  }

  function aceStorageRemove(keys) {
    return new Promise(function (resolve) {
      chrome.storage.local.remove(keys, resolve);
    });
  }

  function aceSnapshotStorageKey(extensionSessionId) {
    return `${ACE_WORD_SNAPSHOT_STORAGE_PREFIX}${extensionSessionId}`;
  }

  function aceExtractGoogleDocText(documentPayload) {
    return aceExtractGoogleDocTextBySource(documentPayload).text;
  }

  function aceExtractGoogleDocTextBySource(documentPayload) {
    const chunks = [];
    const sources = {
      body: [],
      tabs: [],
      headers: [],
      footers: [],
      footnotes: [],
      tables: [],
      tableOfContents: [],
      suggestions: [],
      namedRanges: [],
      inlineObjects: [],
      other: []
    };

    function pushSource(source, content) {
      if (content) {
        sources[source].push(content);
        chunks.push(content);
      }
    }

    function recordSource(source, content) {
      if (content) {
        sources[source].push(content);
      }
    }

    function collectTextRun(paragraphElement, source) {
      const textRun = paragraphElement?.textRun;
      const content = textRun?.content;
      if (!content) {
        return;
      }
      if (Array.isArray(textRun.suggestedInsertionIds) && textRun.suggestedInsertionIds.length) {
        recordSource("suggestions", content);
        return;
      }
      pushSource(source, content);
    }

    function collectStructuralElement(element, source) {
      if (element?.tableOfContents) {
        recordSource("tableOfContents", aceTextFromStructuralElements(element.tableOfContents.content || []));
        return;
      }

      if (element?.paragraph?.elements) {
        element.paragraph.elements.forEach(function (paragraphElement) {
          collectTextRun(paragraphElement, source);
        });
      }

      if (element?.table?.tableRows) {
        element.table.tableRows.forEach(function (row) {
          (row.tableCells || []).forEach(function (cell) {
            collectStructuralElements(cell.content || [], "tables");
          });
        });
      }

      // Google Docs' visible manuscript word count excludes generated table-of-contents text.
    }

    function collectStructuralElements(elements, source) {
      (elements || []).forEach(function (element) {
        collectStructuralElement(element, source);
      });
    }

    function collectDocumentTab(documentTab, source) {
      if (!documentTab) {
        return;
      }

      collectStructuralElements(documentTab.body?.content || [], source);
      Object.values(documentTab.headers || {}).forEach(function (header) {
        recordSource("headers", aceTextFromStructuralElements(header.content || []));
      });
      Object.values(documentTab.footers || {}).forEach(function (footer) {
        recordSource("footers", aceTextFromStructuralElements(footer.content || []));
      });
      Object.values(documentTab.footnotes || {}).forEach(function (footnote) {
        recordSource("footnotes", aceTextFromStructuralElements(footnote.content || []));
      });
    }

    function collectTab(tab) {
      collectDocumentTab(tab?.documentTab, "tabs");
      (tab?.childTabs || []).forEach(collectTab);
    }

    if (Array.isArray(documentPayload.tabs) && documentPayload.tabs.length) {
      documentPayload.tabs.forEach(collectTab);
    } else {
      collectDocumentTab(documentPayload, "body");
    }

    return {
      text: chunks.join(" "),
      sources
    };
  }

  function aceTextFromStructuralElements(elements) {
    const chunks = [];
    (elements || []).forEach(function collect(element) {
      if (element?.paragraph?.elements) {
        element.paragraph.elements.forEach(function (paragraphElement) {
          const content = paragraphElement?.textRun?.content;
          if (content) {
            chunks.push(content);
          }
        });
      }
      if (element?.table?.tableRows) {
        element.table.tableRows.forEach(function (row) {
          (row.tableCells || []).forEach(function (cell) {
            (cell.content || []).forEach(collect);
          });
        });
      }
    });
    return chunks.join(" ");
  }

  function aceSourceTokenCounts(sources) {
    return Object.fromEntries(Object.entries(sources).map(function ([source, chunks]) {
      return [source, aceWordTokensInText((chunks || []).join(" ")).length];
    }));
  }

  function aceSourceTextSamples(sources) {
    return Object.fromEntries(Object.entries(sources).map(function ([source, chunks]) {
      const text = (chunks || []).join(" ").trim();
      return [source, text ? text.slice(0, 500) : ""];
    }));
  }

  function aceDuplicateTextHashes(sources) {
    const seen = new Map();
    Object.entries(sources || {}).forEach(function ([source, chunks]) {
      (chunks || []).forEach(function (chunk, index) {
        const text = String(chunk || "").trim();
        if (!text) {
          return;
        }
        const hash = aceSimpleHash(text);
        const item = seen.get(hash) || {
          hash,
          tokenCount: aceWordTokensInText(text).length,
          occurrences: []
        };
        item.occurrences.push({ source, index, sample: text.slice(0, 120) });
        seen.set(hash, item);
      });
    });
    return Array.from(seen.values()).filter(function (item) {
      return item.occurrences.length > 1;
    });
  }

  function aceSimpleHash(text) {
    let hash = 2166136261;
    for (let index = 0; index < text.length; index += 1) {
      hash ^= text.charCodeAt(index);
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16);
  }

  function aceSuspiciousUnicodeRanges(text) {
    const value = String(text || "");
    const ranges = [];
    for (let index = 0; index < value.length; index += 1) {
      const codePoint = value.codePointAt(index);
      const character = String.fromCodePoint(codePoint);
      const name = aceSuspiciousUnicodeName(codePoint, character);
      if (codePoint > 0xffff) {
        index += 1;
      }
      if (!name) {
        continue;
      }
      const start = Math.max(0, index - 24);
      const end = Math.min(value.length, index + 25);
      const context = value.slice(start, end);
      ranges.push({
        index,
        codePoint: `U+${codePoint.toString(16).toUpperCase().padStart(4, "0")}`,
        name,
        context,
        contextCodePoints: Array.from(context).map(function (item) {
          const itemCodePoint = item.codePointAt(0);
          return `U+${itemCodePoint.toString(16).toUpperCase().padStart(4, "0")}`;
        })
      });
    }
    return ranges;
  }

  function aceSuspiciousUnicodeName(codePoint, character) {
    const names = {
      0x00ad: "SOFT HYPHEN",
      0x00a0: "NO-BREAK SPACE",
      0x2007: "FIGURE SPACE",
      0x2009: "THIN SPACE",
      0x202f: "NARROW NO-BREAK SPACE",
      0x3000: "IDEOGRAPHIC SPACE",
      0x200b: "ZERO WIDTH SPACE",
      0x200c: "ZERO WIDTH NON-JOINER",
      0x200d: "ZERO WIDTH JOINER",
      0x2060: "WORD JOINER",
      0xfeff: "ZERO WIDTH NO-BREAK SPACE"
    };
    if (names[codePoint]) {
      return names[codePoint];
    }
    if (/\p{M}/u.test(character)) {
      return "COMBINING MARK";
    }
    if (/\p{Script=Greek}/u.test(character)) {
      return "GREEK CHARACTER";
    }
    if (/\p{Script=Cyrillic}/u.test(character)) {
      return "CYRILLIC CHARACTER";
    }
    return "";
  }

  function aceCountWordsInText(text) {
    return aceWordTokensInText(text).length;
  }

  function aceWordCountsInText(text) {
    return aceWordCountsFromTokens(aceWordTokensInText(text));
  }

  function aceWordTokensInText(text) {
    const textForCounting = String(text || "")
      .normalize("NFKC")
      .replace(/[\u00ad\u034f\u061c\u115f\u1160\u17b4\u17b5\u180b-\u180f\u200b-\u200f\u202a-\u202e\u2060-\u206f\u3164\ufe00-\ufe0f\ufeff\uffa0]/g, "")
      .replace(/\u00a0/g, " ");
    const matches = textForCounting.match(/[\p{L}\p{N}][\p{L}\p{N}\p{M}]*(?:(?:['’]|[-‐‑‒–—])(?=[\p{L}\p{N}])[\p{L}\p{N}\p{M}]*)*/gu) || [];
    return matches
      .map(function (token) {
        return token.toLocaleLowerCase();
      })
      .filter(function (token) {
        return Boolean(token) && /\p{L}/u.test(token);
      });
  }

  function aceWordCountsFromTokens(tokens) {
    const counts = {};
    (tokens || []).forEach(function (token) {
      counts[token] = (counts[token] || 0) + 1;
    });
    return counts;
  }

  function aceTotalWordCounts(wordCounts) {
    return Object.values(wordCounts || {}).reduce(function (total, count) {
      return total + Math.max(0, Number(count) || 0);
    }, 0);
  }

  function aceCompareWordCounts(startWordCounts, endWordCounts) {
    const keys = new Set([
      ...Object.keys(startWordCounts || {}),
      ...Object.keys(endWordCounts || {})
    ]);
    let wordsAdded = 0;
    let wordsRemoved = 0;

    keys.forEach(function (key) {
      const startCount = Math.max(0, Number(startWordCounts[key]) || 0);
      const endCount = Math.max(0, Number(endWordCounts[key]) || 0);
      if (endCount > startCount) {
        wordsAdded += endCount - startCount;
      } else if (startCount > endCount) {
        wordsRemoved += startCount - endCount;
      }
    });

    return { wordsAdded, wordsRemoved };
  }

  function aceCompareWordTokens(startTokens, endTokens) {
    const before = Array.isArray(startTokens) ? startTokens : [];
    const after = Array.isArray(endTokens) ? endTokens : [];
    let prefixLength = 0;
    while (
      prefixLength < before.length
      && prefixLength < after.length
      && before[prefixLength] === after[prefixLength]
    ) {
      prefixLength += 1;
    }

    let beforeEnd = before.length;
    let afterEnd = after.length;
    while (
      beforeEnd > prefixLength
      && afterEnd > prefixLength
      && before[beforeEnd - 1] === after[afterEnd - 1]
    ) {
      beforeEnd -= 1;
      afterEnd -= 1;
    }

    const removedTokens = before.slice(prefixLength, beforeEnd);
    const addedTokens = after.slice(prefixLength, afterEnd);
    if (!removedTokens.length || !addedTokens.length) {
      return {
        wordsAdded: addedTokens.length,
        wordsRemoved: removedTokens.length,
        method: "google-api-token-sequence"
      };
    }

    const cellCount = removedTokens.length * addedTokens.length;
    if (cellCount > 2000000) {
      return {
        ...aceCompareWordCounts(
          aceWordCountsFromTokens(removedTokens),
          aceWordCountsFromTokens(addedTokens)
        ),
        method: "google-api-token-map-fallback"
      };
    }

    const lcsLength = aceLongestCommonSubsequenceLength(removedTokens, addedTokens);
    return {
      wordsAdded: addedTokens.length - lcsLength,
      wordsRemoved: removedTokens.length - lcsLength,
      method: "google-api-token-sequence"
    };
  }

  function aceLongestCommonSubsequenceLength(before, after) {
    let previous = new Uint32Array(after.length + 1);
    for (let beforeIndex = 0; beforeIndex < before.length; beforeIndex += 1) {
      const current = new Uint32Array(after.length + 1);
      for (let afterIndex = 0; afterIndex < after.length; afterIndex += 1) {
        current[afterIndex + 1] = before[beforeIndex] === after[afterIndex]
          ? previous[afterIndex] + 1
          : Math.max(previous[afterIndex + 1], current[afterIndex]);
      }
      previous = current;
    }
    return previous[after.length];
  }

  if (globalThis.__ACE_TEST_EXPORTS__) {
    Object.assign(globalThis.__ACE_TEST_EXPORTS__, {
      aceStoreGoogleDocStartSnapshot,
      aceFetchFromApp,
      aceFetchGoogleDocWordDiff,
      aceFetchGoogleDocSnapshot,
      aceExtractGoogleDocText,
      aceExtractGoogleDocTextBySource,
      aceWordTokensInText,
      aceWordCountsInText,
      aceSourceTokenCounts,
      aceCompareWordTokens,
      aceCompareWordCounts
    });
  }
})();
