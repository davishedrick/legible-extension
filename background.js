(function () {
  "use strict";

  const ACE_API_BASE_URL = "https://davishedrick.pythonanywhere.com";
  const ACE_API_MESSAGE = "ace-api-fetch";
  const ACE_GOOGLE_DOC_WORD_COUNT_MESSAGE = "ace-google-doc-word-count";
  const ACE_GOOGLE_DOC_START_SNAPSHOT_MESSAGE = "ace-google-doc-start-snapshot";
  const ACE_GOOGLE_DOC_NET_COUNT_MESSAGE = "ace-google-doc-net-count";
  const ACE_CURRENT_TAB_SCOPE_MESSAGE = "ace-current-tab-scope";
  const ACE_GOOGLE_DOCS_SCOPE = "https://www.googleapis.com/auth/documents.readonly";
  const ACE_WORD_SNAPSHOT_STORAGE_PREFIX = "aceWordSnapshot:";
  const ACE_WORD_TOKENIZER_VERSION = "google-docs-like-v3";
  const ACE_GOOGLE_DOCS_SUGGESTIONS_VIEW_MODE = "PREVIEW_WITHOUT_SUGGESTIONS";
  const ACE_SCRIPTOR_SESSION_COOKIE = "session";
  const ACE_SCRIPTOR_SESSION_HEADER = "X-Scriptor-Session";

  chrome.runtime.onMessage.addListener(function (message, sender, sendResponse) {
    if (!message) {
      return false;
    }

    if (message.aceType === ACE_CURRENT_TAB_SCOPE_MESSAGE) {
      sendResponse({
        ok: true,
        chromeTabId: sender?.tab?.id ?? null,
        frameId: sender?.frameId ?? null
      });
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
            status: Number(error?.status) || 0,
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
            status: Number(error?.status) || 0,
            wordCount: null,
            error: error.message || "Google Docs start snapshot failed."
          });
        });
      return true;
    }

    if (message.aceType === ACE_GOOGLE_DOC_NET_COUNT_MESSAGE) {
      aceFetchGoogleDocNetCount(message)
        .then(sendResponse)
        .catch(function (error) {
          console.warn("[ACE] Google Docs net count failed", error);
          sendResponse({
            ok: false,
            status: Number(error?.status) || 0,
            wordCount: null,
            netWordsChanged: 0,
            error: error.message || "Google Docs net count failed."
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
    const method = options.method || "GET";
    const response = await fetch(`${ACE_API_BASE_URL}${path}`, {
      method,
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

    console.info("[ACE] APP API BRIDGE", {
      path,
      method,
      status: response.status,
      hasSessionHeader: Boolean(headers[ACE_SCRIPTOR_SESSION_HEADER]),
      ok: response.ok,
      sessionType: payload?.session?.type || "",
      netWordsChanged: payload?.session?.netWordsChanged
    });

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

    const snapshot = await aceFetchGoogleDocSnapshot(documentId, Boolean(message.interactive), message.tabId);
    console.info("[ACE] GOOGLE DOC WORD COUNT", {
      measurementPath: "google-docs-net-count",
      documentId,
      revisionId: snapshot.revisionId || "",
      apiWordCount: snapshot.wordCount
    });

    return {
      ok: true,
      status: snapshot.status,
      method: "google-docs-api",
      revisionId: snapshot.revisionId || "",
      tabId: snapshot.tabId || "",
      wordCount: snapshot.wordCount,
      wordCountTokenizerVersion: snapshot.wordCountTokenizerVersion
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

    const snapshot = await aceFetchGoogleDocSnapshot(documentId, Boolean(message.interactive), message.tabId);
    await aceStorageSet({
      [aceSnapshotStorageKey(extensionSessionId)]: {
        documentId,
        revisionId: snapshot.revisionId,
        wordCount: snapshot.wordCount,
        tabId: snapshot.tabId || "",
        wordCountTokenizerVersion: snapshot.wordCountTokenizerVersion,
        createdAt: new Date().toISOString(),
        source: "google-docs-api"
      }
    });

    console.info("[ACE] SESSION START", {
      measurementPath: "google-docs-net-count",
      documentId,
      extensionSessionId,
      revisionId: snapshot.revisionId || "",
      apiWordCount: snapshot.wordCount
    });

    return {
      ok: true,
      status: snapshot.status,
      method: "google-docs-api",
      revisionId: snapshot.revisionId || "",
      tabId: snapshot.tabId || "",
      wordCount: snapshot.wordCount,
      wordCountTokenizerVersion: snapshot.wordCountTokenizerVersion
    };
  }

  async function aceFetchGoogleDocNetCount(message) {
    const documentId = String(message.documentId || "").trim();
    const extensionSessionId = String(message.extensionSessionId || "").trim();
    if (!documentId || !extensionSessionId) {
      return {
        ok: false,
        status: 400,
        wordCount: null,
        netWordsChanged: 0,
        error: "E-NET-INPUT: Google Docs document ID and session ID are required."
      };
    }

    if (!aceGoogleOAuthConfigured()) {
      return {
        ok: false,
        status: 400,
        wordCount: null,
        netWordsChanged: 0,
        error: "E-GOOGLE-OAUTH-CONFIG: Google OAuth client ID is not configured."
      };
    }

    const stored = await aceStorageGet(aceSnapshotStorageKey(extensionSessionId));
    const startSnapshot = stored[aceSnapshotStorageKey(extensionSessionId)];
    if (!Number.isFinite(Number(startSnapshot?.wordCount)) || startSnapshot.documentId !== documentId) {
      return {
        ok: false,
        status: 404,
        wordCount: null,
        netWordsChanged: 0,
        error: "E-START-COUNT-MISSING: No start word count exists for this session. Reload the Google Doc after choosing a project, then start a new session."
      };
    }
    const requestedTabId = String(message.tabId || "").trim();
    const snapshotTabId = String(startSnapshot.tabId || "").trim();
    if (snapshotTabId && requestedTabId && snapshotTabId !== requestedTabId) {
      return {
        ok: false,
        status: 409,
        wordCount: null,
        netWordsChanged: 0,
        error: `E-START-SCOPE-MISMATCH: Session started on tab ${snapshotTabId} but completion requested tab ${requestedTabId}.`
      };
    }

    const endSnapshot = await aceFetchGoogleDocSnapshot(documentId, Boolean(message.interactive), message.tabId);
    if (message.clearSnapshot) {
      await aceStorageRemove(aceSnapshotStorageKey(extensionSessionId));
    }

    const startWordCount = Math.max(0, Number(startSnapshot.wordCount) || 0);
    const endWordCount = Math.max(0, Number(endSnapshot.wordCount) || 0);
    const netWordsChanged = endWordCount - startWordCount;
    console.info("[ACE] NET WORD COUNT", {
      measurementPath: "google-docs-net-count",
      documentId,
      extensionSessionId,
      startRevisionId: startSnapshot.revisionId || "",
      endRevisionId: endSnapshot.revisionId || "",
      startWordCount,
      endWordCount,
      netWordsChanged
    });

    return {
      ok: true,
      status: endSnapshot.status,
      method: "google-docs-api",
      revisionId: endSnapshot.revisionId || "",
      startRevisionId: startSnapshot.revisionId || "",
      tabId: endSnapshot.tabId || "",
      wordCount: endWordCount,
      startWordCount,
      netWordsChanged,
      wordCountTokenizerVersion: endSnapshot.wordCountTokenizerVersion
    };
  }

  async function aceFetchGoogleDocSnapshot(documentId, interactive, tabId = "") {
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
    const extracted = aceExtractGoogleDocTextBySource(documentPayload, tabId);
    if (extracted.requestedTabId && extracted.requestedTabId !== "default" && !extracted.tabMatched) {
      const failure = new Error("E-GOOGLE-DOC-TAB-NOT-FOUND: Active Google Docs tab was not found in the API payload.");
      failure.status = 404;
      throw failure;
    }
    const text = extracted.text;
    const wordCount = aceCountWordsInText(text);
    return {
      status: response.status,
      revisionId: documentPayload.revisionId || "",
      wordCountTokenizerVersion: ACE_WORD_TOKENIZER_VERSION,
      tabId: extracted.requestedTabId || "",
      wordCount
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

  function aceExtractGoogleDocText(documentPayload, tabId = "") {
    return aceExtractGoogleDocTextBySource(documentPayload, tabId).text;
  }

  function aceExtractGoogleDocTextBySource(documentPayload, tabId = "") {
    const chunks = [];
    const requestedTabId = String(tabId || "").trim();
    let requestedTabMatched = !requestedTabId || requestedTabId === "default";
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

    function googleDocTabId(tab) {
      return String(
        tab?.tabProperties?.tabId
        || tab?.tabProperties?.id
        || tab?.documentTab?.tabId
        || tab?.documentTab?.id
        || tab?.id
        || ""
      ).trim();
    }

    function tabMatchesRequested(tab) {
      return requestedTabId && requestedTabId !== "default" && googleDocTabId(tab) === requestedTabId;
    }

    function collectTab(tab) {
      collectDocumentTab(tab?.documentTab, "tabs");
      (tab?.childTabs || []).forEach(collectTab);
    }

    if (Array.isArray(documentPayload.tabs) && documentPayload.tabs.length) {
      if (requestedTabId && requestedTabId !== "default") {
        const matchingTabs = [];
        (function findTabs(tabs) {
          (tabs || []).forEach(function (tab) {
            if (tabMatchesRequested(tab)) {
              requestedTabMatched = true;
              matchingTabs.push(tab);
            }
            findTabs(tab?.childTabs || []);
          });
        })(documentPayload.tabs);
        matchingTabs.forEach(function (tab) {
          collectDocumentTab(tab?.documentTab, "tabs");
        });
      } else {
        documentPayload.tabs.forEach(collectTab);
      }
    } else {
      collectDocumentTab(documentPayload, "body");
    }

    return {
      text: chunks.join(" "),
      sources,
      requestedTabId,
      tabMatched: requestedTabMatched
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

  function aceCountWordsInText(text) {
    return aceWordTokensInText(text).length;
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

  if (globalThis.__ACE_TEST_EXPORTS__) {
    Object.assign(globalThis.__ACE_TEST_EXPORTS__, {
      aceStoreGoogleDocStartSnapshot,
      aceFetchFromApp,
      aceFetchGoogleDocNetCount,
      aceFetchGoogleDocSnapshot,
      aceExtractGoogleDocText,
      aceExtractGoogleDocTextBySource,
      aceWordTokensInText,
      aceCountWordsInText
    });
  }
})();
