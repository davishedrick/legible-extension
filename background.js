(function () {
  "use strict";

  const ACE_API_BASE_URL = "https://davishedrick.pythonanywhere.com";
  const ACE_API_MESSAGE = "ace-api-fetch";
  const ACE_GOOGLE_DOC_WORD_COUNT_MESSAGE = "ace-google-doc-word-count";
  const ACE_GOOGLE_DOC_START_SNAPSHOT_MESSAGE = "ace-google-doc-start-snapshot";
  const ACE_GOOGLE_DOC_DIFF_MESSAGE = "ace-google-doc-diff";
  const ACE_GOOGLE_DOCS_SCOPE = "https://www.googleapis.com/auth/documents.readonly";
  const ACE_WORD_SNAPSHOT_STORAGE_PREFIX = "aceWordSnapshot:";

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
            error: error.message || "Author Companion API request failed."
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
        error: "Invalid Author Companion API path."
      };
    }

    const options = message.options || {};
    const response = await fetch(`${ACE_API_BASE_URL}${path}`, {
      method: options.method || "GET",
      credentials: "include",
      headers: {
        "Content-Type": "application/json",
        ...(options.headers || {})
      },
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
      error: payload.error || response.statusText || "Author Companion API request failed."
    };
  }

  async function aceFetchGoogleDocWordCount(message) {
    const documentId = String(message.documentId || "").trim();
    if (!documentId) {
      return {
        ok: false,
        status: 400,
        wordCount: null,
        error: "Google Docs document ID is required."
      };
    }

    if (!aceGoogleOAuthConfigured()) {
      return {
        ok: false,
        status: 400,
        wordCount: null,
        error: "Google OAuth client ID is not configured."
      };
    }

    const snapshot = await aceFetchGoogleDocSnapshot(documentId, Boolean(message.interactive));
    console.info("[ACE] Google Docs word count", {
      documentId,
      revisionId: snapshot.revisionId || "",
      wordCount: snapshot.wordCount
    });

    return {
      ok: true,
      status: snapshot.status,
      method: "google-docs-api",
      revisionId: snapshot.revisionId || "",
      wordCount: snapshot.wordCount
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
        error: "Google Docs document ID and session ID are required."
      };
    }

    if (!aceGoogleOAuthConfigured()) {
      return {
        ok: false,
        status: 400,
        wordCount: null,
        error: "Google OAuth client ID is not configured."
      };
    }

    const snapshot = await aceFetchGoogleDocSnapshot(documentId, Boolean(message.interactive));
    await aceStorageSet({
      [aceSnapshotStorageKey(extensionSessionId)]: {
        documentId,
        revisionId: snapshot.revisionId,
        wordCount: snapshot.wordCount,
        wordCounts: snapshot.wordCounts,
        createdAt: new Date().toISOString()
      }
    });

    console.info("[ACE] Google Docs start snapshot", {
      documentId,
      extensionSessionId,
      revisionId: snapshot.revisionId || "",
      wordCount: snapshot.wordCount
    });

    return {
      ok: true,
      status: snapshot.status,
      method: "google-docs-api",
      revisionId: snapshot.revisionId || "",
      wordCount: snapshot.wordCount
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
        error: "Google Docs document ID and session ID are required."
      };
    }

    if (!aceGoogleOAuthConfigured()) {
      return {
        ok: false,
        status: 400,
        wordCount: null,
        wordsAdded: 0,
        wordsRemoved: 0,
        error: "Google OAuth client ID is not configured."
      };
    }

    const stored = await aceStorageGet(aceSnapshotStorageKey(extensionSessionId));
    const startSnapshot = stored[aceSnapshotStorageKey(extensionSessionId)];
    if (!startSnapshot?.wordCounts || startSnapshot.documentId !== documentId) {
      return {
        ok: false,
        status: 404,
        wordCount: null,
        wordsAdded: 0,
        wordsRemoved: 0,
        error: "The Google Docs start snapshot is missing. Start a new session."
      };
    }

    const endSnapshot = await aceFetchGoogleDocSnapshot(documentId, Boolean(message.interactive));
    const diff = aceCompareWordCounts(startSnapshot.wordCounts, endSnapshot.wordCounts);

    if (message.clearSnapshot) {
      await aceStorageRemove(aceSnapshotStorageKey(extensionSessionId));
    }

    console.info("[ACE] Google Docs word diff", {
      documentId,
      extensionSessionId,
      startRevisionId: startSnapshot.revisionId || "",
      endRevisionId: endSnapshot.revisionId || "",
      wordsAdded: diff.wordsAdded,
      wordsRemoved: diff.wordsRemoved
    });

    return {
      ok: true,
      status: endSnapshot.status,
      method: "google-docs-api",
      revisionId: endSnapshot.revisionId || "",
      startRevisionId: startSnapshot.revisionId || "",
      wordCount: endSnapshot.wordCount,
      startWordCount: startSnapshot.wordCount,
      wordsAdded: diff.wordsAdded,
      wordsRemoved: diff.wordsRemoved,
      netWordsChanged: endSnapshot.wordCount - startSnapshot.wordCount
    };
  }

  async function aceFetchGoogleDocSnapshot(documentId, interactive) {
    const token = await aceGetGoogleAuthToken(interactive);
    const response = await fetch(
      `https://docs.googleapis.com/v1/documents/${encodeURIComponent(documentId)}?includeTabsContent=true`,
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

      const failure = new Error(error);
      failure.status = response.status;
      throw failure;
    }

    const documentPayload = await response.json();
    const wordCounts = aceWordCountsInText(aceExtractGoogleDocText(documentPayload));
    return {
      status: response.status,
      revisionId: documentPayload.revisionId || "",
      wordCounts,
      wordCount: aceTotalWordCounts(wordCounts)
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
            reject(new Error(chrome.runtime.lastError?.message || "Google sign-in was not completed."));
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
    const chunks = [];

    function collectTextRun(paragraphElement) {
      const content = paragraphElement?.textRun?.content;
      if (content) {
        chunks.push(content);
      }
    }

    function collectStructuralElement(element) {
      if (element?.paragraph?.elements) {
        element.paragraph.elements.forEach(collectTextRun);
      }

      if (element?.table?.tableRows) {
        element.table.tableRows.forEach(function (row) {
          (row.tableCells || []).forEach(function (cell) {
            collectStructuralElements(cell.content || []);
          });
        });
      }

      if (element?.tableOfContents?.content) {
        collectStructuralElements(element.tableOfContents.content);
      }
    }

    function collectStructuralElements(elements) {
      (elements || []).forEach(collectStructuralElement);
    }

    function collectDocumentTab(documentTab) {
      if (!documentTab) {
        return;
      }

      collectStructuralElements(documentTab.body?.content || []);
      Object.values(documentTab.headers || {}).forEach(function (header) {
        collectStructuralElements(header.content || []);
      });
      Object.values(documentTab.footers || {}).forEach(function (footer) {
        collectStructuralElements(footer.content || []);
      });
      Object.values(documentTab.footnotes || {}).forEach(function (footnote) {
        collectStructuralElements(footnote.content || []);
      });
    }

    function collectTab(tab) {
      collectDocumentTab(tab?.documentTab);
      (tab?.childTabs || []).forEach(collectTab);
    }

    if (Array.isArray(documentPayload.tabs) && documentPayload.tabs.length) {
      documentPayload.tabs.forEach(collectTab);
    } else {
      collectDocumentTab(documentPayload);
    }

    return chunks.join(" ");
  }

  function aceCountWordsInText(text) {
    return aceTotalWordCounts(aceWordCountsInText(text));
  }

  function aceWordCountsInText(text) {
    const counts = {};
    const matches = String(text || "").match(/[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu) || [];
    matches.forEach(function (word) {
      const normalized = word.toLocaleLowerCase();
      counts[normalized] = (counts[normalized] || 0) + 1;
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
})();
