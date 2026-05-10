(function () {
  "use strict";

  const ACE_API_BASE_URL = "https://davishedrick.pythonanywhere.com";
  const ACE_APP_URL = ACE_API_BASE_URL;
  const ACE_WIDGET_ID = "ace-widget";
  const ACE_TIMER_INTERVAL_MS = 1000;
  const ACE_GOOGLE_DOC_SETTLE_DELAY_MS = 500;
  const ACE_GOOGLE_DOC_POLL_DELAY_MS = 700;
  const ACE_GOOGLE_DOC_POLL_ATTEMPTS = 6;
  const ACE_ACTIVITY_MESSAGE = "ace-writing-activity";
  const ACE_QUOTE_FIND_MESSAGE = "ace-find-quote";
  const ACE_QUOTE_FIND_RESULT_MESSAGE = "ace-find-quote-result";
  const ACE_GOOGLE_DOC_WORD_COUNT_MESSAGE = "ace-google-doc-word-count";
  const ACE_GOOGLE_DOC_START_SNAPSHOT_MESSAGE = "ace-google-doc-start-snapshot";
  const ACE_GOOGLE_DOC_DIFF_MESSAGE = "ace-google-doc-diff";
  const ACE_IS_TOP_FRAME = window.top === window;
  const ACE_ISSUE_TITLE_WORD_LIMIT = 8;

  const ACE_SESSION_STORAGE = {
    widgetPosition: "ace-widget-position"
  };

  const ACE_LOCAL_STORAGE = {
    activeSession: "aceActiveSession",
    pendingSessions: "acePendingSessions",
    documentBindings: "aceDocumentBindings",
    documentBaselines: "aceDocumentBaselines"
  };

  const ACE_DEFAULT_POSITION = "middle-right";
  const ACE_SNAP_MARGIN = 18;
  const ACE_POSITIONS = [
    "top-left",
    "top-right",
    "middle-left",
    "middle-right",
    "bottom-left",
    "bottom-right"
  ];

  const ACE_IGNORED_KEYS = new Set([
    "Shift",
    "Control",
    "Alt",
    "Meta",
    "CapsLock",
    "Tab",
    "Escape",
    "ArrowUp",
    "ArrowDown",
    "ArrowLeft",
    "ArrowRight",
    "Home",
    "End",
    "PageUp",
    "PageDown",
    "Insert",
    "F1",
    "F2",
    "F3",
    "F4",
    "F5",
    "F6",
    "F7",
    "F8",
    "F9",
    "F10",
    "F11",
    "F12"
  ]);

  const ACE_ISSUE_SECTION_NUMBER_WORDS = {
    zero: 0,
    one: 1,
    two: 2,
    three: 3,
    four: 4,
    five: 5,
    six: 6,
    seven: 7,
    eight: 8,
    nine: 9,
    ten: 10,
    eleven: 11,
    twelve: 12,
    thirteen: 13,
    fourteen: 14,
    fifteen: 15,
    sixteen: 16,
    seventeen: 17,
    eighteen: 18,
    nineteen: 19,
    twenty: 20,
    thirty: 30,
    forty: 40,
    fifty: 50,
    sixty: 60,
    seventy: 70,
    eighty: 80,
    ninety: 90
  };
  const ACE_ISSUE_SECTION_REGEX = new RegExp(
    "(chapter|scene|section)\\s+(?:\\d+|(?:zero|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety)(?:[-\\s](?:one|two|three|four|five|six|seven|eight|nine))?)\\b",
    "i"
  );
  const ACE_ISSUE_TYPE_RULES = [
    { type: "Dialogue", keywords: ["dialogue", "conversation", "line"] },
    { type: "Pacing", keywords: ["slow", "fast", "drag", "rush"] },
    { type: "Clarity", keywords: ["confusing", "unclear", "hard to follow"] },
    { type: "Character", keywords: ["motivation", "arc", "character"] },
    { type: "Grammar", keywords: ["grammar", "typo", "spelling"] }
  ];
  const ACE_ISSUE_HIGH_PRIORITY_KEYWORDS = ["major", "critical", "big"];

  let aceState = "idle";
  let aceActiveSession = null;
  let aceCompletedSession = null;
  let aceTimerId = null;
  let aceProjects = [];
  let aceSelectedProject = null;
  let aceProjectPickerMode = "completed";
  let aceSyncStatus = "";
  let aceSyncMessage = "";
  let aceWidget = null;
  let aceWidgetPosition = sessionStorage.getItem(ACE_SESSION_STORAGE.widgetPosition) || ACE_DEFAULT_POSITION;
  let aceDragState = null;
  let aceLastPointerActionAt = 0;
  let acePromptError = "";
  let aceExitHandled = false;
  let aceAutoSyncRestoreId = "";
  let aceCatchUpCandidate = null;
  let aceIssueDraft = null;
  let aceIssueReturnState = "idle";
  let aceCurrentIssues = [];
  let aceIssueStatus = "";

  if (ACE_IS_TOP_FRAME) {
    if (document.getElementById(ACE_WIDGET_ID)) {
      return;
    }

    aceWidget = document.createElement("div");
    aceWidget.id = ACE_WIDGET_ID;
    aceWidget.className = "ace-widget";
    aceWidget.setAttribute("role", "status");
    aceWidget.setAttribute("aria-live", "polite");
    aceWidget.title = "Drag to move";
    document.documentElement.appendChild(aceWidget);
  }

  function aceStorageGet(keys) {
    return new Promise(function (resolve) {
      const runtime = aceChromeRuntime();
      if (!runtime?.storage?.local) {
        resolve({});
        return;
      }

      try {
        runtime.storage.local.get(keys, function (value) {
          resolve(runtime.runtime.lastError ? {} : value);
        });
      } catch (_error) {
        resolve({});
      }
    });
  }

  function aceStorageSet(values) {
    return new Promise(function (resolve) {
      const runtime = aceChromeRuntime();
      if (!runtime?.storage?.local) {
        resolve();
        return;
      }

      try {
        runtime.storage.local.set(values, function () {
          resolve();
        });
      } catch (_error) {
        resolve();
      }
    });
  }

  function aceStorageRemove(keys) {
    return new Promise(function (resolve) {
      const runtime = aceChromeRuntime();
      if (!runtime?.storage?.local) {
        resolve();
        return;
      }

      try {
        runtime.storage.local.remove(keys, function () {
          resolve();
        });
      } catch (_error) {
        resolve();
      }
    });
  }

  function aceChromeRuntime() {
    if (typeof chrome === "undefined") {
      return null;
    }

    return chrome;
  }

  function aceExtensionContextAvailable() {
    const runtime = aceChromeRuntime();
    return Boolean(runtime?.runtime?.id && runtime?.runtime?.sendMessage);
  }

  function aceClosest(target, selector) {
    return target?.closest ? target.closest(selector) : null;
  }

  function aceCapitalize(value) {
    return value.charAt(0).toUpperCase() + value.slice(1);
  }

  function aceExtractDocumentId() {
    const match = window.location.pathname.match(/\/document\/d\/([^/]+)/);
    return match ? decodeURIComponent(match[1]) : "";
  }

  function aceDocumentUrl() {
    return window.location.href.split("#")[0];
  }

  function aceCreateExtensionSessionId(documentId) {
    const random = Math.random().toString(36).slice(2, 10);
    return `ace-${documentId || "doc"}-${Date.now()}-${random}`;
  }

  function aceCreateExtensionIssueId(documentId) {
    const random = Math.random().toString(36).slice(2, 10);
    return `ace-issue-${documentId || "doc"}-${Date.now()}-${random}`;
  }

  function aceNormalizeIssueNoteText(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function aceNormalizeIssueComparisonText(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function aceParseIssueSectionNumber(value) {
    const normalizedValue = aceNormalizeIssueNoteText(value);
    if (/^\d+$/.test(normalizedValue)) {
      return normalizedValue;
    }

    const wordTokens = aceNormalizeIssueComparisonText(normalizedValue)
      .split(" ")
      .filter(Boolean);
    if (!wordTokens.length) {
      return normalizedValue;
    }
    if (!wordTokens.every(function (token) {
      return Object.prototype.hasOwnProperty.call(ACE_ISSUE_SECTION_NUMBER_WORDS, token);
    })) {
      return normalizedValue;
    }
    return String(wordTokens.reduce(function (total, token) {
      return total + ACE_ISSUE_SECTION_NUMBER_WORDS[token];
    }, 0));
  }

  function aceFormatDerivedIssueSection(match) {
    const normalizedMatch = aceNormalizeIssueNoteText(match);
    const sectionParts = normalizedMatch.match(/^(chapter|scene|section)\s+(.+)$/i);
    if (!sectionParts) {
      return normalizedMatch || "Unassigned";
    }
    const sectionType = aceCapitalize(sectionParts[1].toLowerCase());
    return `${sectionType} ${aceParseIssueSectionNumber(sectionParts[2])}`;
  }

  function aceDeriveIssueFieldsFromNote(note) {
    const normalizedNote = aceNormalizeIssueNoteText(note);
    const words = normalizedNote.split(" ").filter(Boolean);
    const matchedSection = String(note || "").match(ACE_ISSUE_SECTION_REGEX);
    const normalizedLower = normalizedNote.toLowerCase();
    const matchedType = ACE_ISSUE_TYPE_RULES.find(function (rule) {
      return rule.keywords.some(function (keyword) {
        return normalizedLower.includes(keyword);
      });
    });
    return {
      title: words.length
        ? words.slice(0, Math.min(ACE_ISSUE_TITLE_WORD_LIMIT, words.length)).join(" ")
        : "Untitled issue",
      sectionLabel: matchedSection?.[0]
        ? aceFormatDerivedIssueSection(matchedSection[0])
        : "Unassigned",
      type: matchedType?.type || "General",
      priority: ACE_ISSUE_HIGH_PRIORITY_KEYWORDS.some(function (keyword) {
        return normalizedLower.includes(keyword);
      }) ? "High" : "Medium"
    };
  }

  function aceElapsedMs() {
    if (!aceActiveSession?.startedAt) {
      return 0;
    }

    return Math.max(0, Date.now() - new Date(aceActiveSession.startedAt).getTime());
  }

  function aceDurationMinutes(milliseconds) {
    return Math.max(1, Math.round(milliseconds / 60000));
  }

  function aceFormatTimer(milliseconds) {
    const totalSeconds = Math.floor(milliseconds / 1000);
    const seconds = totalSeconds % 60;
    const totalMinutes = Math.floor(totalSeconds / 60);
    const minutes = totalMinutes % 60;
    const hours = Math.floor(totalMinutes / 60);

    if (hours > 0) {
      return `${hours}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
    }

    return `${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
  }

  function aceFormatCompletedMinutes(minutes) {
    return `${Math.max(1, Math.round(minutes))} min`;
  }

  function aceFormatWords(count) {
    const words = Math.max(0, Number(count) || 0);
    return `${words} word${words === 1 ? "" : "s"}`;
  }

  function aceOpenApp() {
    window.open(ACE_APP_URL, "_blank", "noopener,noreferrer");
  }

  function aceIsExtensionContextError(error) {
    const message = String(error?.message || error || "").toLowerCase();
    return message.includes("extension context")
      || message.includes("context invalidated")
      || message.includes("message port closed")
      || message.includes("receiving end does not exist");
  }

  async function aceDirectApiFetch(path, options) {
    if (!path.startsWith("/api/")) {
      throw new Error("Invalid Author Companion API path.");
    }

    const response = await fetch(`${ACE_API_BASE_URL}${path}`, {
      method: options?.method || "GET",
      credentials: "include",
      headers: {
        "Content-Type": "application/json",
        ...(options?.headers || {})
      },
      body: options?.body || undefined
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

    if (!response.ok) {
      throw new Error(payload.error || response.statusText || "Author Companion API request failed.");
    }
    return payload;
  }

  async function aceApiFetchViaRuntime(path, options) {
    return new Promise(function (resolve, reject) {
      const runtime = aceChromeRuntime();
      if (!runtime?.runtime?.sendMessage) {
        reject(new Error("Extension context is not available. Refresh the Google Doc."));
        return;
      }

      runtime.runtime.sendMessage(
        {
          aceType: "ace-api-fetch",
          path,
          options: {
            method: options?.method || "GET",
            headers: options?.headers || {},
            body: options?.body || ""
          }
        },
        function (messageResponse) {
          if (runtime.runtime.lastError) {
            reject(new Error(runtime.runtime.lastError.message));
            return;
          }
          resolve(messageResponse || {});
        }
      );
    });
  }

  async function aceApiFetch(path, options) {
    let response = {};
    try {
      response = await aceApiFetchViaRuntime(path, options);
    } catch (error) {
      if (aceIsExtensionContextError(error)) {
        return aceDirectApiFetch(path, options);
      }
      throw error;
    }
    if (!response.ok) {
      throw new Error(response.error || "Author Companion API request failed.");
    }
    return response.payload || {};
  }

  async function aceGetProjects() {
    const payload = await aceApiFetch("/api/projects");
    return Array.isArray(payload.projects) ? payload.projects : [];
  }

  async function aceSaveBinding(documentId, projectId) {
    return aceApiFetch("/api/extension/document-binding", {
      method: "PUT",
      body: JSON.stringify({ documentId, projectId })
    });
  }

  async function aceLocalDocumentBindings() {
    const stored = await aceStorageGet(ACE_LOCAL_STORAGE.documentBindings);
    const bindings = stored[ACE_LOCAL_STORAGE.documentBindings];
    return bindings && typeof bindings === "object" && !Array.isArray(bindings)
      ? bindings
      : {};
  }

  async function aceGetLocalDocumentBinding(documentId) {
    if (!documentId) {
      return null;
    }

    const bindings = await aceLocalDocumentBindings();
    return bindings[documentId] || null;
  }

  async function aceSaveLocalDocumentBinding(documentId, project) {
    if (!documentId || !project?.id) {
      return;
    }

    const bindings = await aceLocalDocumentBindings();
    await aceStorageSet({
      [ACE_LOCAL_STORAGE.documentBindings]: {
        ...bindings,
        [documentId]: {
          projectId: project.id,
          project,
          updatedAt: new Date().toISOString()
        }
      }
    });
  }

  async function aceLocalDocumentBaselines() {
    const stored = await aceStorageGet(ACE_LOCAL_STORAGE.documentBaselines);
    const baselines = stored[ACE_LOCAL_STORAGE.documentBaselines];
    return baselines && typeof baselines === "object" && !Array.isArray(baselines)
      ? baselines
      : {};
  }

  async function aceGetDocumentBaseline(documentId) {
    if (!documentId) {
      return null;
    }

    const baselines = await aceLocalDocumentBaselines();
    return baselines[documentId] || null;
  }

  async function aceSaveDocumentBaseline(session, project) {
    const documentId = String(session?.documentId || "").trim();
    const projectId = String(session?.projectId || project?.id || "").trim();
    const endWordCount = Number(session?.endDocumentWordCount);
    if (!documentId || !projectId || !Number.isFinite(endWordCount)) {
      return;
    }

    const baselines = await aceLocalDocumentBaselines();
    await aceStorageSet({
      [ACE_LOCAL_STORAGE.documentBaselines]: {
        ...baselines,
        [documentId]: {
          documentId,
          projectId,
          project: project || null,
          endDocumentWordCount: Math.max(0, endWordCount),
          syncedAt: new Date().toISOString(),
          sessionId: session.extensionSessionId || ""
        }
      }
    });
  }

  async function acePostSession(session) {
    const payload = aceSessionSyncPayload(session);
    console.info("[ACE] Syncing extension session payload", payload);
    return aceApiFetch("/api/extension/sessions", {
      method: "POST",
      body: JSON.stringify(payload)
    });
  }

  async function acePostIssue(issue) {
    console.info("[ACE] Syncing extension issue payload", issue);
    return aceApiFetch("/api/extension/issues", {
      method: "POST",
      body: JSON.stringify(issue)
    });
  }

  async function aceGetExtensionIssues(documentId) {
    return aceApiFetch(
      `/api/extension/issues?documentId=${encodeURIComponent(documentId)}`
    );
  }

  function aceSessionSyncPayload(session) {
    const sessionType = session?.sessionType === "editing" ? "editing" : "writing";
    const startWordCount = Number(session?.startDocumentWordCount);
    const endWordCount = Number(session?.endDocumentWordCount);
    const startDocumentWordCount = Number.isFinite(startWordCount)
      ? startWordCount
      : null;
    const endDocumentWordCount = Number.isFinite(endWordCount)
      ? endWordCount
      : null;
    const measuredNetWordsChanged = Number.isFinite(startDocumentWordCount) && Number.isFinite(endDocumentWordCount)
      ? endDocumentWordCount - startDocumentWordCount
      : Number(session?.netWordsChanged) || 0;
    const editingBreakdown = aceEditingBreakdownForSync(session, measuredNetWordsChanged);

    return {
      ...session,
      sessionType,
      wordsWritten: sessionType === "writing"
        ? Math.max(0, Number(session?.wordsWritten) || measuredNetWordsChanged)
        : 0,
      wordsAdded: sessionType === "editing" ? editingBreakdown.wordsAdded : 0,
      wordsRemoved: sessionType === "editing" ? editingBreakdown.wordsRemoved : 0,
      wordsEdited: sessionType === "editing" ? editingBreakdown.wordsEdited : 0,
      netWordsChanged: measuredNetWordsChanged,
      startDocumentWordCount,
      endDocumentWordCount,
      wordCountMethod: "google-docs-api",
      measurementPending: Boolean(session?.measurementPending)
    };
  }

  function aceEditingBreakdownForSync(session, netWordsChanged) {
    const explicitWordsAdded = Math.max(0, Number(session?.wordsAdded) || 0);
    const explicitWordsRemoved = Math.max(0, Number(session?.wordsRemoved) || 0);
    if (explicitWordsAdded || explicitWordsRemoved) {
      return {
        wordsAdded: explicitWordsAdded,
        wordsRemoved: explicitWordsRemoved,
        wordsEdited: explicitWordsAdded + explicitWordsRemoved
      };
    }

    const wordsEdited = Math.max(0, Number(session?.wordsEdited) || 0);
    if (!wordsEdited) {
      return {
        wordsAdded: 0,
        wordsRemoved: 0,
        wordsEdited: 0
      };
    }

    const derivedWordsAdded = (wordsEdited + netWordsChanged) / 2;
    const derivedWordsRemoved = (wordsEdited - netWordsChanged) / 2;
    if (
      Number.isFinite(derivedWordsAdded)
      && Number.isFinite(derivedWordsRemoved)
      && derivedWordsAdded >= 0
      && derivedWordsRemoved >= 0
    ) {
      return {
        wordsAdded: Math.round(derivedWordsAdded),
        wordsRemoved: Math.round(derivedWordsRemoved),
        wordsEdited
      };
    }

    return {
      wordsAdded: 0,
      wordsRemoved: 0,
      wordsEdited
    };
  }

  async function aceGoogleDocMessage(aceType, payload) {
    const response = await new Promise(function (resolve) {
      const runtime = aceChromeRuntime();
      if (!runtime?.runtime?.sendMessage) {
        resolve({
          ok: false,
          wordCount: null,
          error: "Extension context is not available. Refresh the Google Doc."
        });
        return;
      }

      try {
        runtime.runtime.sendMessage(
          {
            aceType,
            ...payload
          },
          function (messageResponse) {
            if (runtime.runtime.lastError) {
              resolve({
                ok: false,
                wordCount: null,
                error: runtime.runtime.lastError.message
              });
              return;
            }

            resolve(messageResponse || {
              ok: false,
              wordCount: null,
              error: "Google Docs word count failed."
            });
          }
        );
      } catch (error) {
        resolve({
          ok: false,
          wordCount: null,
          error: error.message || "Google Docs word count failed."
        });
      }
    });

    const wordCount = Number(response.wordCount);
    const wordsAdded = Number(response.wordsAdded);
    const wordsRemoved = Number(response.wordsRemoved);
    const netWordsChanged = Number(response.netWordsChanged);
    return {
      ...response,
      revisionId: response.revisionId || "",
      startRevisionId: response.startRevisionId || "",
      wordCount: Number.isFinite(wordCount) ? Math.max(0, wordCount) : null,
      wordsAdded: Number.isFinite(wordsAdded) ? Math.max(0, wordsAdded) : 0,
      wordsRemoved: Number.isFinite(wordsRemoved) ? Math.max(0, wordsRemoved) : 0,
      netWordsChanged: Number.isFinite(netWordsChanged) ? netWordsChanged : 0
    };
  }

  async function aceGoogleDocStartSnapshot(documentId, extensionSessionId, interactive) {
    if (!documentId) {
      return {
        ok: false,
        wordCount: null,
        error: "Google Docs document ID is missing."
      };
    }

    return aceGoogleDocMessage(ACE_GOOGLE_DOC_START_SNAPSHOT_MESSAGE, {
      documentId,
      extensionSessionId,
      interactive: Boolean(interactive)
    });
  }

  async function aceGoogleDocWordCount(documentId, interactive) {
    if (!documentId) {
      return {
        ok: false,
        wordCount: null,
        error: "Google Docs document ID is missing."
      };
    }

    return aceGoogleDocMessage(ACE_GOOGLE_DOC_WORD_COUNT_MESSAGE, {
      documentId,
      interactive: Boolean(interactive)
    });
  }

  function aceDelay(milliseconds) {
    return new Promise(function (resolve) {
      window.setTimeout(resolve, milliseconds);
    });
  }

  function aceNextFrame() {
    return new Promise(function (resolve) {
      window.requestAnimationFrame(resolve);
    });
  }

  async function aceGoogleDocDiffAfterSave(
    documentId,
    extensionSessionId,
    startWordCount,
    startRevisionId,
    hadDocumentActivity
  ) {
    await aceDelay(ACE_GOOGLE_DOC_SETTLE_DELAY_MS);

    const expectedRevisionChange = Boolean(hadDocumentActivity && startRevisionId);
    let bestResponse = {
      ok: false,
      wordCount: null,
      revisionId: "",
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged: 0,
      error: ""
    };
    let bestActivity = Number.NEGATIVE_INFINITY;
    let bestNetMagnitude = Number.NEGATIVE_INFINITY;
    let bestAttempt = 0;

    for (let attempt = 0; attempt < ACE_GOOGLE_DOC_POLL_ATTEMPTS; attempt += 1) {
      const response = await aceGoogleDocMessage(ACE_GOOGLE_DOC_DIFF_MESSAGE, {
        documentId,
        extensionSessionId,
        interactive: false,
        clearSnapshot: false
      });
      if (!response.ok || !Number.isFinite(response.wordCount)) {
        bestResponse = response;
      } else {
        const delta = Number.isFinite(startWordCount)
          ? response.wordCount - startWordCount
          : 0;
        const wordsAdded = Math.max(0, Number(response.wordsAdded) || 0);
        const wordsRemoved = Math.max(0, Number(response.wordsRemoved) || 0);
        const totalActivity = wordsAdded + wordsRemoved;
        const netMagnitude = Math.abs(delta);
        const revisionChanged = Boolean(
          startRevisionId
          && response.revisionId
          && response.revisionId !== startRevisionId
        );
        const isBetterResponse = !bestResponse.ok
          || totalActivity > bestActivity
          || (totalActivity === bestActivity && netMagnitude > bestNetMagnitude);

        if (isBetterResponse) {
          bestResponse = response;
          bestActivity = totalActivity;
          bestNetMagnitude = netMagnitude;
          bestAttempt = attempt + 1;
        }

        if (totalActivity > 0 || delta !== 0 || revisionChanged || !expectedRevisionChange) {
          console.info("[ACE] Selected Google Docs word diff", {
            attempt: attempt + 1,
            revisionChanged,
            startWordCount,
            endWordCount: response.wordCount,
            wordsAdded,
            wordsRemoved,
            netWordsChanged: delta
          });
          return response;
        }
      }

      if (attempt < ACE_GOOGLE_DOC_POLL_ATTEMPTS - 1) {
        await aceDelay(ACE_GOOGLE_DOC_POLL_DELAY_MS);
      }
    }

    if (expectedRevisionChange && bestResponse.ok) {
      return {
        ...bestResponse,
        ok: false,
        wordCount: null,
        error: "Google Docs has not published the latest revision yet. Retry shortly."
      };
    }

    console.info("[ACE] Selected Google Docs word diff", {
      attempt: bestAttempt,
      startWordCount,
      endWordCount: bestResponse.wordCount,
      wordsAdded: Math.max(0, Number(bestResponse.wordsAdded) || 0),
      wordsRemoved: Math.max(0, Number(bestResponse.wordsRemoved) || 0),
      netWordsChanged: Number.isFinite(bestResponse.netWordsChanged)
        ? bestResponse.netWordsChanged
        : Number.isFinite(startWordCount) && Number.isFinite(bestResponse.wordCount)
          ? bestResponse.wordCount - startWordCount
          : 0
    });
    return bestResponse;
  }

  function aceClearActivityTimers() {
    // Writing activity no longer schedules an automatic start prompt.
  }

  function aceClearTimer() {
    if (aceTimerId) {
      window.clearInterval(aceTimerId);
      aceTimerId = null;
    }
  }

  async function acePersistActiveSession() {
    if (!aceActiveSession) {
      await aceStorageRemove(ACE_LOCAL_STORAGE.activeSession);
      return;
    }

    await aceStorageSet({ [ACE_LOCAL_STORAGE.activeSession]: aceActiveSession });
  }

  async function acePendingSessions() {
    const stored = await aceStorageGet(ACE_LOCAL_STORAGE.pendingSessions);
    return Array.isArray(stored[ACE_LOCAL_STORAGE.pendingSessions])
      ? stored[ACE_LOCAL_STORAGE.pendingSessions]
      : [];
  }

  async function aceStorePendingSession(session) {
    const pending = await acePendingSessions();
    const withoutDuplicate = pending.filter(function (item) {
      return item.extensionSessionId !== session.extensionSessionId;
    });
    await aceStorageSet({
      [ACE_LOCAL_STORAGE.pendingSessions]: [...withoutDuplicate, session]
    });
  }

  async function aceRemovePendingSession(extensionSessionId) {
    const pending = await acePendingSessions();
    await aceStorageSet({
      [ACE_LOCAL_STORAGE.pendingSessions]: pending.filter(function (item) {
        return item.extensionSessionId !== extensionSessionId;
      })
    });
  }

  function aceResetPromptState() {
    acePromptError = "";
  }

  function aceRenderLoading(message, detail) {
    const detailCopy = detail
      ? `<div class="ace-loading-detail">${aceEscapeHtml(detail)}</div>`
      : "";
    aceWidget.className = "ace-widget ace-widget--loading";
    aceWidget.innerHTML = `
      <div class="ace-loading-line">
        <span class="ace-spinner" aria-hidden="true"></span>
        <span>${aceEscapeHtml(message)}</span>
      </div>
      ${detailCopy}
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderIdle() {
    aceWidget.className = "ace-widget ace-widget--idle";
    aceWidget.innerHTML = `
      <button class="ace-icon-button" type="button" data-ace-action="show-controls" aria-label="Open Author Companion session controls">
        <svg class="ace-pencil-icon" viewBox="0 0 24 24" aria-hidden="true">
          <path d="M4 20h4.2L19.4 8.8a2.1 2.1 0 0 0 0-3L18.2 4.6a2.1 2.1 0 0 0-3 0L4 15.8V20z"></path>
          <path d="M13.8 6 18 10.2"></path>
        </svg>
      </button>
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderPrompt() {
    const errorCopy = acePromptError
      ? `<div class="ace-sync-copy ace-sync-copy--pending">${aceEscapeHtml(aceShortDiagnostic(acePromptError))}</div>`
      : "";
    aceWidget.className = "ace-widget ace-widget--prompt";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      <div class="ace-prompt-copy">Start session?</div>
      ${errorCopy}
      <div class="ace-actions">
        <button class="ace-button ace-button--primary" type="button" data-ace-action="confirm-start">Yes</button>
        <button class="ace-button" type="button" data-ace-action="show-issue-form">Add issue</button>
        <button class="ace-button" type="button" data-ace-action="show-issues">Issues</button>
        <button class="ace-button" type="button" data-ace-action="decline-start">No</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderCatchUpPrompt() {
    const words = Math.max(0, Number(aceCatchUpCandidate?.wordsWritten) || 0);
    const statusCopy = aceSyncStatus
      ? `<div class="ace-sync-copy ace-sync-copy--${aceSyncStatus}">${aceEscapeHtml(aceSyncMessage)}</div>`
      : "";
    aceWidget.className = "ace-widget ace-widget--catch-up";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      <div class="ace-prompt-copy">Looks like ${aceFormatWords(words)} were added since your last session. Add them?</div>
      <div class="ace-project-copy">This creates a 1 min catch-up writing session.</div>
      ${statusCopy}
      <div class="ace-actions">
        <button class="ace-button ace-button--primary" type="button" data-ace-action="add-catch-up">Add missed words</button>
        <button class="ace-button" type="button" data-ace-action="skip-catch-up">Skip</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceSessionWordCount(session = aceActiveSession) {
    if (!session) {
      return 0;
    }

    return session.sessionType === "editing"
      ? Math.max(0, Number(session.wordsEdited) || 0)
      : Math.max(0, Number(session.wordsWritten) || 0);
  }

  function aceSessionWordsCopy(session) {
    if (!session) {
      return "";
    }

    if (session.sessionType === "editing") {
      const wordsAdded = Math.max(0, Number(session.wordsAdded) || 0);
      const wordsRemoved = Math.max(0, Number(session.wordsRemoved) || 0);
      return ` · (+${wordsAdded} words - ${wordsRemoved})`;
    }

    return ` · ${aceFormatWords(aceSessionWordCount(session))}`;
  }

  function aceRenderActive() {
    if (!aceActiveSession) {
      aceClearTimer();
      if (aceCompletedSession) {
        aceState = "completed";
        aceRenderCompleted();
      } else {
        aceState = "idle";
        aceRenderIdle();
      }
      return;
    }

    const label = aceCapitalize(aceActiveSession?.sessionType || "writing");
    const nextType = aceActiveSession?.sessionType === "writing" ? "Editing" : "Writing";

    aceWidget.className = "ace-widget ace-widget--active";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      <div class="ace-session-line">
        <span class="ace-session-type">${label}</span>
        <span class="ace-time">${aceFormatTimer(aceElapsedMs())}</span>
      </div>
      <div class="ace-actions">
        <button class="ace-button ace-button--end" type="button" data-ace-action="end">End</button>
        <button class="ace-button" type="button" data-ace-action="show-issue-form">Add issue</button>
        <button class="ace-button" type="button" data-ace-action="show-issues">Issues</button>
        <button class="ace-button ace-button--switch" type="button" data-ace-action="switch">Switch to ${nextType}</button>
        <button class="ace-button ace-button--switch" type="button" data-ace-action="change-project">Change project</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderCompleted() {
    const label = aceCapitalize(aceCompletedSession?.sessionType || "writing");
    const wordsCopy = aceCompletedSession?.measurementPending ? "" : aceSessionWordsCopy(aceCompletedSession);
    const projectCopy = aceSelectedProject
      ? `<div class="ace-project-copy">Project: ${aceEscapeHtml(aceSelectedProject.bookTitle)}</div>`
      : "";
    const methodCopy = aceCompletedSession?.measurementPending
      ? '<div class="ace-project-copy">Waiting on Google Docs word count.</div>'
      : '<div class="ace-project-copy">Words measured from Google Docs.</div>';
    const wordCountError = aceShortDiagnostic(aceCompletedSession?.wordCountError);
    const diagnosticCopy = wordCountError
      ? `<div class="ace-sync-copy ace-sync-copy--pending">Google count unavailable: ${aceEscapeHtml(wordCountError)}</div>`
      : "";
    const statusCopy = aceSyncStatus
      ? `<div class="ace-sync-copy ace-sync-copy--${aceSyncStatus}">${aceEscapeHtml(aceSyncMessage)}</div>`
      : "";
    const contextInvalid = aceSyncMessage.toLowerCase().includes("extension context");
    const retryDisabled = contextInvalid ? "disabled" : "";
    const changeProjectDisabled = contextInvalid ? "disabled" : "";

    aceWidget.className = "ace-widget ace-widget--completed";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      <div class="ace-completed-copy">${label} session tracked: ${aceFormatCompletedMinutes(aceCompletedSession?.durationMinutes || 1)}${wordsCopy}</div>
      ${projectCopy}
      ${methodCopy}
      ${diagnosticCopy}
      ${statusCopy}
      <div class="ace-actions">
        <button class="ace-button" type="button" data-ace-action="open">Open app</button>
        <button class="ace-button" type="button" data-ace-action="show-issue-form">Add issue</button>
        <button class="ace-button" type="button" data-ace-action="show-issues">Issues</button>
        <button class="ace-button" type="button" data-ace-action="retry-sync" ${retryDisabled}>Retry sync</button>
        <button class="ace-button" type="button" data-ace-action="change-project" ${changeProjectDisabled}>Change project</button>
        ${contextInvalid ? '<button class="ace-button ace-button--primary" type="button" data-ace-action="refresh-page">Refresh doc</button>' : ""}
        <button class="ace-button" type="button" data-ace-action="start-new">Start new</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderProjectPicker() {
    const copy = aceProjectPickerMode === "active"
      ? "Change project"
      : aceProjectPickerMode === "catch-up"
        ? "Choose project for missed words"
        : aceProjectPickerMode === "issue"
          ? "Choose project for this issue"
          : "Choose project";
    const rows = aceProjects.length
      ? aceProjects.map(function (project) {
          return `
            <button class="ace-project-option" type="button" data-ace-project-id="${aceEscapeHtml(project.id)}">
              <span>${aceEscapeHtml(project.bookTitle)}</span>
              <small>${aceEscapeHtml(project.manuscriptType || project.status || "Project")}</small>
            </button>
          `;
        }).join("")
      : '<div class="ace-empty">No active projects found. Open the app to create or reopen one.</div>';

    aceWidget.className = "ace-widget ace-widget--picker";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      <div class="ace-prompt-copy">${copy}</div>
      <div class="ace-project-list">${rows}</div>
      <div class="ace-actions">
        <button class="ace-button" type="button" data-ace-action="open">Open app</button>
        <button class="ace-button" type="button" data-ace-action="retry-sync">Retry</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceGetSelectedText() {
    const activeElement = document.activeElement;
    if (
      activeElement
      && typeof activeElement.value === "string"
      && typeof activeElement.selectionStart === "number"
      && typeof activeElement.selectionEnd === "number"
      && activeElement.selectionEnd > activeElement.selectionStart
    ) {
      return activeElement.value.slice(activeElement.selectionStart, activeElement.selectionEnd);
    }

    const selectedText = window.getSelection ? window.getSelection().toString() : "";
    return aceNormalizeIssueNoteText(selectedText).slice(0, 500);
  }

  function aceIssuePreviewCopy(note) {
    const normalizedNote = aceNormalizeIssueNoteText(note);
    if (!normalizedNote) {
      return "";
    }

    const derived = aceDeriveIssueFieldsFromNote(normalizedNote);
    return `Will file under ${derived.sectionLabel} | ${derived.type} | ${derived.priority} priority.`;
  }

  function aceRenderIssueForm() {
    const note = aceIssueDraft?.note || "";
    const snippet = aceIssueDraft?.snippet || "";
    const preview = aceIssuePreviewCopy(note);
    const statusCopy = aceIssueStatus
      ? `<div class="ace-sync-copy ace-sync-copy--pending">${aceEscapeHtml(aceShortDiagnostic(aceIssueStatus))}</div>`
      : "";

    aceWidget.className = "ace-widget ace-widget--issue-form";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      <form class="ace-issue-form" data-ace-issue-form>
        <div class="ace-prompt-copy">Add issue</div>
        ${statusCopy}
        <label class="ace-field">
          <span>Quick note</span>
          <textarea name="note" rows="3" placeholder="chapter 3 slow">${aceEscapeHtml(note)}</textarea>
        </label>
        <label class="ace-field">
          <span>Quote</span>
          <textarea name="snippet" rows="3" placeholder="Selected or pasted text">${aceEscapeHtml(snippet)}</textarea>
        </label>
        <div class="ace-preview" aria-live="polite">${aceEscapeHtml(preview)}</div>
        <div class="ace-actions">
          <button class="ace-button ace-button--primary" type="button" data-ace-action="save-issue">Save issue</button>
          <button class="ace-button" type="button" data-ace-action="cancel-issue">Cancel</button>
        </div>
      </form>
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderIssuesList(message = "") {
    const rows = aceCurrentIssues.length
      ? aceCurrentIssues.map(function (issue) {
          const snippet = aceNormalizeIssueNoteText(issue.snippet || issue.quoteLocator?.quote || "");
          return `
            <div class="ace-issue-row">
              <div class="ace-issue-title">${aceEscapeHtml(issue.title || "Untitled issue")}</div>
              <div class="ace-issue-meta">${aceEscapeHtml([issue.type, issue.priority, issue.sectionLabel].filter(Boolean).join(" | "))}</div>
              ${snippet ? `<blockquote>${aceEscapeHtml(aceShortDiagnostic(snippet))}</blockquote>` : ""}
              <div class="ace-actions">
                <button class="ace-button" type="button" data-ace-action="find-quote" data-ace-issue-id="${aceEscapeHtml(issue.id)}" ${snippet ? "" : "disabled"}>Find quote</button>
                <button class="ace-button" type="button" data-ace-action="copy-quote" data-ace-issue-id="${aceEscapeHtml(issue.id)}" ${snippet ? "" : "disabled"}>Copy quote</button>
              </div>
            </div>
          `;
        }).join("")
      : '<div class="ace-empty">No open issues saved from this doc yet.</div>';
    const messageCopy = message
      ? `<div class="ace-sync-copy ace-sync-copy--pending">${aceEscapeHtml(message)}</div>`
      : "";

    aceWidget.className = "ace-widget ace-widget--issues";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      <div class="ace-prompt-copy">Doc issues</div>
      ${messageCopy}
      <div class="ace-issue-list">${rows}</div>
      <div class="ace-actions">
        <button class="ace-button ace-button--primary" type="button" data-ace-action="show-issue-form">Add issue</button>
        <button class="ace-button" type="button" data-ace-action="open">Open app</button>
        <button class="ace-button" type="button" data-ace-action="cancel-issue">Done</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceEscapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function aceShortDiagnostic(value) {
    const text = String(value || "").replace(/\s+/g, " ").trim();
    return text.length > 120 ? `${text.slice(0, 117)}...` : text;
  }

  function aceCloseButtonHtml() {
    return '<button class="ace-close-button" type="button" data-ace-action="close-popup" aria-label="Close Author Companion controls">&times;</button>';
  }

  function aceIsActiveSessionCurrent(extensionSessionId) {
    return Boolean(extensionSessionId && aceActiveSession?.extensionSessionId === extensionSessionId);
  }

  function aceIsActiveProjectPickerCurrent(extensionSessionId) {
    return aceIsActiveSessionCurrent(extensionSessionId)
      && aceState === "project-picker"
      && aceProjectPickerMode === "active";
  }

  function aceIsCompletedSessionCurrent(extensionSessionId) {
    return Boolean(extensionSessionId && aceCompletedSession?.extensionSessionId === extensionSessionId);
  }

  function aceRunAsync(promise, label) {
    Promise.resolve(promise).catch(function (error) {
      console.warn(`Author Companion: ${label} failed.`, error);
    });
  }

  function aceStartTimer() {
    aceClearTimer();
    if (!aceActiveSession) {
      aceRenderActive();
      return;
    }

    aceRenderActive();
    if (aceActiveSession) {
      aceTimerId = window.setInterval(aceRenderActive, ACE_TIMER_INTERVAL_MS);
    }
  }

  function aceShowStartPrompt() {
    if (aceState !== "idle") {
      return;
    }

    aceState = "prompt";
    aceClearActivityTimers();
    aceRenderPrompt();
  }

  async function aceShowControls() {
    if (aceActiveSession) {
      aceState = "active";
      aceStartTimer();
      return;
    }

    if (aceCompletedSession) {
      aceState = "completed";
      aceRenderCompleted();
      return;
    }

    await aceCheckCatchUpBeforeStartPrompt();
  }

  function aceRememberIssueReturnState() {
    aceIssueReturnState = aceActiveSession
      ? "active"
      : aceCompletedSession
        ? "completed"
        : aceState === "prompt"
          ? "prompt"
          : "idle";
  }

  function aceReturnFromIssue() {
    aceIssueDraft = null;
    aceIssueStatus = "";
    if (aceActiveSession) {
      aceState = "active";
      aceStartTimer();
      return;
    }

    if (aceCompletedSession) {
      aceState = "completed";
      aceRenderCompleted();
      return;
    }

    if (aceIssueReturnState === "prompt") {
      aceState = "prompt";
      aceRenderPrompt();
      return;
    }

    aceState = "idle";
    aceRenderIdle();
  }

  function aceOpenIssueForm() {
    aceRememberIssueReturnState();
    aceClearTimer();
    const documentId = aceExtractDocumentId();
    const selectedText = aceGetSelectedText();
    aceIssueDraft = {
      documentId,
      documentUrl: aceDocumentUrl(),
      extensionIssueId: aceCreateExtensionIssueId(documentId),
      note: "",
      snippet: selectedText
    };
    aceIssueStatus = selectedText ? "Selected text added as the quote." : "";
    aceState = "issue-form";
    aceRenderIssueForm();
  }

  function aceReadIssueFormDraft() {
    const form = aceWidget.querySelector("[data-ace-issue-form]");
    if (!form) {
      return aceIssueDraft;
    }

    const formData = new FormData(form);
    aceIssueDraft = {
      ...(aceIssueDraft || {}),
      documentId: aceIssueDraft?.documentId || aceExtractDocumentId(),
      documentUrl: aceIssueDraft?.documentUrl || aceDocumentUrl(),
      extensionIssueId: aceIssueDraft?.extensionIssueId || aceCreateExtensionIssueId(aceExtractDocumentId()),
      note: String(formData.get("note") || ""),
      snippet: String(formData.get("snippet") || "")
    };
    return aceIssueDraft;
  }

  async function aceShowIssuesList() {
    aceRememberIssueReturnState();
    aceClearTimer();
    aceState = "issues-loading";
    aceRenderLoading("Loading issues...", "Checking Author Companion.");
    await aceNextFrame();

    try {
      const payload = await aceGetExtensionIssues(aceExtractDocumentId());
      aceCurrentIssues = Array.isArray(payload.issues) ? payload.issues : [];
      aceState = "issues";
      aceRenderIssuesList();
    } catch (error) {
      aceCurrentIssues = [];
      aceState = "issues";
      aceRenderIssuesList(`Could not load issues. ${error.message}`);
    }
  }

  async function aceSaveIssue(projectOverride) {
    const draft = aceReadIssueFormDraft();
    if (!draft) {
      return;
    }

    const note = aceNormalizeIssueNoteText(draft.note);
    if (!note) {
      aceIssueStatus = "A short note is required.";
      aceRenderIssueForm();
      return;
    }

    aceState = "issue-saving";
    aceRenderLoading("Saving issue...", "Sending it to the Edit dashboard.");
    await aceNextFrame();

    let project = projectOverride || null;
    let projectId = String(project?.id || "").trim();
    if (!projectId) {
      const localBinding = await aceGetLocalDocumentBinding(draft.documentId);
      project = localBinding?.project || null;
      projectId = String(localBinding?.projectId || project?.id || "").trim();
    }
    if (!projectId && aceActiveSession?.projectId) {
      projectId = aceActiveSession.projectId;
    }
    if (!projectId && aceCompletedSession?.projectId) {
      projectId = aceCompletedSession.projectId;
    }

    if (!projectId) {
      aceProjectPickerMode = "issue";
      try {
        aceProjects = await aceGetProjects();
        aceState = "project-picker";
        aceRenderProjectPicker();
      } catch (error) {
        aceIssueStatus = `Projects unavailable. ${error.message}`;
        aceState = "issue-form";
        aceRenderIssueForm();
      }
      return;
    }

    const snippet = aceNormalizeIssueNoteText(draft.snippet).slice(0, 500);
    const payload = {
      documentId: draft.documentId,
      projectId,
      extensionIssueId: draft.extensionIssueId,
      note,
      snippet,
      documentUrl: draft.documentUrl,
      source: "chrome-extension",
      quoteLocator: {
        strategy: "quote-finder",
        quote: snippet,
        createdAt: new Date().toISOString()
      }
    };

    try {
      const result = await acePostIssue(payload);
      const selectedProject = result.project || project;
      if (selectedProject) {
        await aceSaveLocalDocumentBinding(draft.documentId, selectedProject);
      }
      aceCurrentIssues = [result.issue, ...aceCurrentIssues.filter(function (issue) {
        return issue.id !== result.issue?.id;
      })].filter(Boolean);
      aceIssueDraft = null;
      aceIssueStatus = "";
      aceState = "issues";
      aceRenderIssuesList(result.duplicate ? "Issue was already saved." : "Issue saved.");
    } catch (error) {
      if (error.message.includes("projectId is required")) {
        aceProjectPickerMode = "issue";
        try {
          aceProjects = await aceGetProjects();
          aceState = "project-picker";
          aceRenderProjectPicker();
        } catch (projectError) {
          aceIssueStatus = `Projects unavailable. ${projectError.message}`;
          aceState = "issue-form";
          aceRenderIssueForm();
        }
        return;
      }

      aceIssueStatus = `Issue not saved. ${error.message}`;
      aceState = "issue-form";
      aceRenderIssueForm();
    }
  }

  async function aceChooseProjectForIssue(projectId) {
    const draft = aceIssueDraft;
    const project = aceProjects.find(function (item) {
      return String(item.id) === String(projectId);
    });
    if (!draft || !project) {
      return;
    }

    try {
      const binding = await aceSaveBinding(draft.documentId, project.id);
      const selectedProject = binding.project || project;
      await aceSaveLocalDocumentBinding(draft.documentId, selectedProject);
      if (aceIssueDraft !== draft) {
        return;
      }
      await aceSaveIssue(selectedProject);
    } catch (error) {
      if (aceIssueDraft !== draft) {
        return;
      }
      aceIssueStatus = `Project not saved. ${error.message}`;
      aceState = "issue-form";
      aceRenderIssueForm();
    }
  }

  async function aceCopyText(text) {
    const value = String(text || "");
    if (!value) {
      return false;
    }

    if (navigator.clipboard?.writeText) {
      try {
        await navigator.clipboard.writeText(value);
        return true;
      } catch (_error) {
        // Fall through to the legacy copy path.
      }
    }

    const textarea = document.createElement("textarea");
    textarea.value = value;
    textarea.setAttribute("readonly", "");
    textarea.style.position = "fixed";
    textarea.style.left = "-9999px";
    document.documentElement.appendChild(textarea);
    textarea.select();
    let copied = false;
    try {
      copied = document.execCommand("copy");
    } catch (_error) {
      copied = false;
    }
    textarea.remove();
    return copied;
  }

  function aceFindIssueById(issueId) {
    return aceCurrentIssues.find(function (issue) {
      return String(issue.id) === String(issueId);
    });
  }

  function aceIssueQuote(issue) {
    return aceNormalizeIssueNoteText(issue?.snippet || issue?.quoteLocator?.quote || "");
  }

  function aceQuoteSearchCandidates(quote) {
    const normalizedQuote = aceNormalizeIssueNoteText(quote);
    const candidates = [
      normalizedQuote,
      normalizedQuote.slice(0, 180),
      normalizedQuote.slice(0, 120),
      normalizedQuote.slice(0, 80)
    ];
    const sentence = normalizedQuote.match(/^.{24,}?[.!?](\s|$)/);
    if (sentence?.[0]) {
      candidates.push(sentence[0]);
    }
    const words = normalizedQuote.split(" ").filter(Boolean);
    for (let index = 0; index < words.length; index += 6) {
      candidates.push(words.slice(index, index + 12).join(" "));
    }

    return [...new Set(candidates.map(aceNormalizeIssueNoteText))]
      .filter(function (candidate) {
        return candidate.length >= 24;
      });
  }

  function aceSelectionNode() {
    const selection = window.getSelection ? window.getSelection() : null;
    if (!selection?.rangeCount) {
      return null;
    }

    const node = selection.getRangeAt(0).startContainer;
    return node?.nodeType === Node.ELEMENT_NODE ? node : node?.parentElement;
  }

  function aceSelectionIsInWidget() {
    const node = aceSelectionNode();
    return Boolean(node && aceWidget?.contains(node));
  }

  function aceScrollSelectionIntoView() {
    const node = aceSelectionNode();
    if (!node || aceSelectionIsInWidget()) {
      return false;
    }

    const target = node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
    target?.scrollIntoView?.({ block: "center", inline: "nearest", behavior: "smooth" });
    return true;
  }

  async function aceFindQuoteInCurrentFrame(quote) {
    if (!window.find) {
      return false;
    }

    document.activeElement?.blur?.();
    for (const candidate of aceQuoteSearchCandidates(quote)) {
      try {
        window.getSelection?.().removeAllRanges?.();
        const found = Boolean(window.find(candidate, false, false, true, false, false, false));
        if (found && !aceSelectionIsInWidget()) {
          await aceNextFrame();
          aceScrollSelectionIntoView();
          return true;
        }
      } catch (_error) {
        // Try the next candidate.
      }
    }

    return false;
  }

  function aceFindQuoteInChildFrames(quote) {
    if (!ACE_IS_TOP_FRAME || !window.frames?.length) {
      return Promise.resolve(false);
    }

    const token = `ace-find-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    return new Promise(function (resolve) {
      let settled = false;
      let pending = window.frames.length;

      function finish(found) {
        if (settled) {
          return;
        }

        if (found || pending <= 0) {
          settled = true;
          window.removeEventListener("message", handleResult);
          resolve(Boolean(found));
        }
      }

      function handleResult(event) {
        if (event.data?.aceType !== ACE_QUOTE_FIND_RESULT_MESSAGE || event.data.token !== token) {
          return;
        }

        pending -= 1;
        finish(Boolean(event.data.found));
      }

      window.addEventListener("message", handleResult);
      for (let index = 0; index < window.frames.length; index += 1) {
        try {
          window.frames[index].postMessage({ aceType: ACE_QUOTE_FIND_MESSAGE, token, quote }, "*");
        } catch (_error) {
          pending -= 1;
        }
      }
      window.setTimeout(function () {
        pending = 0;
        finish(false);
      }, 1200);
    });
  }

  async function aceFindQuoteInDocument(quote) {
    const previousDisplay = aceWidget?.style.display || "";
    const previousVisibility = aceWidget?.style.visibility || "";
    if (aceWidget) {
      aceWidget.style.display = "none";
      aceWidget.style.visibility = "hidden";
    }
    await aceNextFrame();

    const found = await aceFindQuoteInCurrentFrame(quote)
      || await aceFindQuoteInChildFrames(quote);

    if (aceWidget) {
      aceWidget.style.display = previousDisplay;
      aceWidget.style.visibility = previousVisibility;
    }
    return found;
  }

  async function aceCopyIssueQuote(issueId) {
    const issue = aceFindIssueById(issueId);
    const quote = aceIssueQuote(issue);
    const copied = await aceCopyText(quote);
    aceRenderIssuesList(copied ? "Quote copied." : "Could not copy the quote.");
  }

  async function aceFindIssueQuote(issueId) {
    const issue = aceFindIssueById(issueId);
    const quote = aceIssueQuote(issue);
    if (!quote) {
      aceRenderIssuesList("No quote saved for that issue.");
      return;
    }

    const found = await aceFindQuoteInDocument(quote);
    if (found) {
      aceMinimizeWidget();
      return;
    }

    aceRenderIssuesList("Could not jump to that quote. Make sure the quote text is still in this Google Doc.");
  }

  async function aceCheckCatchUpBeforeStartPrompt() {
    aceState = "checking-catch-up";
    acePromptError = "";
    aceRenderLoading("Checking progress...", "Looking for missed words.");
    await aceNextFrame();

    const documentId = aceExtractDocumentId();
    const baseline = await aceGetDocumentBaseline(documentId);
    if (!baseline || !Number.isFinite(Number(baseline.endDocumentWordCount))) {
      aceState = "idle";
      aceShowStartPrompt();
      return;
    }

    const currentSnapshot = await aceGoogleDocWordCount(documentId, true);
    if (!currentSnapshot.ok || !Number.isFinite(currentSnapshot.wordCount)) {
      acePromptError = `Could not check for missed words. ${currentSnapshot.error || "You can still start a session."}`;
      aceState = "idle";
      aceShowStartPrompt();
      return;
    }

    const baselineWordCount = Math.max(0, Number(baseline.endDocumentWordCount) || 0);
    const currentWordCount = Math.max(0, Number(currentSnapshot.wordCount) || 0);
    const wordsWritten = currentWordCount - baselineWordCount;
    if (wordsWritten <= 0) {
      aceState = "idle";
      aceShowStartPrompt();
      return;
    }

    aceCatchUpCandidate = {
      documentId,
      documentUrl: aceDocumentUrl(),
      baseline,
      currentSnapshot,
      startDocumentWordCount: baselineWordCount,
      endDocumentWordCount: currentWordCount,
      wordsWritten
    };
    aceSyncStatus = "";
    aceSyncMessage = "";
    aceState = "catch-up";
    aceRenderCatchUpPrompt();
  }

  function aceSkipCatchUp() {
    aceCatchUpCandidate = null;
    aceSyncStatus = "";
    aceSyncMessage = "";
    aceState = "idle";
    aceShowStartPrompt();
  }

  function aceMinimizeWidget() {
    if (aceActiveSession) {
      aceState = "active-minimized";
      aceClearTimer();
      aceRunAsync(acePersistActiveSession(), "persist minimized session");
      aceRenderIdle();
      return;
    }

    if (aceCompletedSession) {
      aceState = "completed-minimized";
      aceRenderIdle();
      return;
    }

    aceState = "idle";
    aceProjectPickerMode = "completed";
    aceSyncStatus = "";
    aceSyncMessage = "";
    aceProjects = [];
    aceCatchUpCandidate = null;
    aceResetPromptState();
    aceRenderIdle();
  }

  async function aceStartSession() {
    if (aceState !== "prompt") {
      return;
    }

    aceState = "starting";
    acePromptError = "";
    aceRenderLoading("Starting...", "Checking Google Docs.");
    await aceNextFrame();

    const documentId = aceExtractDocumentId();
    const now = new Date().toISOString();
    const extensionSessionId = aceCreateExtensionSessionId(documentId);
    const startSnapshot = await aceGoogleDocStartSnapshot(documentId, extensionSessionId, true);
    if (!startSnapshot.ok || !Number.isFinite(startSnapshot.wordCount)) {
      aceState = "prompt";
      acePromptError = `Google Docs word count is required. ${startSnapshot.error || "Check Google OAuth and try again."}`;
      aceRenderPrompt();
      return;
    }

    acePromptError = "";
    aceState = "active";
    aceActiveSession = {
      startedAt: now,
      sessionType: "writing",
      documentId,
      documentUrl: aceDocumentUrl(),
      extensionSessionId,
      wordsWritten: 0,
      wordsEdited: 0,
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged: 0,
      hadDocumentActivity: false,
      startDocumentWordCount: Number.isFinite(startSnapshot.wordCount)
        ? startSnapshot.wordCount
        : null,
      startDocumentRevisionId: startSnapshot.revisionId || "",
      wordCountMethod: "google-docs-api",
      wordCountError: ""
    };
    aceClearActivityTimers();
    await acePersistActiveSession();
    aceStartTimer();
    aceRunAsync(aceResolveBindingForActiveSession(), "resolve active session binding");
  }

  async function aceRestoreSession() {
    const stored = await aceStorageGet(ACE_LOCAL_STORAGE.activeSession);
    const activeSession = stored[ACE_LOCAL_STORAGE.activeSession];

    if (!activeSession?.startedAt || !activeSession?.extensionSessionId) {
      await aceRestorePendingSessionForCurrentDocument();
      return;
    }

    aceState = "active";
    aceActiveSession = activeSession;
    aceStartTimer();
  }

  async function aceRestorePendingSessionForCurrentDocument() {
    const documentId = aceExtractDocumentId();
    const pending = await acePendingSessions();
    const pendingSession = pending.find(function (session) {
      return session.documentId === documentId;
    });

    aceActiveSession = null;
    if (!pendingSession) {
      aceState = "idle";
      aceRenderIdle();
      return;
    }

    aceState = "completed";
    aceCompletedSession = pendingSession;
    aceSyncStatus = "syncing";
    aceSyncMessage = "Checking Google Docs...";
    aceSelectedProject = null;
    aceRenderCompleted();
    aceRunAsync(aceAutoSyncRestoredSession(), "auto-sync restored session");
  }

  async function aceAutoSyncRestoredSession() {
    if (!aceCompletedSession?.extensionSessionId) {
      return;
    }

    if (aceAutoSyncRestoreId === aceCompletedSession.extensionSessionId) {
      return;
    }

    aceAutoSyncRestoreId = aceCompletedSession.extensionSessionId;
    await aceResolveAndSyncCompletedSession(false);
  }

  async function aceSwitchSessionType() {
    if (aceState !== "active" || !aceActiveSession) {
      return;
    }

    aceActiveSession = {
      ...aceActiveSession,
      sessionType: aceActiveSession.sessionType === "writing" ? "editing" : "writing",
      wordsWritten: Math.max(0, Number(aceActiveSession.wordsWritten) || 0),
      wordsEdited: Math.max(0, Number(aceActiveSession.wordsEdited) || 0),
      wordsAdded: Math.max(0, Number(aceActiveSession.wordsAdded) || 0),
      wordsRemoved: Math.max(0, Number(aceActiveSession.wordsRemoved) || 0),
      netWordsChanged: Number(aceActiveSession.netWordsChanged) || 0
    };
    await acePersistActiveSession();
    aceRenderActive();
  }

  function aceDeclineStart() {
    if (aceState !== "prompt") {
      return;
    }

    aceState = "idle";
    aceClearActivityTimers();
    aceResetPromptState();
    aceRenderIdle();
  }

  async function aceEndSession(options = {}) {
    const fromUnload = Boolean(options.fromUnload);
    const activeSession = aceActiveSession;
    if ((aceState !== "active" && aceState !== "active-minimized") || !activeSession) {
      return;
    }

    aceState = "ending";
    aceClearActivityTimers();
    aceClearTimer();

    const endedAt = new Date().toISOString();
    const elapsedMs = aceElapsedMs();
    const documentId = activeSession.documentId || aceExtractDocumentId();
    if (fromUnload) {
      await aceStoreAutoEndedSession(activeSession, endedAt, elapsedMs, documentId);
      return;
    }

    aceRenderLoading("Ending session...", "Measuring words from Google Docs.");
    await aceNextFrame();

    const wordDiff = await aceGoogleDocDiffAfterSave(
      documentId,
      activeSession.extensionSessionId,
      activeSession.startDocumentWordCount,
      activeSession.startDocumentRevisionId,
      activeSession.hadDocumentActivity
    );
    const endDocumentWordCount = Number.isFinite(wordDiff.wordCount)
      ? wordDiff.wordCount
      : null;
    const wordsAdded = Math.max(0, Number(wordDiff.wordsAdded) || 0);
    const wordsRemoved = Math.max(0, Number(wordDiff.wordsRemoved) || 0);
    const netWordsChanged = Number.isFinite(wordDiff.netWordsChanged)
      ? wordDiff.netWordsChanged
      : Number.isFinite(activeSession.startDocumentWordCount) && Number.isFinite(endDocumentWordCount)
        ? endDocumentWordCount - activeSession.startDocumentWordCount
        : 0;
    const measuredWordsWritten = Math.max(0, netWordsChanged);
    const measuredWordsEdited = wordsAdded + wordsRemoved;
    const measurementPending = !wordDiff.ok || !Number.isFinite(endDocumentWordCount);
    const wordCountError = measurementPending
      ? wordDiff.error || "Google Docs word count unavailable."
      : "";
    aceCompletedSession = {
      documentId: activeSession.documentId || aceExtractDocumentId(),
      projectId: activeSession.projectId || "",
      sessionType: activeSession.sessionType || "writing",
      startedAt: activeSession.startedAt,
      endedAt,
      durationMinutes: aceDurationMinutes(elapsedMs),
      source: "chrome-extension",
      documentUrl: activeSession.documentUrl || aceDocumentUrl(),
      notes: "",
      extensionSessionId: activeSession.extensionSessionId,
      wordsWritten: activeSession.sessionType === "writing" && !measurementPending ? measuredWordsWritten : 0,
      wordsEdited: activeSession.sessionType === "editing" && !measurementPending ? measuredWordsEdited : 0,
      wordsAdded: measurementPending ? 0 : wordsAdded,
      wordsRemoved: measurementPending ? 0 : wordsRemoved,
      netWordsChanged: measurementPending ? 0 : netWordsChanged,
      startDocumentWordCount: Number.isFinite(activeSession.startDocumentWordCount)
        ? activeSession.startDocumentWordCount
        : null,
      startDocumentRevisionId: activeSession.startDocumentRevisionId || "",
      endDocumentWordCount: Number.isFinite(endDocumentWordCount)
        ? endDocumentWordCount
        : null,
      wordCountMethod: "google-docs-api",
      wordCountError,
      hadDocumentActivity: Boolean(activeSession.hadDocumentActivity),
      measurementPending
    };

    aceState = "completed";
    aceSyncStatus = "syncing";
    aceSyncMessage = "Syncing...";
    aceSelectedProject = null;
    aceActiveSession = null;
    await acePersistActiveSession();
    aceRenderCompleted();
    if (measurementPending) {
      await aceMarkSessionUnsynced(`Google Docs count unavailable. ${wordCountError}`);
    } else {
      aceRunAsync(aceResolveAndSyncCompletedSession(false), "sync completed session");
    }
  }

  async function aceStoreAutoEndedSession(activeSession, endedAt, elapsedMs, documentId) {
    const session = {
      documentId: activeSession.documentId || documentId,
      projectId: activeSession.projectId || "",
      sessionType: activeSession.sessionType || "writing",
      startedAt: activeSession.startedAt,
      endedAt,
      durationMinutes: aceDurationMinutes(elapsedMs),
      source: "chrome-extension",
      documentUrl: activeSession.documentUrl || aceDocumentUrl(),
      notes: "",
      extensionSessionId: activeSession.extensionSessionId,
      wordsWritten: 0,
      wordsEdited: 0,
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged: 0,
      startDocumentWordCount: Number.isFinite(activeSession.startDocumentWordCount)
        ? activeSession.startDocumentWordCount
        : null,
      startDocumentRevisionId: activeSession.startDocumentRevisionId || "",
      endDocumentWordCount: null,
      wordCountMethod: "google-docs-api",
      wordCountError: "Document closed before Google Docs word count completed.",
      hadDocumentActivity: Boolean(activeSession.hadDocumentActivity),
      measurementPending: true
    };

    aceCompletedSession = session;
    aceActiveSession = null;
    await acePersistActiveSession();
    await aceStorePendingSession(session);
  }

  async function aceStartNew() {
    aceState = "idle";
    aceActiveSession = null;
    aceCompletedSession = null;
    aceSelectedProject = null;
    aceProjectPickerMode = "completed";
    aceSyncStatus = "";
    aceSyncMessage = "";
    aceProjects = [];
    aceClearActivityTimers();
    aceClearTimer();
    aceResetPromptState();
    await acePersistActiveSession();
    aceRenderIdle();
  }

  async function aceResolveBindingForActiveSession() {
    const extensionSessionId = aceActiveSession?.extensionSessionId;
    const documentId = aceActiveSession?.documentId;
    if (!extensionSessionId || !documentId) {
      return;
    }

    try {
      const binding = await aceGetLocalDocumentBinding(documentId);
      if (!aceIsActiveSessionCurrent(extensionSessionId)) {
        return;
      }

      if (binding?.projectId) {
        aceActiveSession = {
          ...aceActiveSession,
          projectId: binding.projectId
        };
        await acePersistActiveSession();
      }
    } catch (_error) {
      // Binding is resolved again when the session ends.
    }
  }

  async function aceResolveAndSyncCompletedSession(forcePicker) {
    const completedSessionId = aceCompletedSession?.extensionSessionId;
    if (!completedSessionId) {
      return;
    }

    if (aceCompletedSession.measurementPending) {
      const measured = await aceMeasureCompletedSession();
      if (!measured) {
        return;
      }
      if (!aceIsCompletedSessionCurrent(completedSessionId)) {
        return;
      }
    }

    aceSyncStatus = "syncing";
    aceSyncMessage = forcePicker ? "Choose the correct project." : "Syncing...";
    aceRenderCompleted();

    try {
      if (!forcePicker) {
        const documentId = aceCompletedSession?.documentId;
        if (!documentId) {
          return;
        }

        const binding = await aceGetLocalDocumentBinding(documentId);
        if (!aceIsCompletedSessionCurrent(completedSessionId)) {
          return;
        }

        if (binding?.projectId) {
          aceCompletedSession = {
            ...aceCompletedSession,
            projectId: binding.projectId
          };
          aceSelectedProject = binding.project || null;
          await aceSyncCompletedSession();
          return;
        }
      }

      aceProjects = await aceGetProjects();
      if (!aceIsCompletedSessionCurrent(completedSessionId)) {
        return;
      }

      aceState = "project-picker";
      aceSyncStatus = "";
      aceSyncMessage = "";
      aceRenderProjectPicker();
    } catch (error) {
      if (!aceIsCompletedSessionCurrent(completedSessionId)) {
        return;
      }

      await aceMarkSessionUnsynced(error.message);
    }
  }

  async function aceChooseProject(projectId) {
    if (aceProjectPickerMode === "active") {
      await aceChooseProjectForActiveSession(projectId);
      return;
    }

    if (aceProjectPickerMode === "catch-up") {
      await aceChooseProjectForCatchUp(projectId);
      return;
    }

    if (aceProjectPickerMode === "issue") {
      await aceChooseProjectForIssue(projectId);
      return;
    }

    const completedSessionId = aceCompletedSession?.extensionSessionId;
    const documentId = aceCompletedSession?.documentId;
    if (!completedSessionId || !documentId) {
      return;
    }

    const project = aceProjects.find(function (item) {
      return String(item.id) === String(projectId);
    });
    if (!project) {
      return;
    }

    aceSelectedProject = project;
    aceCompletedSession = {
      ...aceCompletedSession,
      projectId: project.id
    };
    aceState = "completed";
    aceSyncStatus = "syncing";
    aceSyncMessage = "Saving project...";
    aceRenderCompleted();

    try {
      const binding = await aceSaveBinding(documentId, project.id);
      if (!aceIsCompletedSessionCurrent(completedSessionId)) {
        return;
      }

      aceSelectedProject = binding.project || project;
      await aceSaveLocalDocumentBinding(documentId, aceSelectedProject);
      if (!aceIsCompletedSessionCurrent(completedSessionId)) {
        return;
      }

      await aceSyncCompletedSession();
    } catch (error) {
      if (!aceIsCompletedSessionCurrent(completedSessionId)) {
        return;
      }

      await aceMarkSessionUnsynced(error.message);
    }
  }

  async function aceChooseProjectForActiveSession(projectId) {
    const extensionSessionId = aceActiveSession?.extensionSessionId;
    const documentId = aceActiveSession?.documentId;
    if (!extensionSessionId || !documentId) {
      return;
    }

    const project = aceProjects.find(function (item) {
      return String(item.id) === String(projectId);
    });
    if (!project) {
      return;
    }

    try {
      const binding = await aceSaveBinding(documentId, project.id);
      if (!aceIsActiveSessionCurrent(extensionSessionId)) {
        return;
      }

      aceActiveSession = {
        ...aceActiveSession,
        projectId: project.id
      };
      aceSelectedProject = binding.project || project;
      await aceSaveLocalDocumentBinding(documentId, aceSelectedProject);
      if (!aceIsActiveSessionCurrent(extensionSessionId)) {
        return;
      }

      await acePersistActiveSession();
      if (!aceIsActiveSessionCurrent(extensionSessionId)) {
        return;
      }

      if (aceState === "active-minimized") {
        aceProjectPickerMode = "completed";
        aceClearTimer();
        aceRenderIdle();
        return;
      }

      if (aceState !== "project-picker" && aceState !== "active") {
        return;
      }

      aceState = "active";
      aceProjectPickerMode = "completed";
      aceStartTimer();
    } catch (error) {
      if (!aceIsActiveSessionCurrent(extensionSessionId)) {
        return;
      }

      if (aceState === "active-minimized") {
        aceProjectPickerMode = "completed";
        aceClearTimer();
        aceRenderIdle();
        return;
      }

      if (aceState !== "project-picker" && aceState !== "active") {
        return;
      }

      aceState = "active";
      aceSyncStatus = "pending";
      aceSyncMessage = `Project not changed. ${error.message}`;
      aceProjectPickerMode = "completed";
      aceRenderActive();
    }
  }

  async function aceChooseProjectForCatchUp(projectId) {
    const catchUpCandidate = aceCatchUpCandidate;
    if (!catchUpCandidate) {
      return;
    }

    const project = aceProjects.find(function (item) {
      return String(item.id) === String(projectId);
    });
    if (!project) {
      return;
    }

    try {
      const binding = await aceSaveBinding(catchUpCandidate.documentId, project.id);
      if (aceCatchUpCandidate !== catchUpCandidate) {
        return;
      }

      const selectedProject = binding.project || project;
      await aceSaveLocalDocumentBinding(catchUpCandidate.documentId, selectedProject);
      if (aceCatchUpCandidate !== catchUpCandidate) {
        return;
      }

      await aceSyncCatchUpSession(selectedProject);
    } catch (error) {
      if (aceCatchUpCandidate !== catchUpCandidate) {
        return;
      }

      aceSyncStatus = "pending";
      aceSyncMessage = `Catch-up not synced. ${error.message}`;
      aceState = "catch-up";
      aceRenderCatchUpPrompt();
    }
  }

  async function aceSyncCatchUpSession(projectOverride) {
    const catchUpCandidate = aceCatchUpCandidate;
    if (!catchUpCandidate) {
      return;
    }

    aceState = "catch-up-syncing";
    aceRenderLoading("Adding missed words...", "Saving catch-up session.");
    await aceNextFrame();

    if (aceCatchUpCandidate !== catchUpCandidate) {
      return;
    }

    let project = projectOverride || catchUpCandidate.baseline?.project || null;
    let projectId = String(project?.id || catchUpCandidate.baseline?.projectId || "").trim();
    if (!projectId) {
      const binding = await aceGetLocalDocumentBinding(catchUpCandidate.documentId);
      if (aceCatchUpCandidate !== catchUpCandidate) {
        return;
      }

      project = binding?.project || project;
      projectId = String(binding?.projectId || project?.id || "").trim();
    }

    if (!projectId) {
      aceProjectPickerMode = "catch-up";
      try {
        aceProjects = await aceGetProjects();
        if (aceCatchUpCandidate !== catchUpCandidate) {
          return;
        }

        aceState = "project-picker";
        aceRenderProjectPicker();
      } catch (error) {
        if (aceCatchUpCandidate !== catchUpCandidate) {
          return;
        }

        aceSyncStatus = "pending";
        aceSyncMessage = `Projects unavailable. ${error.message}`;
        aceState = "catch-up";
        aceRenderCatchUpPrompt();
      }
      return;
    }

    const endedAt = new Date().toISOString();
    const startedAt = new Date(Date.now() - 60000).toISOString();
    const wordsWritten = Math.max(0, Number(catchUpCandidate.wordsWritten) || 0);
    const catchUpSession = {
      documentId: catchUpCandidate.documentId,
      projectId,
      sessionType: "writing",
      startedAt,
      endedAt,
      durationMinutes: 1,
      source: "chrome-extension",
      documentUrl: catchUpCandidate.documentUrl || aceDocumentUrl(),
      notes: "Catch-up: words added outside a tracked session.",
      extensionSessionId: aceCreateExtensionSessionId(catchUpCandidate.documentId),
      wordsWritten,
      wordsEdited: 0,
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged: wordsWritten,
      startDocumentWordCount: catchUpCandidate.startDocumentWordCount,
      startDocumentRevisionId: catchUpCandidate.baseline?.revisionId || "",
      endDocumentWordCount: catchUpCandidate.endDocumentWordCount,
      wordCountMethod: "google-docs-api",
      wordCountError: "",
      hadDocumentActivity: true,
      measurementPending: false
    };

    try {
      const result = await acePostSession(catchUpSession);
      if (aceCatchUpCandidate !== catchUpCandidate) {
        return;
      }

      const syncedProject = result.project || project || catchUpCandidate.baseline?.project || null;
      await aceSaveDocumentBaseline(catchUpSession, syncedProject);
      if (aceCatchUpCandidate !== catchUpCandidate) {
        return;
      }

      aceCatchUpCandidate = null;
      aceProjectPickerMode = "completed";
      aceSelectedProject = syncedProject;
      aceSyncStatus = "";
      aceSyncMessage = "";
      aceState = "idle";
      aceShowStartPrompt();
    } catch (error) {
      if (aceCatchUpCandidate !== catchUpCandidate) {
        return;
      }

      aceSyncStatus = "pending";
      aceSyncMessage = `Catch-up not synced. ${error.message}`;
      aceState = "catch-up";
      aceRenderCatchUpPrompt();
    }
  }

  async function aceShowProjectPickerForActiveSession() {
    const extensionSessionId = aceActiveSession?.extensionSessionId;
    if (!extensionSessionId) {
      return;
    }

    aceClearTimer();
    aceProjectPickerMode = "active";
    aceProjects = [];
    aceState = "project-picker";
    aceWidget.className = "ace-widget ace-widget--picker";
    aceWidget.innerHTML = '<div class="ace-prompt-copy">Loading projects...</div>';
    aceApplyWidgetPosition();

    try {
      aceProjects = await aceGetProjects();
      if (!aceIsActiveProjectPickerCurrent(extensionSessionId)) {
        return;
      }

      aceRenderProjectPicker();
    } catch (error) {
      if (!aceIsActiveProjectPickerCurrent(extensionSessionId)) {
        return;
      }

      aceState = "active";
      aceProjectPickerMode = "completed";
      aceStartTimer();
    }
  }

  async function aceSyncCompletedSession() {
    const completedSessionId = aceCompletedSession?.extensionSessionId;
    if (!completedSessionId) {
      return;
    }

    if (!aceCompletedSession?.projectId) {
      await aceResolveAndSyncCompletedSession(true);
      return;
    }

    if (aceCompletedSession.measurementPending) {
      const measured = await aceMeasureCompletedSession();
      if (!measured) {
        return;
      }
      if (!aceIsCompletedSessionCurrent(completedSessionId)) {
        return;
      }
    }

    aceSyncStatus = "syncing";
    aceSyncMessage = "Syncing...";
    aceRenderCompleted();

    try {
      const sessionToSync = aceCompletedSession;
      const result = await acePostSession(sessionToSync);
      if (!aceIsCompletedSessionCurrent(completedSessionId)) {
        return;
      }

      aceSelectedProject = result.project || aceSelectedProject;
      aceSyncStatus = "synced";
      aceSyncMessage = result.duplicate ? "Already synced." : "Synced.";
      await aceSaveDocumentBaseline(sessionToSync, aceSelectedProject);
      if (!aceIsCompletedSessionCurrent(completedSessionId)) {
        return;
      }

      await aceRemovePendingSession(completedSessionId);
    } catch (error) {
      if (!aceIsCompletedSessionCurrent(completedSessionId)) {
        return;
      }

      await aceMarkSessionUnsynced(error.message);
      return;
    }

    if (!aceIsCompletedSessionCurrent(completedSessionId)) {
      return;
    }

    aceRenderCompleted();
  }

  async function aceMeasureCompletedSession() {
    const completedSession = aceCompletedSession;
    const completedSessionId = completedSession?.extensionSessionId;
    if (!completedSession?.measurementPending) {
      return true;
    }

    aceSyncStatus = "syncing";
    aceSyncMessage = "Checking Google Docs...";
    aceRenderCompleted();

    const wordDiff = await aceGoogleDocDiffAfterSave(
      completedSession.documentId,
      completedSession.extensionSessionId,
      completedSession.startDocumentWordCount,
      completedSession.startDocumentRevisionId,
      completedSession.hadDocumentActivity
    );
    if (!aceIsCompletedSessionCurrent(completedSessionId)) {
      return false;
    }

    if (!wordDiff.ok || !Number.isFinite(wordDiff.wordCount)) {
      await aceMarkSessionUnsynced(`Google Docs count unavailable. ${wordDiff.error || "Try again shortly."}`);
      return false;
    }

    const wordsAdded = Math.max(0, Number(wordDiff.wordsAdded) || 0);
    const wordsRemoved = Math.max(0, Number(wordDiff.wordsRemoved) || 0);
    const netWordsChanged = Number.isFinite(wordDiff.netWordsChanged)
      ? wordDiff.netWordsChanged
      : Number(wordDiff.wordCount) - Number(completedSession.startDocumentWordCount);

    aceCompletedSession = {
      ...aceCompletedSession,
      wordsWritten: completedSession.sessionType === "writing" ? Math.max(0, netWordsChanged) : 0,
      wordsEdited: completedSession.sessionType === "editing" ? wordsAdded + wordsRemoved : 0,
      wordsAdded,
      wordsRemoved,
      netWordsChanged,
      endDocumentWordCount: wordDiff.wordCount,
      wordCountMethod: "google-docs-api",
      wordCountError: "",
      measurementPending: false
    };
    return true;
  }

  async function aceMarkSessionUnsynced(message) {
    if (aceCompletedSession) {
      await aceStorePendingSession(aceCompletedSession);
    }
    aceState = "completed";
    aceSyncStatus = "pending";
    aceSyncMessage = `Not synced yet. ${message || "Try again when the app is reachable."}`;
    aceRenderCompleted();
  }

  function aceIsWritingActivity(event) {
    if (event.ctrlKey || event.altKey || event.metaKey) {
      return false;
    }

    if (event.key === "Backspace" || event.key === "Delete" || event.key === "Enter") {
      return true;
    }

    if (ACE_IGNORED_KEYS.has(event.key)) {
      return false;
    }

    if (event.key.length === 1) {
      return true;
    }

    return event.key === "Unidentified" || event.code.startsWith("Key") || event.code.startsWith("Digit");
  }

  function aceRegisterWritingActivity() {
    aceNoteActiveDocumentActivity();

    if (ACE_IS_TOP_FRAME) {
      aceSchedulePrompt();
      return;
    }

    window.top.postMessage({ aceType: ACE_ACTIVITY_MESSAGE }, "*");
  }

  function aceNoteActiveDocumentActivity() {
    if (
      (aceState !== "active" && aceState !== "active-minimized")
      || !aceActiveSession
      || aceActiveSession.hadDocumentActivity
    ) {
      return;
    }

    aceActiveSession.hadDocumentActivity = true;
    aceRunAsync(acePersistActiveSession(), "persist document activity");
  }

  function aceSchedulePrompt() {
    // Sessions are started explicitly from the corner icon.
  }

  function aceHandleKeydown(event) {
    if (aceWidget?.contains(event.target)) {
      return;
    }

    if (!aceIsWritingActivity(event)) {
      return;
    }

    aceRegisterWritingActivity();
  }

  function aceHandleInputLikeActivity(event) {
    if (aceWidget?.contains(event.target)) {
      return;
    }

    if (event.inputType && event.inputType.startsWith("format")) {
      return;
    }

    aceRegisterWritingActivity();
  }

  function aceHandleClipboardActivity() {
    if (aceWidget?.contains(document.activeElement)) {
      return;
    }

    aceRegisterWritingActivity();
  }

  async function aceHandleQuoteFindMessage(message) {
    if (ACE_IS_TOP_FRAME || !message?.token) {
      return;
    }

    const found = await aceFindQuoteInCurrentFrame(message.quote || "");
    window.top.postMessage({
      aceType: ACE_QUOTE_FIND_RESULT_MESSAGE,
      token: message.token,
      found
    }, "*");
  }

  function aceGetAnchorPoint(position) {
    const rect = aceWidget.getBoundingClientRect();
    const width = rect.width || 120;
    const height = rect.height || 44;
    const left = ACE_SNAP_MARGIN;
    const right = Math.max(ACE_SNAP_MARGIN, window.innerWidth - width - ACE_SNAP_MARGIN);
    const top = ACE_SNAP_MARGIN;
    const middle = Math.max(ACE_SNAP_MARGIN, (window.innerHeight - height) / 2);
    const bottom = Math.max(ACE_SNAP_MARGIN, window.innerHeight - height - ACE_SNAP_MARGIN);

    const points = {
      "top-left": { left, top },
      "top-right": { left: right, top },
      "middle-left": { left, top: middle },
      "middle-right": { left: right, top: middle },
      "bottom-left": { left, top: bottom },
      "bottom-right": { left: right, top: bottom }
    };

    return points[position] || points[ACE_DEFAULT_POSITION];
  }

  function aceApplyWidgetPosition() {
    if (!ACE_IS_TOP_FRAME || !aceWidget || aceDragState) {
      return;
    }

    if (!ACE_POSITIONS.includes(aceWidgetPosition)) {
      aceWidgetPosition = ACE_DEFAULT_POSITION;
    }

    const point = aceGetAnchorPoint(aceWidgetPosition);
    aceWidget.style.left = `${point.left}px`;
    aceWidget.style.top = `${point.top}px`;
    aceWidget.style.right = "auto";
    aceWidget.style.bottom = "auto";
  }

  function aceNearestPosition() {
    const rect = aceWidget.getBoundingClientRect();
    const currentCenterX = rect.left + rect.width / 2;
    const currentCenterY = rect.top + rect.height / 2;
    let nearest = aceWidgetPosition;
    let nearestDistance = Number.POSITIVE_INFINITY;

    ACE_POSITIONS.forEach(function (position) {
      const point = aceGetAnchorPoint(position);
      const centerX = point.left + rect.width / 2;
      const centerY = point.top + rect.height / 2;
      const distance = Math.hypot(currentCenterX - centerX, currentCenterY - centerY);

      if (distance < nearestDistance) {
        nearest = position;
        nearestDistance = distance;
      }
    });

    return nearest;
  }

  function aceClamp(value, min, max) {
    return Math.min(Math.max(value, min), max);
  }

  function aceHandlePointerDown(event) {
    if (event.button !== 0) {
      return;
    }

    if (aceClosest(event.target, "input, textarea, select, label")) {
      event.stopPropagation();
      return;
    }

    if (aceHandleWidgetDecision(event, "pointer")) {
      return;
    }

    const rect = aceWidget.getBoundingClientRect();
    aceDragState = {
      offsetX: event.clientX - rect.left,
      offsetY: event.clientY - rect.top
    };
    aceWidget.classList.add("ace-widget--dragging");
    aceWidget.setPointerCapture(event.pointerId);
  }

  function aceHandlePointerMove(event) {
    if (!aceDragState) {
      return;
    }

    const rect = aceWidget.getBoundingClientRect();
    const maxLeft = window.innerWidth - rect.width - ACE_SNAP_MARGIN;
    const maxTop = window.innerHeight - rect.height - ACE_SNAP_MARGIN;
    const nextLeft = aceClamp(event.clientX - aceDragState.offsetX, ACE_SNAP_MARGIN, maxLeft);
    const nextTop = aceClamp(event.clientY - aceDragState.offsetY, ACE_SNAP_MARGIN, maxTop);

    aceWidget.style.left = `${nextLeft}px`;
    aceWidget.style.top = `${nextTop}px`;
    aceWidget.style.right = "auto";
    aceWidget.style.bottom = "auto";
  }

  function aceHandlePointerUp(event) {
    if (!aceDragState) {
      return;
    }

    aceWidget.releasePointerCapture(event.pointerId);
    aceDragState = null;
    aceWidget.classList.remove("ace-widget--dragging");
    aceWidgetPosition = aceNearestPosition();
    sessionStorage.setItem(ACE_SESSION_STORAGE.widgetPosition, aceWidgetPosition);
    aceApplyWidgetPosition();
  }

  function aceHandleWidgetDecision(event, source) {
    const projectButton = aceClosest(event.target, "[data-ace-project-id]");
    if (projectButton && aceWidget.contains(projectButton)) {
      aceConsumeWidgetEvent(event);
      if (source === "click" && Date.now() - aceLastPointerActionAt < 600) {
        return true;
      }

      if (source === "pointer") {
        aceLastPointerActionAt = Date.now();
      }

      aceRunAsync(aceChooseProject(projectButton.getAttribute("data-ace-project-id")), "choose project");
      return true;
    }

    const button = aceClosest(event.target, "[data-ace-action]");
    if (!button || !aceWidget.contains(button)) {
      return false;
    }

    aceConsumeWidgetEvent(event);
    if (button.disabled) {
      return true;
    }

    if (source === "click" && Date.now() - aceLastPointerActionAt < 600) {
      return true;
    }

    if (source === "pointer") {
      aceLastPointerActionAt = Date.now();
    }

    const action = button.getAttribute("data-ace-action");

    if (action === "confirm-start") {
      aceRunAsync(aceStartSession(), "start session");
    } else if (action === "show-controls") {
      aceRunAsync(aceShowControls(), "show controls");
    } else if (action === "show-issue-form") {
      aceOpenIssueForm();
    } else if (action === "save-issue") {
      aceRunAsync(aceSaveIssue(), "save issue");
    } else if (action === "cancel-issue") {
      aceReturnFromIssue();
    } else if (action === "show-issues") {
      aceRunAsync(aceShowIssuesList(), "show issues");
    } else if (action === "copy-quote") {
      aceRunAsync(aceCopyIssueQuote(button.getAttribute("data-ace-issue-id")), "copy quote");
    } else if (action === "find-quote") {
      aceRunAsync(aceFindIssueQuote(button.getAttribute("data-ace-issue-id")), "find quote");
    } else if (action === "close-popup") {
      aceMinimizeWidget();
    } else if (action === "add-catch-up") {
      aceRunAsync(aceSyncCatchUpSession(), "sync catch-up session");
    } else if (action === "skip-catch-up") {
      aceSkipCatchUp();
    } else if (action === "decline-start") {
      aceDeclineStart();
    } else if (action === "end") {
      aceRunAsync(aceEndSession(), "end session");
    } else if (action === "switch") {
      aceRunAsync(aceSwitchSessionType(), "switch session type");
    } else if (action === "open") {
      aceOpenApp();
    } else if (action === "refresh-page") {
      window.location.reload();
    } else if (action === "retry-sync") {
      if (aceProjectPickerMode === "issue" && aceIssueDraft) {
        aceRunAsync(aceSaveIssue(), "retry issue project picker");
      } else if (aceProjectPickerMode === "catch-up" && aceCatchUpCandidate) {
        aceRunAsync(aceSyncCatchUpSession(), "retry catch-up project picker");
      } else if (aceCompletedSession?.projectId) {
        aceRunAsync(aceSyncCompletedSession(), "retry completed session sync");
      } else {
        aceRunAsync(aceResolveAndSyncCompletedSession(false), "resolve completed session sync");
      }
    } else if (action === "change-project") {
      if (aceCompletedSession) {
        aceRunAsync(aceResolveAndSyncCompletedSession(true), "change completed session project");
      } else if (aceActiveSession) {
        aceRunAsync(aceShowProjectPickerForActiveSession(), "show active project picker");
      }
    } else if (action === "start-new") {
      aceRunAsync(aceStartNew(), "start new session flow");
    }

    return true;
  }

  function aceConsumeWidgetEvent(event) {
    event.preventDefault();
    event.stopPropagation();
    if (event.stopImmediatePropagation) {
      event.stopImmediatePropagation();
    }
  }

  function aceHandleWidgetInput(event) {
    const form = aceClosest(event.target, "[data-ace-issue-form]");
    if (!form || !aceWidget.contains(form)) {
      return;
    }

    aceReadIssueFormDraft();
    aceIssueStatus = "";
    const preview = aceWidget.querySelector(".ace-preview");
    if (preview) {
      preview.textContent = aceIssuePreviewCopy(aceIssueDraft?.note || "");
    }
    const status = aceWidget.querySelector(".ace-sync-copy--pending");
    if (status && status.textContent === "Selected text added as the quote.") {
      status.remove();
    }
  }

  document.addEventListener("beforeinput", aceHandleInputLikeActivity, true);
  document.addEventListener("input", aceHandleInputLikeActivity, true);
  document.addEventListener("keydown", aceHandleKeydown, true);
  document.addEventListener("paste", aceHandleClipboardActivity, true);
  document.addEventListener("cut", aceHandleClipboardActivity, true);

  window.addEventListener("message", function (event) {
    if (event.data?.aceType !== ACE_QUOTE_FIND_MESSAGE) {
      return;
    }

    aceRunAsync(aceHandleQuoteFindMessage(event.data), "find quote in document frame");
  });

  if (ACE_IS_TOP_FRAME) {
    window.addEventListener("message", function (event) {
      if (event.source === window || !event.data || event.data.aceType !== ACE_ACTIVITY_MESSAGE) {
        return;
      }

      aceNoteActiveDocumentActivity();
      aceSchedulePrompt();
    });

    window.addEventListener("resize", aceApplyWidgetPosition);
    aceWidget.addEventListener("pointerdown", aceHandlePointerDown);
    aceWidget.addEventListener("pointermove", aceHandlePointerMove);
    aceWidget.addEventListener("pointerup", aceHandlePointerUp);
    aceWidget.addEventListener("pointercancel", aceHandlePointerUp);

    aceWidget.addEventListener("click", function (event) {
      aceHandleWidgetDecision(event, "click");
    }, true);
    aceWidget.addEventListener("input", aceHandleWidgetInput, true);

    window.addEventListener("pagehide", aceHandleDocumentExit);
    window.addEventListener("beforeunload", aceHandleDocumentExit);

    aceRunAsync(aceRestoreSession(), "restore session");
  }

  function aceHandleDocumentExit() {
    if (aceExitHandled || !aceActiveSession) {
      return;
    }

    aceExitHandled = true;
    aceRunAsync(aceEndSession({ fromUnload: true }), "auto-end session on document exit");
  }
})();
