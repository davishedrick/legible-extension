(function () {
  "use strict";

  const ACE_API_BASE_URL = "https://davishedrick.pythonanywhere.com";
  const ACE_APP_URL = ACE_API_BASE_URL;
  const ACE_WIDGET_ID = "ace-widget";
  const ACE_TIMER_INTERVAL_MS = 1000;
  const ACE_GOOGLE_DOC_SETTLE_DELAY_MS = 1500;
  const ACE_GOOGLE_DOC_POLL_DELAY_MS = 1500;
  const ACE_GOOGLE_DOC_POLL_ATTEMPTS = 12;
  const ACE_ACTIVITY_MESSAGE = "ace-writing-activity";
  const ACE_VISIBLE_WORD_COUNT_MESSAGE = "ace-visible-word-count";
  const ACE_VISIBLE_WORD_COUNT_RESULT_MESSAGE = "ace-visible-word-count-result";
  const ACE_GOOGLE_DOC_WORD_COUNT_MESSAGE = "ace-google-doc-word-count";
  const ACE_GOOGLE_DOC_START_SNAPSHOT_MESSAGE = "ace-google-doc-start-snapshot";
  const ACE_GOOGLE_DOC_DIFF_MESSAGE = "ace-google-doc-diff";
  const ACE_WORD_SNAPSHOT_STORAGE_PREFIX = "aceWordSnapshot:";
  const ACE_AUTO_START_UNBOUND_RECHECK_MS = 30000;
  const ACE_WORD_TOKENIZER_VERSION = "google-docs-like-v2";
  const ACE_IS_TOP_FRAME = window.top === window;
  const ACE_ISSUE_TITLE_WORD_LIMIT = 8;

  const ACE_SESSION_STORAGE = {
    widgetPosition: "ace-widget-position"
  };

  const ACE_LOCAL_STORAGE = {
    activeSession: "aceActiveSession",
    pendingSessions: "acePendingSessions",
    documentBindings: "aceDocumentBindings",
    documentBaselines: "aceDocumentBaselines",
    lastSessionType: "aceLastSessionType"
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
  let aceAutoStartInFlight = false;
  let aceAutoStartBindingMisses = {};
  let aceBaselinePrimeInFlight = false;

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

  function aceNormalizeSessionType(value) {
    return value === "editing" ? "editing" : "writing";
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

  function aceSnapshotStorageKey(extensionSessionId) {
    return `${ACE_WORD_SNAPSHOT_STORAGE_PREFIX}${extensionSessionId}`;
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

  function aceFormatNumber(count) {
    return Math.max(0, Number(count) || 0).toLocaleString("en-US");
  }

  function aceFormatSignedNumber(count) {
    const value = Number(count) || 0;
    return `${value > 0 ? "+" : ""}${value}`;
  }

  function aceOpenApp() {
    window.open(ACE_APP_URL, "_blank", "noopener,noreferrer");
  }

  function aceOpenEditDashboard() {
    window.open(`${ACE_APP_URL}/?view=edit`, "_blank", "noopener,noreferrer");
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
      throw new Error("Invalid Scriptor API path.");
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
      throw new Error(payload.error || response.statusText || "Scriptor API request failed.");
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
      throw new Error(response.error || "Scriptor API request failed.");
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

  async function aceGetServerDocumentBinding(documentId) {
    if (!documentId) {
      return null;
    }
    try {
      const payload = await aceApiFetch(`/api/extension/document-binding?documentId=${encodeURIComponent(documentId)}`);
      return payload?.project || null;
    } catch (_error) {
      return null;
    }
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

  async function aceGetBoundProjectForDocument(documentId) {
    if (!documentId) {
      return null;
    }

    const localBinding = await aceGetLocalDocumentBinding(documentId);
    if (localBinding?.projectId) {
      return localBinding;
    }

    const miss = aceAutoStartBindingMisses[documentId];
    if (miss && Date.now() - miss < ACE_AUTO_START_UNBOUND_RECHECK_MS) {
      return null;
    }

    const serverProject = await aceGetServerDocumentBinding(documentId);
    if (serverProject?.id) {
      await aceSaveLocalDocumentBinding(documentId, serverProject);
      delete aceAutoStartBindingMisses[documentId];
      return {
        projectId: serverProject.id,
        project: serverProject
      };
    }

    aceAutoStartBindingMisses[documentId] = Date.now();
    return null;
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
          endDocumentWordCounts: session?.endDocumentWordCounts || session?.endWordCounts || null,
          endDocumentWordTokens: session?.endDocumentWordTokens || session?.endWordTokens || null,
          endDocumentWordCountTokenizerVersion: session?.wordCountTokenizerVersion || session?.endDocumentWordCountTokenizerVersion || "",
          revisionId: session?.endDocumentRevisionId || session?.revisionId || "",
          syncedAt: new Date().toISOString(),
          sessionId: session.extensionSessionId || ""
        }
      }
    });
  }

  async function aceSaveGoogleDocBaseline(documentId, project) {
    const projectId = String(project?.id || "").trim();
    if (!documentId || !projectId) {
      return;
    }

    const snapshot = await aceGoogleDocWordCount(documentId, false);
    const wordCount = Number(snapshot?.wordCount);
    if (!snapshot.ok || !Number.isFinite(wordCount)) {
      return;
    }

    const baselines = await aceLocalDocumentBaselines();
    await aceStorageSet({
      [ACE_LOCAL_STORAGE.documentBaselines]: {
        ...baselines,
        [documentId]: {
          documentId,
          projectId,
          project,
          endDocumentWordCount: Math.max(0, wordCount),
          endDocumentWordCounts: snapshot.wordCounts || null,
          endDocumentWordTokens: snapshot.wordTokens || null,
          endDocumentWordCountTokenizerVersion: snapshot.wordCountTokenizerVersion || "",
          revisionId: snapshot.revisionId || "",
          syncedAt: new Date().toISOString(),
          sessionId: ""
        }
      }
    });
  }

  async function acePrimeBaselineForCurrentDocument() {
    if (!ACE_IS_TOP_FRAME || aceBaselinePrimeInFlight) {
      return;
    }

    const documentId = aceExtractDocumentId();
    if (!documentId) {
      return;
    }

    if (
      aceState !== "idle"
      || aceActiveSession
      || aceCompletedSession
      || aceCatchUpCandidate
    ) {
      return;
    }

    const existingBaseline = await aceGetDocumentBaseline(documentId);
    if (
      existingBaseline?.endDocumentWordTokens
      && existingBaseline.endDocumentWordCountTokenizerVersion === ACE_WORD_TOKENIZER_VERSION
      && Number.isFinite(Number(existingBaseline.endDocumentWordCount))
    ) {
      return;
    }

    aceBaselinePrimeInFlight = true;
    try {
      const binding = await aceGetBoundProjectForDocument(documentId);
      if (!binding?.project) {
        return;
      }

      await aceSaveGoogleDocBaseline(documentId, binding.project);
    } catch (_error) {
      // A missing non-interactive Google token should not interrupt the document.
    } finally {
      aceBaselinePrimeInFlight = false;
    }
  }

  async function acePostSession(session) {
    const payload = aceSessionSyncPayload(session);
    console.info("[ACE] PRE-SYNC", {
      measurementPath: aceMeasurementPathForSession(payload),
      wordsAdded: payload.wordsAdded,
      wordsRemoved: payload.wordsRemoved,
      wordsEdited: payload.wordsEdited,
      netWordsChanged: payload.netWordsChanged,
      payload
    });
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
    const wordChangeBreakdown = aceWordChangeBreakdownForSync(session, measuredNetWordsChanged);

    const payload = {
      ...session,
      sessionType,
      wordsWritten: sessionType === "writing"
        ? Math.max(0, Number(session?.wordsWritten) || measuredNetWordsChanged)
        : 0,
      wordsAdded: wordChangeBreakdown.wordsAdded,
      wordsRemoved: wordChangeBreakdown.wordsRemoved,
      wordsEdited: sessionType === "editing" ? wordChangeBreakdown.wordsEdited : 0,
      netWordsChanged: measuredNetWordsChanged,
      startDocumentWordCount,
      endDocumentWordCount,
      wordCountMethod: session?.wordCountMethod || "google-docs-api",
      measurementPending: Boolean(session?.measurementPending)
    };
    delete payload.endDocumentWordTokens;
    delete payload.endWordTokens;
    delete payload.wordTokens;
    return payload;
  }

  function aceWordChangeBreakdownForSync(session, netWordsChanged) {
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
        wordsAdded: Math.max(0, Number(netWordsChanged) || 0),
        wordsRemoved: Math.max(0, -(Number(netWordsChanged) || 0)),
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

  function aceMeasurementPathForSession(session) {
    if (session?.measurementPending) {
      return "measurement-unavailable";
    }
    if (session?.wordDiffMethod === "google-api-token-map-fallback") {
      return "api-word-map-diff";
    }
    if (session?.wordDiffMethod === "google-api-token-sequence") {
      return "exact-api-sequence-diff";
    }
    if (session?.wordCountMethod === "visible-total-fallback") {
      return "visible-total-fallback";
    }
    if (session?.wordCountMethod === "visible-total-baseline") {
      return "saved-total-baseline";
    }
    return session?.wordCountMethod === "google-docs-api"
      ? "exact-api-sequence-diff"
      : "measurement-unavailable";
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

    const visibleWordCount = await aceVisibleGoogleDocWordCount();
    const responseWordCount = Number(response.wordCount);
    const wordsAdded = Number(response.wordsAdded);
    const wordsRemoved = Number(response.wordsRemoved);
    const netWordsChanged = Number(response.netWordsChanged);
    return {
      ...response,
      revisionId: response.revisionId || "",
      startRevisionId: response.startRevisionId || "",
      wordCount: Number.isFinite(responseWordCount) ? Math.max(0, responseWordCount) : null,
      apiWordCount: Number.isFinite(responseWordCount) ? Math.max(0, responseWordCount) : null,
      visibleWordCount: Number.isFinite(visibleWordCount) ? Math.max(0, visibleWordCount) : null,
      wordCounts: response.wordCounts || null,
      wordTokens: response.wordTokens || null,
      endWordCounts: response.endWordCounts || null,
      endWordTokens: response.endWordTokens || null,
      wordCountTokenizerVersion: response.wordCountTokenizerVersion || "",
      wordDiffMethod: response.wordDiffMethod || "",
      wordsAdded: Number.isFinite(wordsAdded) ? Math.max(0, wordsAdded) : 0,
      wordsRemoved: Number.isFinite(wordsRemoved) ? Math.max(0, wordsRemoved) : 0,
      netWordsChanged: Number.isFinite(netWordsChanged) ? netWordsChanged : 0
    };
  }

  async function aceVisibleGoogleDocWordCount() {
    const localCandidate = aceVisibleGoogleDocWordCountCandidateLocal();
    const frameCandidate = ACE_IS_TOP_FRAME
      ? await aceVisibleGoogleDocWordCountCandidateFromFrames()
      : null;
    const best = aceBestVisibleWordCountCandidate([localCandidate, frameCandidate]);
    return Number.isFinite(best?.count) ? best.count : null;
  }

  function aceVisibleGoogleDocWordCountLocal() {
    const candidates = [];
    const visitedWindows = new Set();

    function collectFromWindow(targetWindow, depth = 0) {
      if (!targetWindow || visitedWindows.has(targetWindow) || depth > 4) {
        return;
      }

      visitedWindows.add(targetWindow);
      try {
        aceVisibleWordCountCandidatesInDocument(targetWindow.document, targetWindow).forEach(function (candidate) {
          candidates.push(candidate);
        });

        for (let index = 0; index < targetWindow.frames.length; index += 1) {
          collectFromWindow(targetWindow.frames[index], depth + 1);
        }
      } catch (_error) {
        // Some Google Docs frames are not readable from this content script.
      }
    }

    collectFromWindow(window.top || window);
    collectFromWindow(window);

    return aceBestVisibleWordCountCandidate(candidates)?.count ?? null;
  }

  function aceVisibleGoogleDocWordCountCandidateFromFrames() {
    if (!ACE_IS_TOP_FRAME || !window.frames?.length) {
      return Promise.resolve(null);
    }

    const token = `ace-word-count-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    return new Promise(function (resolve) {
      const results = [];
      let pending = window.frames.length;
      let settled = false;

      function finish() {
        if (settled) {
          return;
        }
        settled = true;
        window.removeEventListener("message", handleResult);
        const best = aceBestVisibleWordCountCandidate(results);
        resolve(best || null);
      }

      function handleResult(event) {
        if (event.data?.aceType !== ACE_VISIBLE_WORD_COUNT_RESULT_MESSAGE || event.data.token !== token) {
          return;
        }

        const candidate = event.data.wordCountCandidate || null;
        const count = Number(candidate?.count ?? event.data.wordCount);
        if (Number.isFinite(count)) {
          results.push({
            count: Math.max(0, count),
            score: Number(candidate?.score) || 0,
            bottom: Number(candidate?.bottom) || 0,
            left: Number(candidate?.left) || 0
          });
        }
        pending -= 1;
        if (pending <= 0) {
          finish();
        }
      }

      window.addEventListener("message", handleResult);
      for (let index = 0; index < window.frames.length; index += 1) {
        try {
          window.frames[index].postMessage({ aceType: ACE_VISIBLE_WORD_COUNT_MESSAGE, token }, "*");
        } catch (_error) {
          pending -= 1;
        }
      }

      if (pending <= 0) {
        finish();
        return;
      }

      window.setTimeout(finish, 900);
    });
  }

  function aceHandleVisibleWordCountMessage(message, source) {
    if (!message?.token || !source?.postMessage) {
      return;
    }

    const wordCountCandidate = aceVisibleGoogleDocWordCountCandidateLocal();
    source.postMessage({
      aceType: ACE_VISIBLE_WORD_COUNT_RESULT_MESSAGE,
      token: message.token,
      wordCount: Number.isFinite(wordCountCandidate?.count) ? wordCountCandidate.count : null,
      wordCountCandidate
    }, "*");
  }

  function aceVisibleGoogleDocWordCountCandidateLocal() {
    const candidates = [];
    const visitedWindows = new Set();

    function collectFromWindow(targetWindow, depth = 0) {
      if (!targetWindow || visitedWindows.has(targetWindow) || depth > 4) {
        return;
      }

      visitedWindows.add(targetWindow);
      try {
        aceVisibleWordCountCandidatesInDocument(targetWindow.document, targetWindow).forEach(function (candidate) {
          candidates.push(candidate);
        });

        for (let index = 0; index < targetWindow.frames.length; index += 1) {
          collectFromWindow(targetWindow.frames[index], depth + 1);
        }
      } catch (_error) {
        // Some Google Docs frames are not readable from this content script.
      }
    }

    collectFromWindow(window.top || window);
    collectFromWindow(window);

    return aceBestVisibleWordCountCandidate(candidates);
  }

  function aceBestVisibleWordCountCandidate(candidates) {
    const validCandidates = (candidates || [])
      .filter(function (candidate) {
        return candidate && Number.isFinite(candidate.count);
      })
      .map(function (candidate) {
        return {
          ...candidate,
          count: Math.max(0, Number(candidate.count) || 0)
        };
      });
    const candidatesToRank = validCandidates.some(function (candidate) {
      return candidate.count > 0;
    })
      ? validCandidates.filter(function (candidate) {
        return candidate.count > 0;
      })
      : validCandidates;

    return candidatesToRank
      .sort(function (a, b) {
        return (Number(b.score) || 0) - (Number(a.score) || 0)
          || (Number(b.bottom) || 0) - (Number(a.bottom) || 0)
          || (Number(a.left) || 0) - (Number(b.left) || 0);
      })[0] || null;
  }

  function aceVisibleWordCountCandidatesInDocument(targetDocument, targetWindow) {
    if (!targetDocument || !targetWindow) {
      return [];
    }

    return Array.from(targetDocument.querySelectorAll("div, span, button, [aria-label], [role='button']"))
      .map(function (element) {
        if (element.closest?.("#ace-widget")) {
          return null;
        }

        const count = aceVisibleWordCountFromElement(element);
        if (!Number.isFinite(count)) {
          return null;
        }

        const rect = element.getBoundingClientRect();
        const viewportWidth = targetWindow.innerWidth || targetDocument.documentElement?.clientWidth || 0;
        const viewportHeight = targetWindow.innerHeight || targetDocument.documentElement?.clientHeight || 0;
        const isVisible = rect.width > 0
          && rect.height > 0
          && rect.bottom >= 0
          && rect.right >= 0
          && rect.top <= viewportHeight
          && rect.left <= viewportWidth;
        if (!isVisible) {
          return null;
        }

        const isBottomLeftCounter = rect.left <= Math.min(360, viewportWidth * 0.4)
          && rect.top >= viewportHeight * 0.55
          && rect.bottom >= viewportHeight * 0.65;
        if (!isBottomLeftCounter) {
          return null;
        }

        return {
          count,
          bottom: rect.bottom,
          left: rect.left,
          score: (viewportHeight ? rect.bottom / viewportHeight : 0)
            + (viewportWidth ? (1 - rect.left / viewportWidth) : 0)
        };
      })
      .filter(function (candidate) {
        return candidate && Number.isFinite(candidate.count);
      });
  }

  function aceVisibleWordCountFromElement(element) {
    const texts = [
      element.textContent || "",
      element.getAttribute?.("aria-label") || "",
      element.getAttribute?.("title") || ""
    ];

    for (const rawText of texts) {
      const text = aceNormalizeIssueNoteText(rawText);
      if (!text || text.length > 120) {
        continue;
      }

      const match = text.match(/(?:^|[^\p{L}\p{N}])([\d,]+)\s+words?(?:[^\p{L}\p{N}]|$)/iu);
      if (!match) {
        continue;
      }

      const count = Number(match[1].replace(/,/g, ""));
      if (Number.isFinite(count)) {
        return Math.max(0, count);
      }
    }

    return null;
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

  async function aceVisibleStartSnapshot(documentId, reason = "visible-start-fallback") {
    const visibleWordCount = await aceVisibleGoogleDocWordCount();
    if (!Number.isFinite(visibleWordCount)) {
      return null;
    }

    return {
      ok: true,
      status: 200,
      method: "visible-total-baseline",
      revisionId: "",
      wordCount: Math.max(0, visibleWordCount),
      visibleWordCount: Math.max(0, visibleWordCount),
      wordCounts: null,
      wordTokens: null,
      wordCountTokenizerVersion: "",
      wordCountDiagnostic: `W-START-VISIBLE-FALLBACK: started from visible Google Docs count ${Math.max(0, visibleWordCount)} because ${reason}.`
    };
  }

  async function aceStoreVisibleStartSnapshot(documentId, extensionSessionId, visibleSnapshot) {
    await aceStorageSet({
      [aceSnapshotStorageKey(extensionSessionId)]: {
        documentId,
        revisionId: "",
        wordCount: visibleSnapshot.wordCount,
        wordCounts: null,
        wordTokens: null,
        wordCountTokenizerVersion: "",
        createdAt: new Date().toISOString(),
        source: "visible-total-baseline"
      }
    });
  }

  function aceSessionTypeRequiresExactWordDiff(sessionType) {
    return sessionType === "editing";
  }

  function aceSnapshotHasExactWordTokens(snapshot) {
    return Boolean(
      Array.isArray(snapshot?.wordTokens)
      && snapshot.wordCountTokenizerVersion === ACE_WORD_TOKENIZER_VERSION
    );
  }

  function aceExactWordDiffStartError(startSnapshot) {
    return startSnapshot?.error
      || `E-EXACT-START-TOKENS-MISSING: Editing sessions need a current ${ACE_WORD_TOKENIZER_VERSION} Google Docs before snapshot. Reload the doc, wait for the visible word count to settle, then start again.`;
  }

  async function aceStartSnapshotWithVisibleFallback(documentId, extensionSessionId, interactive, context, options = {}) {
    const allowVisibleFallback = options.allowVisibleFallback !== false;
    const startSnapshot = await aceGoogleDocStartSnapshot(documentId, extensionSessionId, interactive);
    if (startSnapshot.ok && Number.isFinite(startSnapshot.wordCount)) {
      const visibleWordCount = Number(startSnapshot.visibleWordCount);
      if (
        Number.isFinite(visibleWordCount)
        && visibleWordCount > 0
        && Math.max(0, visibleWordCount) !== Math.max(0, Number(startSnapshot.wordCount))
      ) {
        if (!allowVisibleFallback) {
          return {
            ...startSnapshot,
            ok: false,
            error: `E-EXACT-START-MISMATCH: Google Docs API start count ${Math.max(0, Number(startSnapshot.wordCount) || 0)} does not match the visible count ${Math.max(0, visibleWordCount)}. Wait for Google Docs to finish saving, then start a new editing session.`
          };
        }
        const visibleSnapshot = await aceVisibleStartSnapshot(
          documentId,
          `${context} Google API start count ${Math.max(0, Number(startSnapshot.wordCount) || 0)} differed from visible count ${Math.max(0, visibleWordCount)}`
        );
        if (visibleSnapshot) {
          await aceStoreVisibleStartSnapshot(documentId, extensionSessionId, visibleSnapshot);
          return visibleSnapshot;
        }
      }
      return startSnapshot;
    }

    if (!allowVisibleFallback) {
      return {
        ...startSnapshot,
        error: aceIsExtensionContextError(startSnapshot.error)
          ? `E-CONTEXT-INVALIDATED: ${startSnapshot.error} Reload this Google Doc after reloading the extension.`
          : `E-EXACT-START-UNAVAILABLE: ${startSnapshot.error || "Google Docs API before snapshot was unavailable."}`
      };
    }

    const visibleSnapshot = await aceVisibleStartSnapshot(
      documentId,
      startSnapshot.error || `${context} Google API start snapshot was unavailable`
    );
    if (visibleSnapshot) {
      await aceStoreVisibleStartSnapshot(documentId, extensionSessionId, visibleSnapshot);
      return visibleSnapshot;
    }

    return {
      ...startSnapshot,
      error: aceIsExtensionContextError(startSnapshot.error)
        ? `E-CONTEXT-INVALIDATED: ${startSnapshot.error} Reload this Google Doc after reloading the extension.`
        : startSnapshot.error
    };
  }

  async function aceSeedGoogleDocStartSnapshotFromBaseline(documentId, extensionSessionId, baseline) {
    const wordCount = aceOptionalWordCount(baseline?.endDocumentWordCount);
    const wordCounts = baseline?.endDocumentWordCounts || baseline?.wordCounts || null;
    const wordTokens = baseline?.endDocumentWordTokens || baseline?.wordTokens || null;
    if (!documentId || !extensionSessionId || !Number.isFinite(wordCount)) {
      return null;
    }
    const baselineTokenizerVersion = baseline?.endDocumentWordCountTokenizerVersion || baseline?.wordCountTokenizerVersion || "";
    const hasWordTokens = Boolean(
      Array.isArray(wordTokens)
      && baselineTokenizerVersion === ACE_WORD_TOKENIZER_VERSION
    );

    await aceStorageSet({
      [aceSnapshotStorageKey(extensionSessionId)]: {
        documentId,
        revisionId: baseline.revisionId || "",
        wordCount,
        wordCounts: hasWordTokens ? wordCounts : null,
        wordTokens: hasWordTokens ? wordTokens : null,
        wordCountTokenizerVersion: hasWordTokens ? ACE_WORD_TOKENIZER_VERSION : "",
        createdAt: new Date().toISOString(),
        source: hasWordTokens ? "local-baseline" : "visible-total-baseline"
      }
    });

    return {
      ok: true,
      status: 200,
      method: hasWordTokens ? "local-baseline" : "visible-total-baseline",
      revisionId: baseline.revisionId || "",
      wordCount,
      wordCounts: hasWordTokens ? wordCounts : null,
      wordTokens: hasWordTokens ? wordTokens : null,
      wordCountTokenizerVersion: hasWordTokens ? ACE_WORD_TOKENIZER_VERSION : "",
      wordCountDiagnostic: hasWordTokens
        ? ""
        : `W-START-SAVED-TOTAL-BASELINE: started from saved total ${wordCount}; exact added/removed split requires a current ${ACE_WORD_TOKENIZER_VERSION} Google API token snapshot.`
    };
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

  function aceGoogleDocDiffDiagnostic({
    code = "D-API-DIFF",
    attempt = 0,
    startWordCount = null,
    apiEndWordCount = null,
    visibleEndWordCount = null,
    wordsAdded = 0,
    wordsRemoved = 0,
    netWordsChanged = 0,
    revisionChanged = false,
    startRevisionId = "",
    endRevisionId = "",
    source = "google-api-word-map"
  } = {}) {
    const visibleCopy = Number.isFinite(Number(visibleEndWordCount))
      ? Number(visibleEndWordCount)
      : "unavailable";
    const revisionCopy = revisionChanged
      ? "changed"
      : startRevisionId && endRevisionId
        ? "unchanged"
        : "unknown";
    const attemptCopy = attempt
      ? `${attempt}/${ACE_GOOGLE_DOC_POLL_ATTEMPTS}`
      : `0/${ACE_GOOGLE_DOC_POLL_ATTEMPTS}`;

    return `${code}: start ${Number.isFinite(Number(startWordCount)) ? Number(startWordCount) : "unknown"}; API end ${Number.isFinite(Number(apiEndWordCount)) ? Number(apiEndWordCount) : "unknown"}; visible end ${visibleCopy}; added +${Math.max(0, Number(wordsAdded) || 0)}; removed -${Math.max(0, Number(wordsRemoved) || 0)}; net ${aceFormatSignedNumber(netWordsChanged)}; revision ${revisionCopy}; attempt ${attemptCopy}; source ${source}.`;
  }

  function aceVisibleWordCountFallbackResponse(response, {
    attempt = 0,
    startWordCount = null,
    startRevisionId = "",
    revisionChanged = false
  } = {}) {
    const visibleWordCount = Number(response?.visibleWordCount);
    const apiWordCount = Number(response?.apiWordCount ?? response?.wordCount);
    const netWordsChanged = Number.isFinite(startWordCount) && Number.isFinite(visibleWordCount)
      ? visibleWordCount - startWordCount
      : 0;
    const wordsAdded = Math.max(0, netWordsChanged);
    const wordsRemoved = Math.max(0, -netWordsChanged);

    return {
      ...response,
      ok: true,
      wordCount: Math.max(0, visibleWordCount),
      wordsAdded,
      wordsRemoved,
      netWordsChanged,
      wordCounts: null,
      endWordCounts: null,
      wordCountDiagnostic: aceGoogleDocDiffDiagnostic({
        code: "W-VISIBLE-FALLBACK",
        attempt,
        startWordCount,
        apiEndWordCount: apiWordCount,
        visibleEndWordCount: visibleWordCount,
        wordsAdded,
        wordsRemoved,
        netWordsChanged,
        revisionChanged,
        startRevisionId,
        endRevisionId: response?.revisionId || "",
        source: "visible-total-fallback"
      })
    };
  }

  function aceVisibleCatchUpFallbackResponse(response, baselineWordCount, attempt = 0) {
    const visibleWordCount = Number(response?.visibleWordCount);
    const apiWordCount = Number(response?.apiWordCount ?? response?.wordCount);
    const netWordsChanged = Number.isFinite(visibleWordCount)
      ? visibleWordCount - baselineWordCount
      : 0;
    const wordsAdded = Math.max(0, netWordsChanged);
    const wordsRemoved = Math.max(0, -netWordsChanged);

    return {
      ...response,
      ok: true,
      wordCount: Math.max(0, visibleWordCount),
      wordsAdded,
      wordsRemoved,
      netWordsChanged,
      wordCounts: null,
      wordCountDiagnostic: aceGoogleDocDiffDiagnostic({
        code: "W-CATCHUP-VISIBLE-FALLBACK",
        attempt,
        startWordCount: baselineWordCount,
        apiEndWordCount: apiWordCount,
        visibleEndWordCount: visibleWordCount,
        wordsAdded,
        wordsRemoved,
        netWordsChanged,
        revisionChanged: false,
        startRevisionId: "",
        endRevisionId: response?.revisionId || "",
        source: "visible-total-fallback"
      })
    };
  }

  async function aceGoogleDocDiffAfterSave(
    documentId,
    extensionSessionId,
    startWordCount,
    startRevisionId,
    hadDocumentActivity,
    options = {}
  ) {
    await aceDelay(ACE_GOOGLE_DOC_SETTLE_DELAY_MS);

    const expectedRevisionChange = Boolean(hadDocumentActivity && startRevisionId);
    const requireExactWordDiff = Boolean(options.requireExactWordDiff);
    const allowVisibleFallback = options.allowVisibleFallback !== false && !requireExactWordDiff;
    let bestResponse = {
      ok: false,
      wordCount: null,
      revisionId: "",
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged: 0,
      wordCountDiagnostic: "",
      error: ""
    };
    let bestActivity = Number.NEGATIVE_INFINITY;
    let bestNetMagnitude = Number.NEGATIVE_INFINITY;
    let bestAttempt = 0;
    let bestVisibleFallbackResponse = null;

    for (let attempt = 0; attempt < ACE_GOOGLE_DOC_POLL_ATTEMPTS; attempt += 1) {
      const response = await aceGoogleDocMessage(ACE_GOOGLE_DOC_DIFF_MESSAGE, {
        documentId,
        extensionSessionId,
        interactive: false,
        clearSnapshot: false
      });
      if (!response.ok || !Number.isFinite(response.wordCount)) {
        const visibleWordCount = Number(response.visibleWordCount);
        if (
          allowVisibleFallback
          && (
          Number.isFinite(visibleWordCount)
          && visibleWordCount > 0
          && Number.isFinite(startWordCount)
          )
        ) {
          bestVisibleFallbackResponse = aceVisibleWordCountFallbackResponse(response, {
            attempt: attempt + 1,
            startWordCount,
            startRevisionId,
            revisionChanged: false
          });
        }
        bestResponse = response;
      } else {
        const visibleWordCount = Number(response.visibleWordCount);
        const apiWordCount = Number(response.apiWordCount ?? response.wordCount);
        const apiMatchesVisible = !Number.isFinite(visibleWordCount)
          || Math.max(0, apiWordCount) === Math.max(0, visibleWordCount);
        const hasExactWordTokens = Array.isArray(response.endWordTokens)
          && response.wordCountTokenizerVersion === ACE_WORD_TOKENIZER_VERSION;
        const suspiciousZeroWordCount = Number(startWordCount) >= 100
          && Number(apiWordCount) === 0
          && !Number.isFinite(response.visibleWordCount);
        if (suspiciousZeroWordCount) {
          bestResponse = {
            ...response,
            ok: false,
            wordCount: null,
            wordsAdded: 0,
            wordsRemoved: 0,
            netWordsChanged: 0,
            wordCountDiagnostic: aceGoogleDocDiffDiagnostic({
              code: "E-API-ZERO",
              attempt: attempt + 1,
              startWordCount,
              apiEndWordCount: apiWordCount,
              visibleEndWordCount: visibleWordCount,
              wordsAdded: 0,
              wordsRemoved: 0,
              netWordsChanged: 0,
              startRevisionId,
              endRevisionId: response.revisionId || "",
              source: "google-api-word-map"
            }),
            error: "E-API-ZERO: Google Docs API returned 0 words for a non-empty document, so I did not save the count. Retry after Google Docs updates its API result."
          };
          continue;
        }

        if (requireExactWordDiff && !hasExactWordTokens) {
          bestResponse = {
            ...response,
            ok: false,
            wordCount: null,
            wordsAdded: 0,
            wordsRemoved: 0,
            netWordsChanged: 0,
            wordCountDiagnostic: aceGoogleDocDiffDiagnostic({
              code: "E-EXACT-END-TOKENS-MISSING",
              attempt: attempt + 1,
              startWordCount,
              apiEndWordCount: apiWordCount,
              visibleEndWordCount: visibleWordCount,
              wordsAdded: 0,
              wordsRemoved: 0,
              netWordsChanged: 0,
              startRevisionId,
              endRevisionId: response.revisionId || "",
              source: "google-api-token-sequence"
            }),
            error: `E-EXACT-END-TOKENS-MISSING: Editing sessions require a current ${ACE_WORD_TOKENIZER_VERSION} Google Docs token snapshot. Reload the doc, wait for the visible word count to settle, then retry sync.`
          };
          continue;
        }

        if (requireExactWordDiff && !apiMatchesVisible) {
          bestResponse = {
            ...response,
            ok: false,
            wordCount: null,
            wordsAdded: 0,
            wordsRemoved: 0,
            netWordsChanged: 0,
            wordCountDiagnostic: aceGoogleDocDiffDiagnostic({
              code: "E-EXACT-END-MISMATCH",
              attempt: attempt + 1,
              startWordCount,
              apiEndWordCount: apiWordCount,
              visibleEndWordCount: visibleWordCount,
              wordsAdded: Math.max(0, Number(response.wordsAdded) || 0),
              wordsRemoved: Math.max(0, Number(response.wordsRemoved) || 0),
              netWordsChanged: Number.isFinite(startWordCount) ? apiWordCount - startWordCount : 0,
              revisionChanged: Boolean(
                startRevisionId
                && response.revisionId
                && response.revisionId !== startRevisionId
              ),
              startRevisionId,
              endRevisionId: response.revisionId || "",
              source: response.wordDiffMethod || "google-api-token-sequence"
            }),
            error: `E-EXACT-END-MISMATCH: Google Docs API end count ${Math.max(0, apiWordCount)} does not match the visible count ${Math.max(0, visibleWordCount)}. I did not save an editing breakdown because the API is still stale. Retry shortly.`
          };
          continue;
        }

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
        const wordCountDiagnostic = aceGoogleDocDiffDiagnostic({
          code: Number.isFinite(visibleWordCount) && Math.max(0, apiWordCount) !== Math.max(0, visibleWordCount)
            ? "D-VISIBLE-MISMATCH"
            : "D-API-DIFF",
          attempt: attempt + 1,
          startWordCount,
          apiEndWordCount: response.wordCount,
          visibleEndWordCount: visibleWordCount,
          wordsAdded,
          wordsRemoved,
          netWordsChanged: delta,
          revisionChanged,
          startRevisionId,
          endRevisionId: response.revisionId || "",
          source: response.wordDiffMethod || (response.endWordTokens ? "google-api-token-sequence" : "api-total-fallback")
        });
        const isBetterResponse = !bestResponse.ok
          || totalActivity > bestActivity
          || (totalActivity === bestActivity && netMagnitude > bestNetMagnitude);

        const measuredResponse = {
          ...response,
          wordsAdded,
          wordsRemoved,
          netWordsChanged: delta,
          wordCountDiagnostic
        };

        if (isBetterResponse) {
          bestResponse = measuredResponse;
          bestActivity = totalActivity;
          bestNetMagnitude = netMagnitude;
          bestAttempt = attempt + 1;
        }

        if (
          apiMatchesVisible
          && (totalActivity > 0 || delta !== 0 || revisionChanged || !expectedRevisionChange)
        ) {
          console.info("[ACE] Selected Google Docs word diff", {
            attempt: attempt + 1,
            revisionChanged,
            startWordCount,
            endWordCount: response.wordCount,
            visibleWordCount,
            wordsAdded,
            wordsRemoved,
            netWordsChanged: delta,
            wordCountDiagnostic
          });
          return measuredResponse;
        }

        if (
          allowVisibleFallback
          && !apiMatchesVisible
          && Number.isFinite(visibleWordCount)
          && visibleWordCount > 0
        ) {
          bestVisibleFallbackResponse = aceVisibleWordCountFallbackResponse(response, {
            attempt: attempt + 1,
            startWordCount,
            startRevisionId,
            revisionChanged
          });
        }
      }

      if (attempt < ACE_GOOGLE_DOC_POLL_ATTEMPTS - 1) {
        await aceDelay(ACE_GOOGLE_DOC_POLL_DELAY_MS);
      }
    }

    if (bestVisibleFallbackResponse) {
      console.info("[ACE] Selected visible word count fallback", {
        attempt: bestVisibleFallbackResponse.wordCountDiagnostic || "",
        startWordCount,
        apiEndWordCount: bestVisibleFallbackResponse.apiWordCount,
        visibleEndWordCount: bestVisibleFallbackResponse.visibleWordCount,
        wordsAdded: bestVisibleFallbackResponse.wordsAdded,
        wordsRemoved: bestVisibleFallbackResponse.wordsRemoved,
        netWordsChanged: bestVisibleFallbackResponse.netWordsChanged,
        wordCountDiagnostic: bestVisibleFallbackResponse.wordCountDiagnostic
      });
      return bestVisibleFallbackResponse;
    }

    const bestRevisionChanged = Boolean(
      startRevisionId
      && bestResponse.revisionId
      && bestResponse.revisionId !== startRevisionId
    );
    const bestNetWordsChanged = Number.isFinite(bestResponse.netWordsChanged)
      ? bestResponse.netWordsChanged
      : Number.isFinite(startWordCount) && Number.isFinite(bestResponse.wordCount)
        ? bestResponse.wordCount - startWordCount
        : 0;
    const bestTotalActivity = Math.max(0, Number(bestResponse.wordsAdded) || 0)
      + Math.max(0, Number(bestResponse.wordsRemoved) || 0);

    if (
      requireExactWordDiff
      && hadDocumentActivity
      && bestResponse.ok
      && bestTotalActivity === 0
      && bestNetWordsChanged === 0
    ) {
      return {
        ...bestResponse,
        ok: false,
        wordCount: null,
        wordCountDiagnostic: bestResponse.wordCountDiagnostic || aceGoogleDocDiffDiagnostic({
          code: "E-EXACT-NO-TOKEN-CHANGE",
          attempt: bestAttempt,
          startWordCount,
          apiEndWordCount: bestResponse.wordCount,
          visibleEndWordCount: bestResponse.visibleWordCount,
          wordsAdded: 0,
          wordsRemoved: 0,
          netWordsChanged: 0,
          revisionChanged: bestRevisionChanged,
          startRevisionId,
          endRevisionId: bestResponse.revisionId || "",
          source: bestResponse.wordDiffMethod || "google-api-token-sequence"
        }),
        error: "E-EXACT-NO-TOKEN-CHANGE: The editing session had document activity, but the Google Docs before/after token sequence is unchanged. I did not save +0/-0; retry sync after Google Docs finishes saving, or send this diagnostic if the visible count has settled."
      };
    }

    if (
      expectedRevisionChange
      && bestResponse.ok
      && bestTotalActivity === 0
      && bestNetWordsChanged === 0
      && !bestRevisionChanged
    ) {
      const fallbackNetWordsChanged = Number.isFinite(bestResponse.netWordsChanged)
        ? bestResponse.netWordsChanged
        : Number.isFinite(startWordCount) && Number.isFinite(bestResponse.wordCount)
          ? bestResponse.wordCount - startWordCount
          : 0;
      return {
        ...bestResponse,
        ok: false,
        wordCount: null,
        wordCountDiagnostic: bestResponse.wordCountDiagnostic || aceGoogleDocDiffDiagnostic({
          code: "E-API-NO-CHANGE",
          attempt: bestAttempt,
          startWordCount,
          apiEndWordCount: bestResponse.wordCount,
          visibleEndWordCount: bestResponse.visibleWordCount,
          wordsAdded: Math.max(0, Number(bestResponse.wordsAdded) || 0),
          wordsRemoved: Math.max(0, Number(bestResponse.wordsRemoved) || 0),
          netWordsChanged: fallbackNetWordsChanged,
          revisionChanged: false,
          startRevisionId,
          endRevisionId: bestResponse.revisionId || "",
          source: bestResponse.wordDiffMethod || "google-api-token-sequence"
        }),
        error: "E-API-NO-CHANGE: Google Docs API still matches the before snapshot after polling. Retry sync, or run a deletion-only/addition-only test and send the diagnostic line."
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
          : 0,
      wordCountDiagnostic: bestResponse.wordCountDiagnostic || ""
    });
    return bestResponse;
  }

  function aceCompareWordCounts(startWordCounts, endWordCounts, netWordsChanged) {
    if (!startWordCounts || !endWordCounts) {
      return {
        wordsAdded: Math.max(0, Number(netWordsChanged) || 0),
        wordsRemoved: Math.max(0, -(Number(netWordsChanged) || 0))
      };
    }

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

  function aceWordCountsSignature(wordCounts) {
    if (!wordCounts || typeof wordCounts !== "object") {
      return "";
    }

    return Object.keys(wordCounts)
      .sort()
      .map(function (key) {
        return `${key}:${Math.max(0, Number(wordCounts[key]) || 0)}`;
      })
      .join("|");
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

  function aceWordCountsFromTokens(tokens) {
    const counts = {};
    (tokens || []).forEach(function (token) {
      counts[token] = (counts[token] || 0) + 1;
    });
    return counts;
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

  async function aceGoogleDocWordCountAfterSettle(documentId, baseline) {
    await aceDelay(ACE_GOOGLE_DOC_SETTLE_DELAY_MS);

    const baselineWordCount = Math.max(0, Number(baseline?.endDocumentWordCount) || 0);
    const baselineTokenizerVersion = baseline?.endDocumentWordCountTokenizerVersion || baseline?.wordCountTokenizerVersion || "";
    const baselineWordCounts = baselineTokenizerVersion === ACE_WORD_TOKENIZER_VERSION
      ? baseline?.endDocumentWordCounts || baseline?.wordCounts || null
      : null;
    const baselineWordTokens = baselineTokenizerVersion === ACE_WORD_TOKENIZER_VERSION
      ? baseline?.endDocumentWordTokens || baseline?.wordTokens || null
      : null;
    const allowVisibleFallback = true;
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
    let previousSignature = "";
    let stableSnapshots = 0;
    let bestVisibleFallbackResponse = null;

    for (let attempt = 0; attempt < ACE_GOOGLE_DOC_POLL_ATTEMPTS; attempt += 1) {
      const response = await aceGoogleDocWordCount(documentId, attempt === 0);
      const wordCount = Number(response.wordCount);
      if (!response.ok || !Number.isFinite(wordCount)) {
        bestResponse = response;
      } else {
        const visibleWordCount = Number(response.visibleWordCount);
        const apiMatchesVisible = !Number.isFinite(visibleWordCount)
          || Math.max(0, wordCount) === Math.max(0, visibleWordCount);
        const delta = wordCount - baselineWordCount;
        const diff = Array.isArray(baselineWordTokens) && Array.isArray(response.wordTokens)
          ? aceCompareWordTokens(baselineWordTokens, response.wordTokens)
          : aceCompareWordCounts(baselineWordCounts, response.wordCounts, delta);
        const enrichedResponse = {
          ...response,
          wordsAdded: diff.wordsAdded,
          wordsRemoved: diff.wordsRemoved,
          netWordsChanged: delta,
          wordCountDiagnostic: aceGoogleDocDiffDiagnostic({
            code: apiMatchesVisible ? "D-CATCHUP-API" : "D-CATCHUP-VISIBLE-MISMATCH",
            attempt: attempt + 1,
            startWordCount: baselineWordCount,
            apiEndWordCount: wordCount,
            visibleEndWordCount: visibleWordCount,
            wordsAdded: diff.wordsAdded,
            wordsRemoved: diff.wordsRemoved,
            netWordsChanged: delta,
            revisionChanged: false,
            startRevisionId: "",
            endRevisionId: response.revisionId || "",
            source: diff.method || (response.wordCounts ? "google-api-word-map" : "api-total-fallback")
          })
        };
        const totalActivity = diff.wordsAdded + diff.wordsRemoved;
        const netMagnitude = Math.abs(delta);
        const isBetterResponse = !bestResponse.ok
          || totalActivity > bestActivity
          || (totalActivity === bestActivity && netMagnitude > bestNetMagnitude);

        if (isBetterResponse) {
          bestResponse = enrichedResponse;
          bestActivity = totalActivity;
          bestNetMagnitude = netMagnitude;
        }

        const signature = [
          response.revisionId || "",
          wordCount,
          aceWordCountsSignature(response.wordCounts)
        ].join("::");
        stableSnapshots = signature && signature === previousSignature
          ? stableSnapshots + 1
          : 0;
        previousSignature = signature;

        if (
          allowVisibleFallback
          && !apiMatchesVisible
          && Number.isFinite(visibleWordCount)
          && visibleWordCount > 0
        ) {
          bestVisibleFallbackResponse = aceVisibleCatchUpFallbackResponse(
            response,
            baselineWordCount,
            attempt + 1
          );
        }

        if (apiMatchesVisible && stableSnapshots >= 1) {
          console.info("[ACE] Selected stable Google Docs catch-up word count", {
            attempt: attempt + 1,
            baselineWordCount,
            currentWordCount: bestResponse.wordCount,
            visibleWordCount,
            wordsAdded: bestResponse.wordsAdded,
            wordsRemoved: bestResponse.wordsRemoved,
            netWordsChanged: bestResponse.netWordsChanged,
            wordCountDiagnostic: bestResponse.wordCountDiagnostic || ""
          });
          return bestResponse;
        }
      }

      if (attempt < ACE_GOOGLE_DOC_POLL_ATTEMPTS - 1) {
        await aceDelay(ACE_GOOGLE_DOC_POLL_DELAY_MS);
      }
    }

    if (bestVisibleFallbackResponse) {
      console.info("[ACE] Selected visible catch-up word count fallback", {
        baselineWordCount,
        currentWordCount: bestVisibleFallbackResponse.wordCount,
        apiWordCount: bestVisibleFallbackResponse.apiWordCount,
        visibleWordCount: bestVisibleFallbackResponse.visibleWordCount,
        wordsAdded: bestVisibleFallbackResponse.wordsAdded,
        wordsRemoved: bestVisibleFallbackResponse.wordsRemoved,
        netWordsChanged: bestVisibleFallbackResponse.netWordsChanged,
        wordCountDiagnostic: bestVisibleFallbackResponse.wordCountDiagnostic
      });
      return bestVisibleFallbackResponse;
    }

    console.info("[ACE] Selected Google Docs catch-up word count", {
      baselineWordCount,
      currentWordCount: bestResponse.wordCount,
      visibleWordCount: bestResponse.visibleWordCount,
      wordsAdded: Math.max(0, Number(bestResponse.wordsAdded) || 0),
      wordsRemoved: Math.max(0, Number(bestResponse.wordsRemoved) || 0),
      netWordsChanged: Number.isFinite(Number(bestResponse.netWordsChanged))
        ? Number(bestResponse.netWordsChanged)
        : 0,
      wordCountDiagnostic: bestResponse.wordCountDiagnostic || ""
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

  async function aceLastSessionType() {
    const stored = await aceStorageGet(ACE_LOCAL_STORAGE.lastSessionType);
    return aceNormalizeSessionType(stored[ACE_LOCAL_STORAGE.lastSessionType]);
  }

  async function aceRememberLastSessionType(sessionType) {
    await aceStorageSet({
      [ACE_LOCAL_STORAGE.lastSessionType]: aceNormalizeSessionType(sessionType)
    });
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
      ${acePanelHeaderHtml("Scriptor", "Working")}
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
      <button class="ace-icon-button" type="button" data-ace-action="show-controls" aria-label="Open Scriptor session controls">
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
      ${acePanelHeaderHtml("Scriptor", "Google Docs")}
      <div class="ace-prompt-copy">Start a Scriptor session for this doc?</div>
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
    const activity = aceCatchUpActivity(aceCatchUpCandidate);
    const startWordCount = Math.max(0, Number(aceCatchUpCandidate?.startDocumentWordCount) || 0);
    const endWordCount = aceOptionalWordCount(aceCatchUpCandidate?.endDocumentWordCount);
    const needsWordCountConfirmation = Boolean(aceCatchUpCandidate?.needsWordCountConfirmation)
      || !Number.isFinite(endWordCount);
    const promptCopy = needsWordCountConfirmation
      ? "Google Docs returned 0 words, so I do not trust the automatic count. Enter the current word count shown in Google Docs to log catch-up."
      : `I think this doc is at <span data-ace-current-word-count-preview>${aceFormatNumber(endWordCount)}</span> words now. Log catch-up?`;
    const changeCopy = needsWordCountConfirmation
      ? "Detected change: <span data-ace-catch-up-preview>enter a count to calculate changes</span>."
      : `Detected change: <span data-ace-catch-up-preview>${activity.activityCopy}</span>.`;
    const sessionTypeCopy = needsWordCountConfirmation
      ? "catch-up"
      : `catch-up ${activity.isEditingCatchUp ? "editing" : "writing"}`;
    const statusCopy = aceSyncStatus
      ? `<div class="ace-sync-copy ace-sync-copy--${aceSyncStatus}">${aceEscapeHtml(aceSyncMessage)}</div>`
      : "";
    const diagnostic = aceShortDiagnostic(aceCatchUpCandidate?.currentSnapshot?.wordCountDiagnostic, 240);
    const diagnosticCopy = diagnostic
      ? `<div class="ace-project-copy">Diagnostic: ${aceEscapeHtml(diagnostic)}</div>`
      : "";
    aceWidget.className = "ace-widget ace-widget--catch-up";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      ${acePanelHeaderHtml("Catch-up", "Scriptor")}
      <div class="ace-prompt-copy">${promptCopy}</div>
      <div class="ace-project-copy">Scriptor starts from ${aceFormatNumber(startWordCount)} words. ${changeCopy}</div>
      <label class="ace-field ace-field--compact">
        <span>Current word count</span>
        <input type="text" inputmode="numeric" autocomplete="off" data-ace-catch-up-word-count value="${Number.isFinite(endWordCount) ? aceEscapeHtml(aceFormatNumber(endWordCount)) : ""}">
      </label>
      <div class="ace-project-copy">Change the count if Google Docs shows something different. This creates a 1 min ${sessionTypeCopy} session.</div>
      ${diagnosticCopy}
      ${statusCopy}
      <div class="ace-actions">
        <button class="ace-button ace-button--primary" type="button" data-ace-action="add-catch-up">Log catch-up</button>
        <button class="ace-button" type="button" data-ace-action="skip-catch-up">Skip</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceCatchUpActivity(candidate) {
    const wordsAdded = Math.max(0, Number(candidate?.wordsAdded) || 0);
    const wordsRemoved = Math.max(0, Number(candidate?.wordsRemoved) || 0);
    const totalWords = wordsAdded + wordsRemoved;
    const isEditingCatchUp = candidate?.sessionType === "editing";
    const activityCopy = isEditingCatchUp
      ? wordsAdded > 0 && wordsRemoved > 0
        ? `${aceFormatWords(wordsAdded)} added and ${aceFormatWords(wordsRemoved)} removed`
        : `${aceFormatWords(wordsRemoved)} removed`
      : `${aceFormatWords(wordsAdded)} added`;

    return {
      wordsAdded,
      wordsRemoved,
      totalWords,
      isEditingCatchUp,
      activityCopy,
      verb: totalWords === 1 ? "was" : "were"
    };
  }

  function aceParseWordCountInput(value) {
    const normalized = String(value || "").replace(/,/g, "").trim();
    if (!/^\d+$/.test(normalized)) {
      return null;
    }

    const wordCount = Number(normalized);
    return Number.isFinite(wordCount) ? Math.max(0, Math.round(wordCount)) : null;
  }

  function aceOptionalWordCount(value) {
    if (value === null || value === undefined || value === "") {
      return null;
    }

    const wordCount = Number(value);
    return Number.isFinite(wordCount) ? Math.max(0, Math.round(wordCount)) : null;
  }

  function aceReadCatchUpEndWordCount() {
    const input = aceWidget.querySelector("[data-ace-catch-up-word-count]");
    if (!input) {
      return aceOptionalWordCount(aceCatchUpCandidate?.endDocumentWordCount);
    }

    return aceParseWordCountInput(input.value);
  }

  function aceRecalculateCatchUpCandidate(candidate, endWordCount) {
    const startWordCount = Math.max(0, Number(candidate?.startDocumentWordCount) || 0);
    const currentWordCount = Math.max(0, Number(endWordCount) || 0);
    const netWordsChanged = currentWordCount - startWordCount;
    const wordsAdded = Math.max(0, netWordsChanged);
    const wordsRemoved = Math.max(0, -netWordsChanged);
    const sessionType = wordsRemoved > 0 ? "editing" : "writing";

    return {
      ...candidate,
      currentSnapshot: {
        ...(candidate?.currentSnapshot || {}),
        wordCount: currentWordCount,
        wordsAdded,
        wordsRemoved,
        netWordsChanged
      },
      endDocumentWordCount: currentWordCount,
      endDocumentWordCounts: null,
      wordsWritten: wordsRemoved > 0 ? 0 : wordsAdded,
      wordsAdded,
      wordsRemoved,
      wordsEdited: wordsAdded + wordsRemoved,
      netWordsChanged,
      sessionType,
      needsWordCountConfirmation: false
    };
  }

  function aceUnsafeCatchUpMessage(candidate) {
    const startWordCount = Math.max(0, Number(candidate?.startDocumentWordCount) || 0);
    const endWordCount = Math.max(0, Number(candidate?.endDocumentWordCount) || 0);
    if (startWordCount >= 100 && endWordCount === 0) {
      return "That would set the manuscript to 0 words, so I stopped it. Enter the Google Docs word count shown in the lower-left corner.";
    }

    return "";
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

    const wordsAdded = Math.max(0, Number(session.wordsAdded) || 0);
    const wordsRemoved = Math.max(0, Number(session.wordsRemoved) || 0);
    const hasBreakdown = wordsRemoved > 0 || wordsAdded > aceSessionWordCount(session);
    return hasBreakdown
      ? ` · ${aceFormatWords(aceSessionWordCount(session))} (+${wordsAdded} words - ${wordsRemoved})`
      : ` · ${aceFormatWords(aceSessionWordCount(session))}`;
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
      ${acePanelHeaderHtml("Session", label)}
      <div class="ace-session-metric">
        <span class="ace-session-type">${label}</span>
        <strong class="ace-time">${aceFormatTimer(aceElapsedMs())}</strong>
        <small>Tracking from Google Docs</small>
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
    const wordCountDiagnostic = aceShortDiagnostic(aceCompletedSession?.wordCountDiagnostic, 240);
    const wordCountDiagnosticCopy = wordCountDiagnostic
      ? `<div class="ace-project-copy">Diagnostic: ${aceEscapeHtml(wordCountDiagnostic)}</div>`
      : "";
    const statusCopy = aceSyncStatus
      ? `<div class="ace-sync-copy ace-sync-copy--${aceSyncStatus}">${aceEscapeHtml(aceSyncMessage)}</div>`
      : "";
    const contextInvalid = aceSyncMessage.toLowerCase().includes("extension context");
    const retryDisabled = contextInvalid ? "disabled" : "";
    const changeProjectDisabled = contextInvalid ? "disabled" : "";
    console.info("[ACE] UI RENDER", {
      measurementPath: aceMeasurementPathForSession(aceCompletedSession),
      wordsAdded: aceCompletedSession?.wordsAdded,
      wordsRemoved: aceCompletedSession?.wordsRemoved,
      wordsEdited: aceCompletedSession?.wordsEdited,
      netWordsChanged: aceCompletedSession?.netWordsChanged,
      display: wordsCopy
    });

    aceWidget.className = "ace-widget ace-widget--completed";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      ${acePanelHeaderHtml("Session saved", label)}
      <div class="ace-completed-copy">${label} session tracked: ${aceFormatCompletedMinutes(aceCompletedSession?.durationMinutes || 1)}${wordsCopy}</div>
      ${projectCopy}
      ${methodCopy}
      ${wordCountDiagnosticCopy}
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
      ${acePanelHeaderHtml(copy, "Projects")}
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
        ${acePanelHeaderHtml("Add issue", "Edit dashboard")}
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
              <button class="ace-issue-summary" type="button" data-ace-action="show-issue-detail" data-ace-issue-id="${aceEscapeHtml(issue.id)}">
                <span class="ace-issue-title">${aceEscapeHtml(issue.title || "Untitled issue")}</span>
                <span class="ace-issue-meta">${aceEscapeHtml(aceIssueMetaLine(issue))}</span>
                ${snippet ? `<blockquote>${aceEscapeHtml(aceShortDiagnostic(snippet))}</blockquote>` : ""}
              </button>
              <div class="ace-actions ace-issue-row-actions">
                ${snippet ? `<button class="ace-button" type="button" data-ace-action="copy-quote" data-ace-issue-id="${aceEscapeHtml(issue.id)}">Copy quote</button>` : ""}
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
      <div class="ace-issue-header">
        ${acePanelHeaderHtml("Doc issues", "Edit dashboard")}
        <button class="ace-link-button" type="button" data-ace-action="open-edit-dashboard">Open in app</button>
      </div>
      ${messageCopy}
      <div class="ace-issue-list">${rows}</div>
      <div class="ace-actions ace-issue-footer-actions">
        <button class="ace-button ace-button--primary" type="button" data-ace-action="show-issue-form">Add issue</button>
        <button class="ace-button" type="button" data-ace-action="cancel-issue">Done</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderIssueDetail(issueId, message = "") {
    const issue = aceFindIssueById(issueId);
    if (!issue) {
      aceRenderIssuesList("Issue not found.");
      return;
    }

    const snippet = aceIssueQuote(issue);
    const note = aceNormalizeIssueNoteText(issue.notes || issue.note || "");
    const detailRows = aceIssueDetailRows(issue)
      .map(function (row) {
        return `
          <div class="ace-detail-row">
            <span>${aceEscapeHtml(row.label)}</span>
            <strong>${aceEscapeHtml(row.value)}</strong>
          </div>
        `;
      }).join("");
    const messageCopy = message
      ? `<div class="ace-sync-copy ace-sync-copy--pending">${aceEscapeHtml(message)}</div>`
      : "";

    aceWidget.className = "ace-widget ace-widget--issue-detail";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      ${acePanelHeaderHtml("Issue details", "Edit dashboard")}
      ${messageCopy}
      <div class="ace-issue-detail">
        <div class="ace-detail-heading">
          <div class="ace-issue-title">${aceEscapeHtml(issue.title || "Untitled issue")}</div>
          <div class="ace-issue-meta">${aceEscapeHtml(aceIssueMetaLine(issue))}</div>
        </div>
        ${detailRows ? `<div class="ace-detail-grid">${detailRows}</div>` : ""}
        ${snippet ? `
          <section class="ace-detail-section">
            <div class="ace-detail-label">Quote</div>
            <blockquote>${aceEscapeHtml(snippet)}</blockquote>
            <div class="ace-actions">
              <button class="ace-button" type="button" data-ace-action="copy-quote-detail" data-ace-issue-id="${aceEscapeHtml(issue.id)}">Copy quote</button>
            </div>
          </section>
        ` : ""}
        ${note ? `
          <section class="ace-detail-section">
            <div class="ace-detail-label">Note</div>
            <p>${aceEscapeHtml(note)}</p>
          </section>
        ` : ""}
      </div>
      <div class="ace-actions ace-issue-footer-actions">
        <button class="ace-button ace-button--primary" type="button" data-ace-action="back-to-issues">Back</button>
        <button class="ace-button" type="button" data-ace-action="cancel-issue">Done</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceIssueMetaLine(issue) {
    return [issue.type, issue.priority, issue.sectionLabel].filter(Boolean).join(" | ");
  }

  function aceIssueDetailRows(issue) {
    return [
      { label: "Priority", value: issue.priority },
      { label: "Status", value: issue.status },
      { label: "Section", value: issue.sectionLabel },
      { label: "Location", value: issue.textLocation }
    ].filter(function (row) {
      return aceNormalizeIssueNoteText(row.value);
    });
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
    const maxLength = Number(arguments[1]) || 120;
    const text = String(value || "").replace(/\s+/g, " ").trim();
    return text.length > maxLength ? `${text.slice(0, maxLength - 3)}...` : text;
  }

  function aceCloseButtonHtml() {
    return '<button class="ace-close-button" type="button" data-ace-action="close-popup" aria-label="Close Scriptor controls">&times;</button>';
  }

  function acePanelHeaderHtml(title, meta) {
    return `
      <div class="ace-panel-head">
        <span class="ace-brand-mark" aria-hidden="true"></span>
        <div>
          <strong>${aceEscapeHtml(title)}</strong>
          ${meta ? `<span>${aceEscapeHtml(meta)}</span>` : ""}
        </div>
      </div>
    `;
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
      console.warn(`Scriptor: ${label} failed.`, error);
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
    aceRenderLoading("Loading issues...", "Checking Scriptor.");
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

  async function aceCopyIssueQuote(issueId, returnToDetail = false) {
    const issue = aceFindIssueById(issueId);
    const quote = aceIssueQuote(issue);
    const copied = await aceCopyText(quote);
    const message = copied ? "Quote copied." : "Could not copy the quote.";
    if (returnToDetail) {
      aceRenderIssueDetail(issueId, message);
      return;
    }
    aceRenderIssuesList(message);
  }

  async function aceBuildCatchUpCandidate(documentId) {
    const baseline = await aceGetDocumentBaseline(documentId);
    if (!baseline || !Number.isFinite(Number(baseline.endDocumentWordCount))) {
      return { candidate: null, error: "" };
    }

    const baselineWordCount = Math.max(0, Number(baseline.endDocumentWordCount) || 0);
    const boundProject = await aceGetServerDocumentBinding(documentId);
    if (boundProject?.id) {
      await aceSaveLocalDocumentBinding(documentId, boundProject);
    }
    const currentSnapshot = await aceGoogleDocWordCountAfterSettle(documentId, baseline);
    if (!currentSnapshot.ok || !Number.isFinite(currentSnapshot.wordCount)) {
      return {
        candidate: null,
        error: `Could not check for missed words. ${currentSnapshot.error || "You can still start a session."}`
      };
    }

    const currentWordCount = Math.max(0, Number(currentSnapshot.wordCount) || 0);
    const suspiciousZeroWordCount = baselineWordCount >= 100
      && currentWordCount === 0
      && !Number.isFinite(currentSnapshot.visibleWordCount);
    const appWordCount = Number(boundProject?.currentWordCount);
    const shouldUseAppWordCount = Number.isFinite(appWordCount) && Math.max(0, appWordCount) !== currentWordCount;
    const catchUpStartWordCount = shouldUseAppWordCount ? Math.max(0, appWordCount) : baselineWordCount;
    const catchUpBaseline = shouldUseAppWordCount
      ? {
          ...baseline,
          projectId: boundProject.id,
          project: boundProject,
          endDocumentWordCount: catchUpStartWordCount,
          endDocumentWordCounts: null
        }
      : baseline;

    if (suspiciousZeroWordCount) {
      return {
        candidate: {
          documentId,
          documentUrl: aceDocumentUrl(),
          baseline: catchUpBaseline,
          currentSnapshot,
          startDocumentWordCount: catchUpStartWordCount,
          endDocumentWordCount: null,
          endDocumentWordCounts: null,
          wordsWritten: 0,
          wordsAdded: 0,
          wordsRemoved: 0,
          wordsEdited: 0,
          netWordsChanged: 0,
          sessionType: "writing",
          needsWordCountConfirmation: true
        },
        error: ""
      };
    }

    const netWordsChanged = currentWordCount - catchUpStartWordCount;
    const wordsAdded = shouldUseAppWordCount
      ? Math.max(0, netWordsChanged)
      : Math.max(0, Number(currentSnapshot.wordsAdded) || 0);
    const wordsRemoved = shouldUseAppWordCount
      ? Math.max(0, -netWordsChanged)
      : Math.max(0, Number(currentSnapshot.wordsRemoved) || 0);
    const wordsEdited = wordsAdded + wordsRemoved;
    if (wordsEdited <= 0) {
      return { candidate: null, error: "" };
    }

    return {
      candidate: {
        documentId,
        documentUrl: aceDocumentUrl(),
        baseline: catchUpBaseline,
        currentSnapshot,
        startDocumentWordCount: catchUpStartWordCount,
        endDocumentWordCount: currentWordCount,
        endDocumentWordCounts: currentSnapshot.wordCounts || null,
        endDocumentWordTokens: currentSnapshot.wordTokens || null,
        wordsWritten: wordsRemoved > 0 ? 0 : wordsAdded,
        wordsAdded,
        wordsRemoved,
        wordsEdited,
        netWordsChanged,
        sessionType: wordsRemoved > 0 ? "editing" : "writing"
      },
      error: ""
    };
  }

  async function aceCheckCatchUpBeforeStartPrompt() {
    aceState = "checking-catch-up";
    acePromptError = "";
    aceRenderLoading("Checking progress...", "Looking for missed words.");
    await aceNextFrame();

    const catchUpResult = await aceBuildCatchUpCandidate(aceExtractDocumentId());
    if (catchUpResult.error) {
      acePromptError = catchUpResult.error;
      aceState = "idle";
      aceShowStartPrompt();
      return;
    }

    if (!catchUpResult.candidate) {
      aceState = "idle";
      aceShowStartPrompt();
      return;
    }

    aceCatchUpCandidate = catchUpResult.candidate;
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

  async function aceActivateSessionFromSnapshot({
    documentId,
    extensionSessionId,
    sessionType,
    startSnapshot,
    projectId = "",
    hadDocumentActivity = false
  }) {
    const now = new Date().toISOString();
    acePromptError = "";
    aceState = "active";
    aceActiveSession = {
      startedAt: now,
      sessionType: aceNormalizeSessionType(sessionType),
      documentId,
      projectId: projectId || "",
      documentUrl: aceDocumentUrl(),
      extensionSessionId,
      wordsWritten: 0,
      wordsEdited: 0,
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged: 0,
      hadDocumentActivity: Boolean(hadDocumentActivity),
      startDocumentWordCount: Number.isFinite(startSnapshot.wordCount)
        ? startSnapshot.wordCount
        : null,
      startDocumentRevisionId: startSnapshot.revisionId || "",
      startDocumentWordCountTokenizerVersion: startSnapshot.wordCountTokenizerVersion || "",
      startDocumentHasWordTokens: aceSnapshotHasExactWordTokens(startSnapshot),
      wordCountMethod: startSnapshot.method || "google-docs-api",
      wordCountError: "",
      wordCountDiagnostic: startSnapshot.wordCountDiagnostic || ""
    };
    aceClearActivityTimers();
    await acePersistActiveSession();
    await aceRememberLastSessionType(sessionType);
    console.info("[ACE] SESSION START", {
      sessionType,
      visibleWordCount: startSnapshot.visibleWordCount ?? null,
      apiWordCount: startSnapshot.apiWordCount ?? startSnapshot.wordCount ?? null,
      revisionId: startSnapshot.revisionId || "",
      tokenCount: Array.isArray(startSnapshot.wordTokens) ? startSnapshot.wordTokens.length : null,
      measurementPath: aceMeasurementPathForSession(aceActiveSession)
    });
    aceStartTimer();
    if (!projectId) {
      aceRunAsync(aceResolveBindingForActiveSession(), "resolve active session binding");
    }
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
    const catchUpResult = await aceBuildCatchUpCandidate(documentId);
    if (catchUpResult.candidate) {
      aceCatchUpCandidate = catchUpResult.candidate;
      aceSyncStatus = "";
      aceSyncMessage = "";
      aceState = "catch-up";
      aceRenderCatchUpPrompt();
      return;
    }

    const extensionSessionId = aceCreateExtensionSessionId(documentId);
    const sessionType = await aceLastSessionType();
    const requiresExactWordDiff = aceSessionTypeRequiresExactWordDiff(sessionType);
    const startSnapshot = await aceStartSnapshotWithVisibleFallback(
      documentId,
      extensionSessionId,
      true,
      "manual start",
      { allowVisibleFallback: !requiresExactWordDiff }
    );
    if (!startSnapshot.ok || !Number.isFinite(startSnapshot.wordCount)) {
      aceState = "prompt";
      acePromptError = `Google Docs word count is required. ${startSnapshot.error || "Check Google OAuth and try again."}`;
      aceRenderPrompt();
      return;
    }
    if (requiresExactWordDiff && !aceSnapshotHasExactWordTokens(startSnapshot)) {
      aceState = "prompt";
      acePromptError = aceExactWordDiffStartError(startSnapshot);
      aceRenderPrompt();
      return;
    }

    await aceActivateSessionFromSnapshot({
      documentId,
      extensionSessionId,
      sessionType,
      startSnapshot
    });
  }

  function aceCanAutoStartBoundSession() {
    return ACE_IS_TOP_FRAME
      && aceState === "idle"
      && !aceActiveSession
      && !aceCompletedSession
      && !aceCatchUpCandidate
      && !aceAutoStartInFlight;
  }

  async function aceAutoStartBoundSessionFromActivity() {
    if (!aceCanAutoStartBoundSession()) {
      return;
    }

    const documentId = aceExtractDocumentId();
    if (!documentId) {
      return;
    }

    aceAutoStartInFlight = true;
    try {
      const binding = await aceGetBoundProjectForDocument(documentId);
      if (
        !binding?.projectId
        || aceState !== "idle"
        || aceActiveSession
        || aceCompletedSession
        || aceCatchUpCandidate
      ) {
        return;
      }

      aceState = "starting";
      acePromptError = "";
      aceRenderLoading("Starting session...", "Connected Scriptor project found.");
      await aceNextFrame();

      const extensionSessionId = aceCreateExtensionSessionId(documentId);
      const sessionType = await aceLastSessionType();
      const requiresExactWordDiff = aceSessionTypeRequiresExactWordDiff(sessionType);
      const baseline = await aceGetDocumentBaseline(documentId);
      let startSnapshot = requiresExactWordDiff
        ? null
        : await aceSeedGoogleDocStartSnapshotFromBaseline(
          documentId,
          extensionSessionId,
          baseline
        );
      if (!startSnapshot) {
        startSnapshot = await aceStartSnapshotWithVisibleFallback(
          documentId,
          extensionSessionId,
          false,
          "auto-start",
          { allowVisibleFallback: !requiresExactWordDiff }
        );
      }

      if (!startSnapshot.ok || !Number.isFinite(startSnapshot.wordCount)) {
        aceState = "prompt";
        acePromptError = `Could not auto-start session. ${startSnapshot.error || "Check Google OAuth and try again."}`;
        aceRenderPrompt();
        return;
      }
      if (requiresExactWordDiff && !aceSnapshotHasExactWordTokens(startSnapshot)) {
        aceState = "prompt";
        acePromptError = aceExactWordDiffStartError(startSnapshot);
        aceRenderPrompt();
        return;
      }

      const suspiciousZeroWordCount = Number(baseline?.endDocumentWordCount) >= 100
        && Number(startSnapshot.apiWordCount ?? startSnapshot.wordCount) === 0
        && !Number.isFinite(startSnapshot.visibleWordCount);
      if (suspiciousZeroWordCount) {
        aceState = "prompt";
        acePromptError = "Could not auto-start session because Google Docs returned 0 words.";
        aceRenderPrompt();
        return;
      }

      await aceActivateSessionFromSnapshot({
        documentId,
        extensionSessionId,
        sessionType,
        startSnapshot,
        projectId: binding.projectId,
        hadDocumentActivity: true
      });
    } catch (error) {
      if (aceState === "starting") {
        aceState = "prompt";
        acePromptError = `Could not auto-start session. ${error.message || "Try starting manually."}`;
        aceRenderPrompt();
      }
    } finally {
      aceAutoStartInFlight = false;
    }
  }

  async function aceTrackOrPromptFromActivity() {
    if (!aceCanAutoStartBoundSession()) {
      return;
    }

    await aceAutoStartBoundSessionFromActivity();
    if (
      aceState === "idle"
      && !aceActiveSession
      && !aceCompletedSession
      && !aceCatchUpCandidate
    ) {
      aceShowStartPrompt();
    }
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

    const sessionType = aceActiveSession.sessionType === "writing" ? "editing" : "writing";
    aceActiveSession = {
      ...aceActiveSession,
      sessionType,
      wordsWritten: Math.max(0, Number(aceActiveSession.wordsWritten) || 0),
      wordsEdited: Math.max(0, Number(aceActiveSession.wordsEdited) || 0),
      wordsAdded: Math.max(0, Number(aceActiveSession.wordsAdded) || 0),
      wordsRemoved: Math.max(0, Number(aceActiveSession.wordsRemoved) || 0),
      netWordsChanged: Number(aceActiveSession.netWordsChanged) || 0
    };
    await acePersistActiveSession();
    await aceRememberLastSessionType(sessionType);
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
      activeSession.hadDocumentActivity,
      {
        requireExactWordDiff: aceSessionTypeRequiresExactWordDiff(activeSession.sessionType),
        allowVisibleFallback: !aceSessionTypeRequiresExactWordDiff(activeSession.sessionType)
      }
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
    const wordCountDiagnostic = wordDiff.wordCountDiagnostic || activeSession.wordCountDiagnostic || aceGoogleDocDiffDiagnostic({
      code: measurementPending ? "E-API-UNAVAILABLE" : "D-API-DIFF",
      startWordCount: activeSession.startDocumentWordCount,
      apiEndWordCount: endDocumentWordCount,
      visibleEndWordCount: wordDiff.visibleWordCount,
      wordsAdded,
      wordsRemoved,
      netWordsChanged,
      revisionChanged: Boolean(
        activeSession.startDocumentRevisionId
        && wordDiff.revisionId
        && wordDiff.revisionId !== activeSession.startDocumentRevisionId
      ),
      startRevisionId: activeSession.startDocumentRevisionId || "",
      endRevisionId: wordDiff.revisionId || "",
      source: wordDiff.wordDiffMethod || (wordDiff.endWordTokens ? "google-api-token-sequence" : "api-total-fallback")
    });
    console.info("[ACE] SESSION END", {
      sessionType: activeSession.sessionType,
      visibleWordCount: wordDiff.visibleWordCount ?? null,
      apiWordCount: wordDiff.apiWordCount ?? wordDiff.wordCount ?? null,
      revisionId: wordDiff.revisionId || "",
      tokenCount: Array.isArray(wordDiff.endWordTokens) ? wordDiff.endWordTokens.length : null,
      measurementPath: wordDiff.wordDiffMethod || (measurementPending ? "measurement-unavailable" : "exact-api-sequence-diff")
    });
    console.info("[ACE] DIFF RESULT", {
      measurementPath: wordDiff.wordDiffMethod || (measurementPending ? "measurement-unavailable" : "exact-api-sequence-diff"),
      wordsAdded,
      wordsRemoved,
      wordsEdited: measuredWordsEdited,
      netWordsChanged
    });
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
      endDocumentWordCounts: wordDiff.endWordCounts || null,
      endDocumentWordTokens: wordDiff.endWordTokens || null,
      endDocumentRevisionId: wordDiff.revisionId || "",
      wordCountTokenizerVersion: wordDiff.wordCountTokenizerVersion || "",
      wordDiffMethod: wordDiff.wordDiffMethod || "",
      wordCountMethod: "google-docs-api",
      wordCountError,
      wordCountDiagnostic,
      hadDocumentActivity: Boolean(activeSession.hadDocumentActivity),
      measurementPending
    };

    aceState = "completed";
    aceSyncStatus = "syncing";
    aceSyncMessage = "Syncing...";
    aceSelectedProject = null;
    aceActiveSession = null;
    await aceRememberLastSessionType(activeSession.sessionType);
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
      wordCountDiagnostic: activeSession.wordCountDiagnostic || "E-DOC-CLOSED: The document closed before the extension could take the Google API after snapshot.",
      hadDocumentActivity: Boolean(activeSession.hadDocumentActivity),
      measurementPending: true
    };

    aceCompletedSession = session;
    aceActiveSession = null;
    await aceRememberLastSessionType(activeSession.sessionType);
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
      let binding = await aceGetLocalDocumentBinding(documentId);
      if (!binding?.projectId) {
        const serverProject = await aceGetServerDocumentBinding(documentId);
        if (serverProject?.id) {
          await aceSaveLocalDocumentBinding(documentId, serverProject);
          binding = {
            projectId: serverProject.id,
            project: serverProject
          };
        }
      }
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

        let binding = await aceGetLocalDocumentBinding(documentId);
        if (!binding?.projectId) {
          const serverProject = await aceGetServerDocumentBinding(documentId);
          if (serverProject?.id) {
            await aceSaveLocalDocumentBinding(documentId, serverProject);
            binding = {
              projectId: serverProject.id,
              project: serverProject
            };
          }
        }
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
    const originalCatchUpCandidate = aceCatchUpCandidate;
    const manualEndWordCount = aceReadCatchUpEndWordCount();
    if (!originalCatchUpCandidate) {
      return;
    }

    if (!Number.isFinite(manualEndWordCount)) {
      aceSyncStatus = "pending";
      aceSyncMessage = "Enter the current Google Docs word count before logging catch-up.";
      aceState = "catch-up";
      aceRenderCatchUpPrompt();
      return;
    }

    const adjustedCatchUpCandidate = aceRecalculateCatchUpCandidate(originalCatchUpCandidate, manualEndWordCount);
    const unsafeMessage = aceUnsafeCatchUpMessage(adjustedCatchUpCandidate);
    if (unsafeMessage) {
      aceCatchUpCandidate = adjustedCatchUpCandidate;
      aceSyncStatus = "pending";
      aceSyncMessage = unsafeMessage;
      aceState = "catch-up";
      aceRenderCatchUpPrompt();
      return;
    }

    if (adjustedCatchUpCandidate.wordsEdited <= 0) {
      aceCatchUpCandidate = null;
      aceSyncStatus = "";
      aceSyncMessage = "";
      aceState = "idle";
      aceShowStartPrompt();
      return;
    }

    aceCatchUpCandidate = adjustedCatchUpCandidate;
    const catchUpCandidate = adjustedCatchUpCandidate;
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
    const sessionType = catchUpCandidate.sessionType === "editing" ? "editing" : "writing";
    const wordsAdded = Math.max(0, Number(catchUpCandidate.wordsAdded) || 0);
    const wordsRemoved = Math.max(0, Number(catchUpCandidate.wordsRemoved) || 0);
    const wordsWritten = sessionType === "writing"
      ? Math.max(0, Number(catchUpCandidate.wordsWritten) || wordsAdded)
      : 0;
    const wordsEdited = sessionType === "editing" ? wordsAdded + wordsRemoved : 0;
    const catchUpSession = {
      documentId: catchUpCandidate.documentId,
      projectId,
      sessionType,
      startedAt,
      endedAt,
      durationMinutes: 1,
      source: "chrome-extension",
      documentUrl: catchUpCandidate.documentUrl || aceDocumentUrl(),
      notes: sessionType === "editing"
        ? "Catch-up: words edited outside a tracked session."
        : "Catch-up: words added outside a tracked session.",
      extensionSessionId: aceCreateExtensionSessionId(catchUpCandidate.documentId),
      wordsWritten,
      wordsEdited,
      wordsAdded,
      wordsRemoved,
      netWordsChanged: catchUpCandidate.netWordsChanged,
      startDocumentWordCount: catchUpCandidate.startDocumentWordCount,
      startDocumentRevisionId: catchUpCandidate.baseline?.revisionId || "",
      endDocumentWordCount: catchUpCandidate.endDocumentWordCount,
      endDocumentWordCounts: catchUpCandidate.endDocumentWordCounts || null,
      endDocumentWordTokens: catchUpCandidate.endDocumentWordTokens || null,
      endDocumentRevisionId: catchUpCandidate.currentSnapshot?.revisionId || "",
      wordCountTokenizerVersion: catchUpCandidate.currentSnapshot?.wordCountTokenizerVersion || "",
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
      console.info("[ACE] POST-SYNC RESPONSE", {
        measurementPath: aceMeasurementPathForSession(result.session || sessionToSync),
        wordsAdded: result.session?.wordsAdded,
        wordsRemoved: result.session?.wordsRemoved,
        wordsEdited: result.session?.wordsEdited,
        netWordsChanged: result.session?.netWordsChanged,
        duplicate: Boolean(result.duplicate)
      });
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
      completedSession.hadDocumentActivity,
      {
        requireExactWordDiff: aceSessionTypeRequiresExactWordDiff(completedSession.sessionType),
        allowVisibleFallback: !aceSessionTypeRequiresExactWordDiff(completedSession.sessionType)
      }
    );
    if (!aceIsCompletedSessionCurrent(completedSessionId)) {
      return false;
    }

    if (!wordDiff.ok || !Number.isFinite(wordDiff.wordCount)) {
      const visibleFallback = aceSessionTypeRequiresExactWordDiff(completedSession.sessionType)
        ? null
        : await aceVisibleStartSnapshot(
          completedSession.documentId,
          wordDiff.error || "retry Google API diff was unavailable"
        );
      if (visibleFallback && Number.isFinite(completedSession.startDocumentWordCount)) {
        const netWordsChanged = visibleFallback.wordCount - completedSession.startDocumentWordCount;
        const wordsAdded = Math.max(0, netWordsChanged);
        const wordsRemoved = Math.max(0, -netWordsChanged);
        aceCompletedSession = {
          ...aceCompletedSession,
          wordsWritten: completedSession.sessionType === "writing" ? Math.max(0, netWordsChanged) : 0,
          wordsEdited: completedSession.sessionType === "editing" ? wordsAdded + wordsRemoved : 0,
          wordsAdded,
          wordsRemoved,
          netWordsChanged,
          endDocumentWordCount: visibleFallback.wordCount,
          endDocumentWordCounts: null,
          endDocumentWordTokens: null,
          endDocumentRevisionId: "",
          wordCountTokenizerVersion: "",
          wordCountMethod: "visible-total-fallback",
          wordCountError: "",
          wordCountDiagnostic: aceGoogleDocDiffDiagnostic({
            code: "W-END-VISIBLE-FALLBACK",
            startWordCount: completedSession.startDocumentWordCount,
            apiEndWordCount: wordDiff.wordCount,
            visibleEndWordCount: visibleFallback.wordCount,
            wordsAdded,
            wordsRemoved,
            netWordsChanged,
            revisionChanged: false,
            startRevisionId: completedSession.startDocumentRevisionId || "",
            endRevisionId: "",
            source: "visible-total-fallback"
          }),
          measurementPending: false
        };
        return true;
      }

      aceCompletedSession = {
        ...aceCompletedSession,
        wordCountError: wordDiff.error || "Try again shortly.",
        wordCountDiagnostic: wordDiff.wordCountDiagnostic || aceGoogleDocDiffDiagnostic({
          code: "E-API-UNAVAILABLE",
          startWordCount: completedSession.startDocumentWordCount,
          apiEndWordCount: wordDiff.wordCount,
          visibleEndWordCount: wordDiff.visibleWordCount,
          wordsAdded: wordDiff.wordsAdded,
          wordsRemoved: wordDiff.wordsRemoved,
          netWordsChanged: wordDiff.netWordsChanged,
          revisionChanged: Boolean(
            completedSession.startDocumentRevisionId
            && wordDiff.revisionId
            && wordDiff.revisionId !== completedSession.startDocumentRevisionId
          ),
          startRevisionId: completedSession.startDocumentRevisionId || "",
          endRevisionId: wordDiff.revisionId || "",
          source: wordDiff.wordDiffMethod || (wordDiff.endWordTokens ? "google-api-token-sequence" : "api-total-fallback")
        })
      };
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
      endDocumentWordCounts: wordDiff.endWordCounts || null,
      endDocumentWordTokens: wordDiff.endWordTokens || null,
      endDocumentRevisionId: wordDiff.revisionId || "",
      wordCountTokenizerVersion: wordDiff.wordCountTokenizerVersion || "",
      wordDiffMethod: wordDiff.wordDiffMethod || "",
      wordCountMethod: "google-docs-api",
      wordCountError: "",
      wordCountDiagnostic: wordDiff.wordCountDiagnostic || aceGoogleDocDiffDiagnostic({
        code: "D-API-DIFF",
        startWordCount: completedSession.startDocumentWordCount,
        apiEndWordCount: wordDiff.wordCount,
        visibleEndWordCount: wordDiff.visibleWordCount,
        wordsAdded,
        wordsRemoved,
        netWordsChanged,
        revisionChanged: Boolean(
          completedSession.startDocumentRevisionId
          && wordDiff.revisionId
          && wordDiff.revisionId !== completedSession.startDocumentRevisionId
        ),
        startRevisionId: completedSession.startDocumentRevisionId || "",
        endRevisionId: wordDiff.revisionId || "",
        source: wordDiff.wordDiffMethod || (wordDiff.endWordTokens ? "google-api-token-sequence" : "api-total-fallback")
      }),
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
      aceRunAsync(aceTrackOrPromptFromActivity(), "track or prompt from document activity");
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
    } else if (action === "show-issue-detail") {
      aceRenderIssueDetail(button.getAttribute("data-ace-issue-id"));
    } else if (action === "back-to-issues") {
      aceRenderIssuesList();
    } else if (action === "copy-quote") {
      aceRunAsync(aceCopyIssueQuote(button.getAttribute("data-ace-issue-id")), "copy quote");
    } else if (action === "copy-quote-detail") {
      aceRunAsync(aceCopyIssueQuote(button.getAttribute("data-ace-issue-id"), true), "copy quote");
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
    } else if (action === "open-edit-dashboard") {
      aceOpenEditDashboard();
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
    const catchUpInput = aceClosest(event.target, "[data-ace-catch-up-word-count]");
    if (catchUpInput && aceWidget.contains(catchUpInput) && aceCatchUpCandidate) {
      const wordCount = aceParseWordCountInput(catchUpInput.value);
      const currentWordCountPreview = aceWidget.querySelector("[data-ace-current-word-count-preview]");
      if (currentWordCountPreview) {
        currentWordCountPreview.textContent = Number.isFinite(wordCount)
          ? aceFormatNumber(wordCount)
          : "an unknown number of";
      }

      const preview = aceWidget.querySelector("[data-ace-catch-up-preview]");
      if (preview) {
        preview.textContent = Number.isFinite(wordCount)
          ? aceCatchUpActivity(aceRecalculateCatchUpCandidate(aceCatchUpCandidate, wordCount)).activityCopy
          : "an unknown number of words";
      }

      aceSyncStatus = "";
      aceSyncMessage = "";
      const status = aceWidget.querySelector(".ace-sync-copy--pending");
      if (status) {
        status.remove();
      }
      return;
    }

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

  if (globalThis.__ACE_TEST_EXPORTS__) {
    Object.assign(globalThis.__ACE_TEST_EXPORTS__, {
      aceSessionSyncPayload,
      aceWordChangeBreakdownForSync,
      aceMeasurementPathForSession,
      aceSessionWordsCopy,
      aceSessionTypeRequiresExactWordDiff
    });
  }

  document.addEventListener("beforeinput", aceHandleInputLikeActivity, true);
  document.addEventListener("input", aceHandleInputLikeActivity, true);
  document.addEventListener("keydown", aceHandleKeydown, true);
  document.addEventListener("paste", aceHandleClipboardActivity, true);
  document.addEventListener("cut", aceHandleClipboardActivity, true);

  window.addEventListener("message", function (event) {
    if (event.data?.aceType !== ACE_VISIBLE_WORD_COUNT_MESSAGE) {
      return;
    }

    aceHandleVisibleWordCountMessage(event.data, event.source);
  });

  if (ACE_IS_TOP_FRAME) {
    window.addEventListener("message", function (event) {
      if (event.source === window || !event.data || event.data.aceType !== ACE_ACTIVITY_MESSAGE) {
        return;
      }

      aceNoteActiveDocumentActivity();
      aceRunAsync(aceTrackOrPromptFromActivity(), "track or prompt from document activity");
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

    aceRunAsync(aceInitializeSessionState(), "initialize session state");
  }

  async function aceInitializeSessionState() {
    await aceRestoreSession();
    await acePrimeBaselineForCurrentDocument();
  }

  function aceHandleDocumentExit() {
    if (aceExitHandled || !aceActiveSession) {
      return;
    }

    aceExitHandled = true;
    aceRunAsync(aceEndSession({ fromUnload: true }), "auto-end session on document exit");
  }
})();
