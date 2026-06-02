(function () {
  "use strict";

  const ACE_API_BASE_URL = "https://davishedrick.pythonanywhere.com";
  const ACE_APP_URL = ACE_API_BASE_URL;
  const ACE_WIDGET_ID = "ace-widget";
  const ACE_TIMER_INTERVAL_MS = 1000;
  const ACE_GOOGLE_DOC_SETTLE_DELAY_MS = 500;
  const ACE_GOOGLE_DOC_POLL_ATTEMPTS = 1;
  const ACE_VISIBLE_STABLE_DELAY_MS = 300;
  const ACE_VISIBLE_STABLE_TIMEOUT_MS = 2000;
  const ACE_REFRESH_VISIBLE_STABLE_TIMEOUT_MS = 4500;
  const ACE_GOOGLE_DOC_API_TIMEOUT_MS = 2000;
  const ACE_ACTIVE_SESSION_HEARTBEAT_MS = 15000;
  const ACE_ACTIVITY_MESSAGE = "ace-writing-activity";
  const ACE_VISIBLE_WORD_COUNT_MESSAGE = "ace-visible-word-count";
  const ACE_VISIBLE_WORD_COUNT_RESULT_MESSAGE = "ace-visible-word-count-result";
  const ACE_GOOGLE_DOC_WORD_COUNT_MESSAGE = "ace-google-doc-word-count";
  const ACE_GOOGLE_DOC_START_SNAPSHOT_MESSAGE = "ace-google-doc-start-snapshot";
  const ACE_GOOGLE_DOC_NET_COUNT_MESSAGE = "ace-google-doc-net-count";
  const ACE_CURRENT_TAB_SCOPE_MESSAGE = "ace-current-tab-scope";
  const ACE_WORD_SNAPSHOT_STORAGE_PREFIX = "aceWordSnapshot:";
  const ACE_AUTO_START_UNBOUND_RECHECK_MS = 30000;
  const ACE_WORD_TOKENIZER_VERSION = "google-docs-like-v3";
  const ACE_IS_TOP_FRAME = window.top === window;
  const ACE_ISSUE_TITLE_WORD_LIMIT = 8;

  const ACE_SESSION_STORAGE = {
    pageInstanceId: "ace-page-instance-id",
    widgetPosition: "ace-widget-position"
  };

  const ACE_LOCAL_STORAGE = {
    activeSession: "aceActiveSession",
    abandonedSessions: "aceAbandonedSessions",
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
  const ACE_STALE_BINDING_STATUSES = new Set([
    "stale_missing_doc",
    "stale_inaccessible",
    "stale_missing_tab"
  ]);
  const ACE_TYPING_CATCH_UP_SUPPRESSION_MS = 15000;
  const ACE_CATCH_UP_BOUNDARY_TRIGGERS = new Set([
    "bind",
    "pre-session",
    "manual-sync"
  ]);

  let aceState = "idle";
  let aceActiveSession = null;
  let aceCompletedSession = null;
  let aceTimerId = null;
  let aceProjects = [];
  let acePendingClearStaleBinding = null;
  let acePendingDeletedBindingRebind = null;
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
  let aceRecoveryCandidate = null;
  let aceIssueDraft = null;
  let aceIssueReturnState = "idle";
  let aceCurrentIssues = [];
  let aceIssueStatus = "";
  let aceAutoStartInFlight = false;
  let aceAutoStartBindingMisses = {};
  let aceBaselinePrimeInFlight = false;
  let aceCurrentSurface = null;
  let aceCurrentBinding = null;
  let aceCreateProjectDraft = null;
  let aceCreateProjectStep = 0;
  let aceCreateProjectError = "";
  let aceLastResolvedSurfaceId = "";
  let aceSurfaceRefreshInFlight = false;
  let aceSurfaceMonitorId = null;
  let aceLastTypingTimestamp = 0;
  let aceLastActiveHeartbeatAt = 0;
  let acePageInstanceId = sessionStorage.getItem(ACE_SESSION_STORAGE.pageInstanceId) || "";
  if (!acePageInstanceId) {
    acePageInstanceId = `ace-page-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
    try {
      sessionStorage.setItem(ACE_SESSION_STORAGE.pageInstanceId, acePageInstanceId);
    } catch (_error) {
      // Session storage can be unavailable in some embedded contexts; the in-memory id still scopes this page.
    }
  }

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

  function aceNormalizeTabId(tabId) {
    const normalized = String(tabId || "").trim();
    return normalized || "default";
  }

  function aceNormalizeTabTitle(tabTitle) {
    return String(tabTitle || "").replace(/\s+/g, " ").trim();
  }

  function aceCreateManuscriptSurfaceId(documentId, tabId) {
    const normalizedDocumentId = String(documentId || "").trim();
    if (!normalizedDocumentId) {
      return "";
    }
    return `${normalizedDocumentId}:${aceNormalizeTabId(tabId)}`;
  }

  function aceCurrentGoogleDocTabInfo() {
    const currentLocation = window.location || globalThis.location || {};
    const tabParamNames = ["tab", "tabId", "tab_id", "documentTabId"];
    const tabTitleParamNames = ["tabTitle", "tab_title"];
    const searchParams = [];
    const appendParams = function (value) {
      const text = String(value || "").replace(/^#/, "");
      if (!text) {
        return;
      }
      searchParams.push(new URLSearchParams(text.startsWith("?") ? text.slice(1) : text));
    };

    appendParams(currentLocation.search);
    appendParams(currentLocation.hash);
    appendParams(currentLocation.hash?.split("?")[1] || "");

    let tabId = "";
    let tabTitle = "";
    searchParams.some(function (params) {
      tabId = tabParamNames.map(function (name) {
        return params.get(name);
      }).find(Boolean) || "";
      tabTitle = tabTitleParamNames.map(function (name) {
        return params.get(name);
      }).find(Boolean) || "";
      return Boolean(tabId || tabTitle);
    });

    return {
      tabId: aceNormalizeTabId(tabId),
      tabTitle: aceNormalizeTabTitle(tabTitle)
    };
  }

  function aceSurfaceFromParts(parts = {}) {
    const documentId = String(parts.documentId || "").trim();
    const tabId = aceNormalizeTabId(parts.tabId);
    const tabTitle = aceNormalizeTabTitle(parts.tabTitle);
    const manuscriptSurfaceId = String(parts.manuscriptSurfaceId || "").trim()
      || aceCreateManuscriptSurfaceId(documentId, tabId);
    const manuscriptSurfaceLabel = aceNormalizeTabTitle(parts.manuscriptSurfaceLabel)
      || tabTitle
      || "Current manuscript";

    return {
      documentId,
      tabId,
      tabTitle,
      manuscriptSurfaceId,
      manuscriptSurfaceLabel
    };
  }

  function aceCurrentManuscriptSurface(documentId = aceExtractDocumentId()) {
    return aceSurfaceFromParts({
      documentId,
      ...aceCurrentGoogleDocTabInfo()
    });
  }

  function aceNormalizeChromeTabId(tabId) {
    if (tabId === null || tabId === undefined || tabId === "") {
      return "";
    }
    return String(tabId).trim();
  }

  function aceCreateSessionScope(surfaceOrParts = {}, details = {}) {
    const surface = aceSurfaceFromParts(surfaceOrParts);
    return {
      projectId: String(details.projectId || surfaceOrParts.projectId || "").trim(),
      documentId: surface.documentId,
      tabId: surface.tabId,
      tabTitle: surface.tabTitle,
      manuscriptSurfaceId: surface.manuscriptSurfaceId,
      manuscriptSurfaceLabel: surface.manuscriptSurfaceLabel,
      chromeTabId: aceNormalizeChromeTabId(details.chromeTabId ?? surfaceOrParts.chromeTabId),
      pageInstanceId: String(details.pageInstanceId || surfaceOrParts.pageInstanceId || acePageInstanceId || "").trim()
    };
  }

  function aceCurrentChromeTabScope() {
    return new Promise(function (resolve) {
      const runtime = aceChromeRuntime();
      if (!runtime?.runtime?.sendMessage) {
        resolve({ chromeTabId: "", frameId: null });
        return;
      }

      try {
        runtime.runtime.sendMessage({ aceType: ACE_CURRENT_TAB_SCOPE_MESSAGE }, function (response) {
          if (runtime.runtime.lastError || !response?.ok) {
            resolve({ chromeTabId: "", frameId: null });
            return;
          }
          resolve({
            chromeTabId: aceNormalizeChromeTabId(response.chromeTabId),
            frameId: response.frameId ?? null
          });
        });
      } catch (_error) {
        resolve({ chromeTabId: "", frameId: null });
      }
    });
  }

  async function aceGetCurrentDocumentScope(projectId = "") {
    const currentSurface = aceCurrentManuscriptSurface();
    const tabScope = await aceCurrentChromeTabScope();
    return aceCreateSessionScope(currentSurface, {
      projectId,
      chromeTabId: tabScope.chromeTabId,
      pageInstanceId: acePageInstanceId
    });
  }

  function aceSurfaceConfidence(surface) {
    return Boolean(surface?.documentId && surface?.manuscriptSurfaceId && surface?.tabId);
  }

  function aceLogTabDiagnostic(code, detail = {}) {
    console.info(`[ACE] ${code}`, detail);
  }

  function aceSurfaceDiagnostic(surface) {
    return {
      documentId: surface?.documentId || "",
      tabId: surface?.tabId || "",
      tabTitle: surface?.tabTitle || "",
      manuscriptSurfaceId: surface?.manuscriptSurfaceId || "",
      confident: aceSurfaceConfidence(surface)
    };
  }

  function aceSessionManuscriptSurface(session) {
    return aceSurfaceFromParts({
      documentId: session?.documentId || aceExtractDocumentId(),
      tabId: session?.tabId,
      tabTitle: session?.tabTitle,
      manuscriptSurfaceId: session?.manuscriptSurfaceId,
      manuscriptSurfaceLabel: session?.manuscriptSurfaceLabel
    });
  }

  function aceSessionScope(session) {
    const storedScope = session?.sessionScope || {};
    const surface = aceSessionManuscriptSurface({
      ...session,
      documentId: storedScope.documentId || session?.documentId,
      tabId: storedScope.tabId || session?.tabId,
      tabTitle: storedScope.tabTitle || session?.tabTitle,
      manuscriptSurfaceId: storedScope.manuscriptSurfaceId || session?.manuscriptSurfaceId,
      manuscriptSurfaceLabel: storedScope.manuscriptSurfaceLabel || session?.manuscriptSurfaceLabel
    });
    return aceCreateSessionScope(surface, {
      projectId: storedScope.projectId || session?.projectId || "",
      chromeTabId: storedScope.chromeTabId || session?.chromeTabId || "",
      pageInstanceId: storedScope.pageInstanceId || session?.pageInstanceId || ""
    });
  }

  function aceSessionScopeMismatchReason(sessionScope, currentScope) {
    if (!sessionScope?.documentId || !currentScope?.documentId) {
      return "document-unknown";
    }
    if (sessionScope.documentId !== currentScope.documentId) {
      return "document-mismatch";
    }
    if (!sessionScope.manuscriptSurfaceId || !currentScope.manuscriptSurfaceId) {
      return "surface-unknown";
    }
    if (sessionScope.manuscriptSurfaceId !== currentScope.manuscriptSurfaceId) {
      return "surface-mismatch";
    }
    if (
      sessionScope.chromeTabId
      && currentScope.chromeTabId
      && sessionScope.chromeTabId !== currentScope.chromeTabId
    ) {
      return "chrome-tab-mismatch";
    }
    if (sessionScope.projectId && currentScope.projectId && sessionScope.projectId !== currentScope.projectId) {
      return "project-mismatch";
    }
    return "";
  }

  function aceValidateSessionScope(session, currentScope = aceCreateSessionScope(aceCurrentManuscriptSurface())) {
    const sessionScope = aceSessionScope(session);
    const reason = aceSessionScopeMismatchReason(sessionScope, currentScope);
    const originalTab = sessionScope.tabTitle || sessionScope.manuscriptSurfaceLabel || "the original Google Docs tab";
    const message = reason
      ? `This session belongs to another Google Docs tab. Return to "${originalTab}" to end it.`
      : "";
    return {
      ok: !reason,
      reason,
      message,
      sessionScope,
      currentScope
    };
  }

  function aceSessionSurfaceMismatch(session, surface = aceCurrentManuscriptSurface()) {
    return !aceValidateSessionScope(session, aceCreateSessionScope(surface)).ok;
  }

  function aceRecordMatchesSurface(record, surface) {
    return Boolean(
      record?.manuscriptSurfaceId
      && surface?.manuscriptSurfaceId
      && record.manuscriptSurfaceId === surface.manuscriptSurfaceId
    );
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

    const pausedDurationMs = Math.max(0, Number(aceActiveSession.pausedDurationMs) || 0);
    const currentBlockedMs = aceActiveSession.tabBlockedAt
      ? Math.max(0, Date.now() - new Date(aceActiveSession.tabBlockedAt).getTime())
      : 0;
    return Math.max(0, Date.now() - new Date(aceActiveSession.startedAt).getTime() - pausedDurationMs - currentBlockedMs);
  }

  function aceElapsedMsForSession(session) {
    if (!session?.startedAt) {
      return 0;
    }
    const pausedDurationMs = Math.max(0, Number(session.pausedDurationMs) || 0);
    const currentBlockedMs = session.tabBlockedAt
      ? Math.max(0, Date.now() - new Date(session.tabBlockedAt).getTime())
      : 0;
    return Math.max(0, Date.now() - new Date(session.startedAt).getTime() - pausedDurationMs - currentBlockedMs);
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

  function aceDiagnosticWordCountCopy(value) {
    if (value === null || value === undefined || value === "") {
      return "unknown";
    }
    return Number.isFinite(Number(value)) ? Number(value) : "unknown";
  }

  function aceCalculateNetWordDelta(previousWordCount, currentWordCount) {
    const previous = Math.max(0, Number(previousWordCount) || 0);
    const current = Math.max(0, Number(currentWordCount) || 0);
    // Catch-up delta always moves forward in time: current count minus previous baseline.
    return current - previous;
  }

  function aceSessionNetWordsChanged(session) {
    const startWordCount = Number(session?.startDocumentWordCount);
    const endWordCount = Number(session?.endDocumentWordCount);
    if (
      session?.startDocumentWordCount !== null
      && session?.startDocumentWordCount !== undefined
      && session?.endDocumentWordCount !== null
      && session?.endDocumentWordCount !== undefined
      && Number.isFinite(startWordCount)
      && Number.isFinite(endWordCount)
    ) {
      return endWordCount - startWordCount;
    }

    const netWordsChanged = Number(session?.netWordsChanged);
    if (Number.isFinite(netWordsChanged)) {
      return netWordsChanged;
    }

    if (session?.sessionType === "editing") {
      return (Number(session?.wordsAdded) || 0) - (Number(session?.wordsRemoved) || 0);
    }

    return Number(session?.wordsWritten) || 0;
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

  async function aceGetExtensionProjects() {
    const payload = await aceApiFetch("/api/extension/projects");
    return Array.isArray(payload.projects) ? payload.projects : [];
  }

  async function aceCreateProject(payload) {
    const response = await aceApiFetch("/api/extension/projects", {
      method: "POST",
      body: JSON.stringify(payload)
    });
    return response?.project || null;
  }

  async function aceSaveBinding(surfaceOrDocumentId, projectId) {
    const surface = typeof surfaceOrDocumentId === "object"
      ? aceSurfaceFromParts(surfaceOrDocumentId)
      : aceCurrentManuscriptSurface(surfaceOrDocumentId);
    return aceApiFetch("/api/extension/document-binding", {
      method: "PUT",
      body: JSON.stringify({
        documentId: surface.documentId,
        tabId: surface.tabId,
        tabTitle: surface.tabTitle,
        manuscriptSurfaceId: surface.manuscriptSurfaceId,
        documentUrl: aceDocumentUrl(surface.documentId),
        projectId
      })
    });
  }

  async function aceDeleteBinding(surfaceOrDocumentId) {
    const surface = typeof surfaceOrDocumentId === "object"
      ? aceSurfaceFromParts(surfaceOrDocumentId)
      : aceCurrentManuscriptSurface(surfaceOrDocumentId);
    return aceApiFetch("/api/extension/document-binding", {
      method: "DELETE",
      body: JSON.stringify({
        documentId: surface.documentId,
        tabId: surface.tabId,
        manuscriptSurfaceId: surface.manuscriptSurfaceId
      })
    });
  }

  async function aceUpdateBindingStatus(surfaceOrDocumentId, status, staleReason = "") {
    const surface = typeof surfaceOrDocumentId === "object"
      ? aceSurfaceFromParts(surfaceOrDocumentId)
      : aceCurrentManuscriptSurface(surfaceOrDocumentId);
    return aceApiFetch("/api/extension/document-binding/status", {
      method: "PATCH",
      body: JSON.stringify({
        documentId: surface.documentId,
        tabId: surface.tabId,
        tabTitle: surface.tabTitle,
        manuscriptSurfaceId: surface.manuscriptSurfaceId,
        status,
        staleReason
      })
    });
  }

  async function aceGetServerDocumentBinding(surfaceOrDocumentId) {
    const result = await aceGetServerDocumentBindingResult(surfaceOrDocumentId);
    return result.ok ? result.project : null;
  }

  async function aceGetServerDocumentBindingResult(surfaceOrDocumentId) {
    const surface = typeof surfaceOrDocumentId === "object"
      ? aceSurfaceFromParts(surfaceOrDocumentId)
      : aceCurrentManuscriptSurface(surfaceOrDocumentId);
    const documentId = surface.documentId;
    if (!documentId) {
      return { ok: true, project: null };
    }
    try {
      const query = new URLSearchParams({
        documentId,
        tabId: surface.tabId,
        manuscriptSurfaceId: surface.manuscriptSurfaceId
      });
      const payload = await aceApiFetch(`/api/extension/document-binding?${query.toString()}`);
      return { ok: true, project: payload?.project || null };
    } catch (error) {
      return { ok: false, project: null, error };
    }
  }

  async function aceLocalDocumentBindings() {
    const stored = await aceStorageGet(ACE_LOCAL_STORAGE.documentBindings);
    const bindings = stored[ACE_LOCAL_STORAGE.documentBindings];
    return bindings && typeof bindings === "object" && !Array.isArray(bindings)
      ? bindings
      : {};
  }

  function aceRecordSurfaceKey(surfaceOrDocumentId) {
    const surface = typeof surfaceOrDocumentId === "object"
      ? aceSurfaceFromParts(surfaceOrDocumentId)
      : aceCurrentManuscriptSurface(surfaceOrDocumentId);
    return surface;
  }

  function aceSurfaceRecordWithMetadata(record, surface) {
    if (!record) {
      return null;
    }
    return {
      ...record,
      documentId: record.documentId || surface.documentId,
      tabId: record.tabId || surface.tabId,
      tabTitle: record.tabTitle || surface.tabTitle,
      manuscriptSurfaceId: record.manuscriptSurfaceId || surface.manuscriptSurfaceId,
      manuscriptSurfaceLabel: record.manuscriptSurfaceLabel || surface.manuscriptSurfaceLabel
    };
  }

  function aceFindSurfaceRecord(records, surfaceOrDocumentId) {
    const surface = aceRecordSurfaceKey(surfaceOrDocumentId);
    if (!surface.documentId) {
      return null;
    }

    if (surface.manuscriptSurfaceId && records?.[surface.manuscriptSurfaceId]) {
      return aceSurfaceRecordWithMetadata(records[surface.manuscriptSurfaceId], surface);
    }

    if (surface.tabId === "default" && records?.[surface.documentId]) {
      return aceSurfaceRecordWithMetadata(records[surface.documentId], surface);
    }

    return null;
  }

  async function aceGetLocalDocumentBinding(surfaceOrDocumentId) {
    const surface = aceRecordSurfaceKey(surfaceOrDocumentId);
    if (!surface.documentId) {
      return null;
    }

    const bindings = await aceLocalDocumentBindings();
    return aceFindSurfaceRecord(bindings, surface);
  }

  async function aceSaveLocalDocumentBinding(surfaceOrDocumentId, project) {
    const surface = aceRecordSurfaceKey(surfaceOrDocumentId);
    if (!surface.documentId || !surface.manuscriptSurfaceId || !project?.id) {
      return;
    }

    const bindings = await aceLocalDocumentBindings();
    const nextBindings = {};
    Object.entries(bindings).forEach(function ([key, binding]) {
      const existingProjectId = String(binding?.projectId || binding || "").trim();
      if (key !== surface.manuscriptSurfaceId && existingProjectId === project.id) {
        return;
      }
      nextBindings[key] = binding;
    });
    await aceStorageSet({
      [ACE_LOCAL_STORAGE.documentBindings]: {
        ...nextBindings,
        [surface.manuscriptSurfaceId]: {
          documentId: surface.documentId,
          tabId: surface.tabId,
          tabTitle: surface.tabTitle,
          manuscriptSurfaceId: surface.manuscriptSurfaceId,
          manuscriptSurfaceLabel: surface.manuscriptSurfaceLabel,
          projectId: project.id,
          project,
          updatedAt: new Date().toISOString()
        }
      }
    });
  }

  async function aceRemoveLocalDocumentBinding(surfaceOrDocumentId) {
    const surface = aceRecordSurfaceKey(surfaceOrDocumentId);
    if (!surface.documentId) {
      return;
    }

    const bindings = await aceLocalDocumentBindings();
    const nextBindings = { ...bindings };
    if (surface.manuscriptSurfaceId) {
      delete nextBindings[surface.manuscriptSurfaceId];
    }
    if (surface.tabId === "default") {
      delete nextBindings[surface.documentId];
    }
    await aceStorageSet({ [ACE_LOCAL_STORAGE.documentBindings]: nextBindings });
  }

  async function aceGetBoundProjectForDocument(surfaceOrDocumentId) {
    const surface = aceRecordSurfaceKey(surfaceOrDocumentId);
    if (!surface.documentId) {
      return null;
    }

    const localBinding = await aceGetLocalDocumentBinding(surface);
    const miss = aceAutoStartBindingMisses[surface.manuscriptSurfaceId || surface.documentId];
    if (!localBinding?.projectId && miss && Date.now() - miss < ACE_AUTO_START_UNBOUND_RECHECK_MS) {
      return null;
    }

    const serverBinding = await aceGetServerDocumentBindingResult(surface);
    if (!serverBinding.ok) {
      return localBinding?.projectId ? localBinding : null;
    }

    if (serverBinding.project?.id) {
      await aceSaveLocalDocumentBinding(surface, serverBinding.project);
      delete aceAutoStartBindingMisses[surface.manuscriptSurfaceId || surface.documentId];
      return {
        ...surface,
        projectId: serverBinding.project.id,
        project: serverBinding.project
      };
    }

    if (localBinding?.projectId) {
      await aceRemoveLocalDocumentBinding(surface);
    }

    aceAutoStartBindingMisses[surface.manuscriptSurfaceId || surface.documentId] = Date.now();
    return null;
  }

  function aceProjectPickerBinding(item) {
    const binding = item?.binding || item?.documentBinding || null;
    if (!binding || typeof binding !== "object") {
      return null;
    }
    return aceSurfaceFromParts(binding);
  }

  function aceDeletedBindingForProjectItem(item) {
    const deletedBinding = item?.deletedBinding || item?.project?.deletedBinding || null;
    if (deletedBinding && typeof deletedBinding === "object") {
      return deletedBinding;
    }
    if (aceIsStaleBinding(item) && item?.binding && typeof item.binding === "object") {
      return item.binding;
    }
    return null;
  }

  function aceBindingStatus(item) {
    return String(
      item?.bindingStatus
      || item?.binding?.status
      || item?.deletedBinding?.status
      || (item?.isBound || item?.projectId ? "active" : "unbound")
    ).trim() || "unbound";
  }

  function aceIsStaleBinding(item) {
    return ACE_STALE_BINDING_STATUSES.has(aceBindingStatus(item));
  }

  function aceProjectPickerStatusLabel(item) {
    const status = aceBindingStatus(item);
    if (status === "stale_missing_doc") {
      return "Bound · missing doc";
    }
    if (status === "stale_missing_tab") {
      return "Bound · missing tab";
    }
    if (status === "stale_inaccessible") {
      return "Bound · inaccessible";
    }
    if (item?.isBound) {
      return "Bound";
    }
    const project = item?.project || item;
    return project?.manuscriptType || project?.status || "Project";
  }

  function aceValidateBoundDocument(projectBinding, currentTabDocument = {}, apiResponse = null) {
    const binding = projectBinding?.binding || projectBinding?.documentBinding || projectBinding || null;
    const deletedBinding = projectBinding?.deletedBinding || null;
    const boundDocumentId = String(binding?.documentId || deletedBinding?.documentId || "").trim();
    const boundSurfaceId = String(binding?.manuscriptSurfaceId || deletedBinding?.manuscriptSurfaceId || "").trim();
    const currentDocumentId = String(currentTabDocument?.documentId || "").trim();
    const currentSurfaceId = String(currentTabDocument?.manuscriptSurfaceId || "").trim();
    const existingStatus = aceBindingStatus(projectBinding?.binding || projectBinding || {});
    const staleStatus = ACE_STALE_BINDING_STATUSES.has(existingStatus)
      ? existingStatus
      : "";

    if (!binding && !deletedBinding) {
      return {
        status: "unbound",
        reason: "no binding",
        bindingStatus: "unbound",
        boundDocumentId,
        currentDocumentId
      };
    }
    if (staleStatus || deletedBinding) {
      return {
        status: "deleted",
        reason: staleStatus || deletedBinding?.staleReason || "deleted binding marker",
        bindingStatus: staleStatus || deletedBinding?.status || "stale_missing_doc",
        boundDocumentId,
        currentDocumentId
      };
    }
    if (apiResponse) {
      const validation = aceClassifyBindingValidation(apiResponse);
      if (ACE_STALE_BINDING_STATUSES.has(validation.status)) {
        return {
          status: "deleted",
          reason: validation.staleReason || validation.status,
          bindingStatus: validation.status,
          boundDocumentId,
          currentDocumentId
        };
      }
      if (!validation.status) {
        return {
          status: "unknown",
          reason: validation.staleReason || "binding validation unavailable",
          bindingStatus: existingStatus,
          boundDocumentId,
          currentDocumentId
        };
      }
    }
    if (
      currentDocumentId
      && boundDocumentId
      && (
        boundDocumentId !== currentDocumentId
        || (currentSurfaceId && boundSurfaceId && boundSurfaceId !== currentSurfaceId)
      )
    ) {
      return {
        status: "mismatch",
        reason: "bound document differs from current tab",
        bindingStatus: existingStatus,
        boundDocumentId,
        currentDocumentId
      };
    }
    return {
      status: "valid",
      reason: "binding is accessible",
      bindingStatus: "active",
      boundDocumentId,
      currentDocumentId
    };
  }

  function aceClassifyBindingValidation(response) {
    if (response?.ok) {
      return { status: "active", staleReason: "" };
    }
    const status = Number(response?.status) || 0;
    const error = String(response?.error || "");
    if (status === 404 && error.includes("E-GOOGLE-DOC-TAB-NOT-FOUND")) {
      return { status: "stale_missing_tab", staleReason: error };
    }
    if (status === 404) {
      return { status: "stale_missing_doc", staleReason: error };
    }
    if (status === 403) {
      return { status: "stale_inaccessible", staleReason: error };
    }
    return { status: "", staleReason: error };
  }

  async function aceValidateProjectBindingItem(item) {
    const deletedBinding = aceDeletedBindingForProjectItem(item);
    if (!item?.isBound) {
      return deletedBinding
        ? {
            ...item,
            isBound: false,
            bindingStatus: aceBindingStatus(item),
            deletedBinding,
            binding: null
          }
        : item;
    }
    const binding = aceProjectPickerBinding(item);
    if (!binding?.documentId || !binding?.manuscriptSurfaceId) {
      return item;
    }

    const response = await aceGoogleDocWordCount(binding.documentId, false, binding);
    const validation = aceValidateBoundDocument({ ...item, binding }, null, response);
    if (validation.status === "unknown") {
      return item;
    }

    const previousStatus = aceBindingStatus(item);
    const isDeleted = validation.status === "deleted";
    const bindingStatus = isDeleted ? validation.bindingStatus : "active";
    const staleReason = isDeleted ? validation.reason : "";
    const nextDeletedBinding = isDeleted
      ? {
          ...binding,
          documentId: binding.documentId,
          tabId: binding.tabId,
          tabTitle: binding.tabTitle,
          title: binding.tabTitle,
          manuscriptSurfaceId: binding.manuscriptSurfaceId,
          documentUrl: binding.documentUrl || aceDocumentUrl(),
          url: binding.documentUrl || aceDocumentUrl(),
          status: bindingStatus,
          staleReason,
          detectedAt: new Date().toISOString()
        }
      : null;
    const nextItem = {
      ...item,
      isBound: !isDeleted,
      bindingStatus,
      staleReason,
      binding: isDeleted
        ? null
        : {
            ...(item.binding || {}),
            status: bindingStatus,
            staleReason
          },
      deletedBinding: nextDeletedBinding
    };
    if (isDeleted) {
      await aceRemoveLocalDocumentBinding(binding);
    }
    console.info("[ACE] BINDING VALIDATION", {
      projectId: item?.project?.id || item?.id || "",
      previousBoundDocumentId: binding.documentId,
      currentTabDocumentId: aceCurrentManuscriptSurface().documentId,
      validationStatus: validation.status,
      deletionReason: validation.reason,
      localStateUpdated: isDeleted,
      baselineScanRan: false
    });
    if (bindingStatus !== previousStatus || staleReason !== String(item.staleReason || item.binding?.staleReason || "")) {
      try {
        await aceUpdateBindingStatus(binding, bindingStatus, staleReason);
        console.info("[ACE] BINDING VALIDATION SERVER", {
          projectId: item?.project?.id || item?.id || "",
          previousBoundDocumentId: binding.documentId,
          validationStatus: validation.status,
          serverUpdateSucceeded: true
        });
      } catch (error) {
        console.warn("[ACE] Failed to update binding status", error);
      }
    }
    return nextItem;
  }

  async function aceReconcileProjectPickerBindings(projects) {
    const rows = Array.isArray(projects) ? projects : [];
    return Promise.all(rows.map(aceValidateProjectBindingItem));
  }

  async function aceClearStaleProjectBinding(item) {
    const deletedBinding = aceDeletedBindingForProjectItem(item);
    const binding = aceProjectPickerBinding(item) || (deletedBinding ? aceSurfaceFromParts(deletedBinding) : null);
    if (!binding?.documentId || !binding?.manuscriptSurfaceId) {
      throw new Error("Binding metadata is missing.");
    }
    await aceDeleteBinding(binding);
    await aceRemoveLocalDocumentBinding(binding);
  }

  async function aceLocalDocumentBaselines() {
    const stored = await aceStorageGet(ACE_LOCAL_STORAGE.documentBaselines);
    const baselines = stored[ACE_LOCAL_STORAGE.documentBaselines];
    return baselines && typeof baselines === "object" && !Array.isArray(baselines)
      ? baselines
      : {};
  }

  async function aceGetDocumentBaseline(surfaceOrDocumentId) {
    const surface = aceRecordSurfaceKey(surfaceOrDocumentId);
    if (!surface.documentId) {
      return null;
    }

    const baselines = await aceLocalDocumentBaselines();
    return aceFindSurfaceRecord(baselines, surface);
  }

  async function aceGetDocumentBaselineForCatchUp(surfaceOrDocumentId) {
    const surface = aceRecordSurfaceKey(surfaceOrDocumentId);
    if (!surface.documentId || !surface.manuscriptSurfaceId) {
      return { baseline: null, key: "", isLegacy: false };
    }

    const baselines = await aceLocalDocumentBaselines();
    if (baselines?.[surface.manuscriptSurfaceId]) {
      return {
        baseline: aceSurfaceRecordWithMetadata(baselines[surface.manuscriptSurfaceId], surface),
        key: surface.manuscriptSurfaceId,
        isLegacy: false
      };
    }

    if (surface.tabId === "default" && baselines?.[surface.documentId]) {
      return {
        baseline: aceSurfaceRecordWithMetadata(baselines[surface.documentId], surface),
        key: surface.documentId,
        isLegacy: true
      };
    }

    return { baseline: null, key: surface.manuscriptSurfaceId, isLegacy: false };
  }

  async function aceSaveDocumentBaseline(session, project) {
    const surface = aceSessionManuscriptSurface(session);
    const documentId = surface.documentId;
    const projectId = String(session?.projectId || project?.id || "").trim();
    const endWordCount = Number(session?.endDocumentWordCount);
    if (!documentId || !surface.manuscriptSurfaceId || !projectId || !Number.isFinite(endWordCount)) {
      return;
    }

    const baselines = await aceLocalDocumentBaselines();
    await aceStorageSet({
      [ACE_LOCAL_STORAGE.documentBaselines]: {
        ...baselines,
        [surface.manuscriptSurfaceId]: {
          documentId,
          tabId: surface.tabId,
          tabTitle: surface.tabTitle,
          manuscriptSurfaceId: surface.manuscriptSurfaceId,
          manuscriptSurfaceLabel: surface.manuscriptSurfaceLabel,
          projectId,
          project: project || null,
          endDocumentWordCount: Math.max(0, endWordCount),
          wordCountMethod: session?.wordCountMethod || "",
          endDocumentWordCountTokenizerVersion: session?.wordCountTokenizerVersion || session?.endDocumentWordCountTokenizerVersion || "",
          revisionId: session?.endDocumentRevisionId || session?.revisionId || "",
          syncedAt: new Date().toISOString(),
          sessionId: session.extensionSessionId || ""
        }
      }
    });
  }

  async function aceSaveGoogleDocBaseline(surfaceOrDocumentId, project, options = {}) {
    const surface = aceRecordSurfaceKey(surfaceOrDocumentId);
    const documentId = surface.documentId;
    const projectId = String(project?.id || "").trim();
    if (!documentId || !surface.manuscriptSurfaceId || !projectId) {
      return;
    }

    const stableVisible = await aceStableVisibleGoogleDocWordCount({
      timeoutMs: Number.isFinite(Number(options.visibleTimeoutMs))
        ? Number(options.visibleTimeoutMs)
        : ACE_VISIBLE_STABLE_TIMEOUT_MS,
      delayMs: Number.isFinite(Number(options.visibleDelayMs))
        ? Number(options.visibleDelayMs)
        : ACE_VISIBLE_STABLE_DELAY_MS,
      minStableReads: 2,
      readVisibleCount: options.readVisibleCount,
      context: {
        startSource: "binding-baseline",
        endSource: "stable-visible"
      }
    });
    const visibleWordCount = Number(stableVisible.count);
    const hasStableVisible = Boolean(stableVisible.stable && Number.isFinite(visibleWordCount));
    const snapshot = hasStableVisible
      ? {
          ok: true,
          wordCount: Math.max(0, visibleWordCount),
          method: "stable-visible-count",
          revisionId: "",
          wordCountTokenizerVersion: "",
          visibleCountDiagnostic: stableVisible
        }
      : await aceGoogleDocWordCount(documentId, false, surface);
    const wordCount = Number(snapshot?.wordCount);
    if (!snapshot.ok || !Number.isFinite(wordCount)) {
      return;
    }

    const baselines = await aceLocalDocumentBaselines();
    await aceStorageSet({
      [ACE_LOCAL_STORAGE.documentBaselines]: {
        ...baselines,
        [surface.manuscriptSurfaceId]: {
          documentId,
          tabId: surface.tabId,
          tabTitle: surface.tabTitle,
          manuscriptSurfaceId: surface.manuscriptSurfaceId,
          manuscriptSurfaceLabel: surface.manuscriptSurfaceLabel,
          projectId,
          project,
          endDocumentWordCount: Math.max(0, wordCount),
          wordCountMethod: snapshot.method || "google-docs-api",
          visibleCountDiagnostic: snapshot.visibleCountDiagnostic || null,
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
    const surface = aceCurrentManuscriptSurface(documentId);

    if (
      aceState !== "idle"
      || aceActiveSession
      || aceCompletedSession
      || aceCatchUpCandidate
    ) {
      return;
    }

    const existingBaseline = await aceGetDocumentBaseline(surface);
    if (Number.isFinite(Number(existingBaseline?.endDocumentWordCount))) {
      return;
    }

    aceBaselinePrimeInFlight = true;
    try {
      const binding = await aceGetBoundProjectForDocument(surface);
      if (!binding?.project) {
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

      await aceSaveGoogleDocBaseline(surface, binding.project);
    } catch (_error) {
      // A missing non-interactive Google token should not interrupt the document.
    } finally {
      aceBaselinePrimeInFlight = false;
    }
  }

  async function aceEnsureDocumentBaseline(surfaceOrDocumentId, project) {
    const existingBaseline = await aceGetDocumentBaseline(surfaceOrDocumentId);
    if (Number.isFinite(Number(existingBaseline?.endDocumentWordCount))) {
      return existingBaseline;
    }
    await aceSaveGoogleDocBaseline(surfaceOrDocumentId, project);
    return aceGetDocumentBaseline(surfaceOrDocumentId);
  }

  async function aceRefreshDocumentBaselineFromCurrentCount(surfaceOrDocumentId, project) {
    await aceSaveGoogleDocBaseline(surfaceOrDocumentId, project);
    return aceGetDocumentBaseline(surfaceOrDocumentId);
  }

  async function aceSaveBaselineFromCurrentSnapshot(surface, project, currentSnapshot) {
    if (!surface?.manuscriptSurfaceId || !project?.id || !currentSnapshot?.ok || !Number.isFinite(Number(currentSnapshot.wordCount))) {
      return null;
    }
    await aceSaveDocumentBaseline({
      ...surface,
      projectId: project.id,
      project,
      endDocumentWordCount: Math.max(0, Number(currentSnapshot.wordCount)),
      endDocumentRevisionId: currentSnapshot.revisionId || "",
      wordCountTokenizerVersion: currentSnapshot.wordCountTokenizerVersion || "",
      wordCountMethod: currentSnapshot.currentCountSource || currentSnapshot.method || "stable-visible-count"
    }, project);
    return aceGetDocumentBaseline(surface);
  }

  async function aceReconcileInitialBindBaseline(surface, project) {
    const existing = await aceGetDocumentBaselineForCatchUp(surface);
    const baseline = Number.isFinite(Number(existing.baseline?.endDocumentWordCount))
      ? existing.baseline
      : {
          ...surface,
          projectId: project.id,
          project,
          endDocumentWordCount: 0,
          syncedAt: "",
          sessionId: "",
          wordCountMethod: "initial-bind-zero"
        };
    const baselineKey = existing.key || surface.manuscriptSurfaceId;
    const baselineIsLegacy = Boolean(existing.isLegacy && existing.baseline);
    const currentSnapshot = await aceGoogleDocWordCountAfterSettle(surface.documentId, baseline, { trigger: "bind" });
    const currentWordCount = Math.max(0, Number(currentSnapshot?.wordCount) || 0);
    const baselineWordCount = Math.max(0, Number(baseline.endDocumentWordCount) || 0);
    const pendingSession = await acePendingSessionForCatchUpChange(surface, baselineWordCount, currentWordCount);
    const completedSession = aceSessionCoversCatchUpChange(aceCompletedSession, surface, baselineWordCount, currentWordCount)
      ? aceCompletedSession
      : null;
    const result = aceEvaluateCatchUpCandidate({
      trigger: "bind",
      surface,
      baseline: {
        ...baseline,
        projectId: project.id,
        project
      },
      baselineKey,
      baselineIsLegacy,
      currentSnapshot,
      pendingSession,
      completedSession,
      binding: {
        ...surface,
        projectId: project.id,
        project
      }
    });
    aceLogCatchUpDecision(result.trace);
    if (result.candidate) {
      aceCatchUpCandidate = result.candidate;
      aceSyncStatus = "";
      aceSyncMessage = "";
      aceState = "catch-up";
      aceRenderCatchUpPrompt();
      return { candidate: result.candidate, baselineUpdated: false };
    }
    if (currentSnapshot?.ok && Number.isFinite(Number(currentSnapshot.wordCount))) {
      await aceSaveBaselineFromCurrentSnapshot(surface, project, currentSnapshot);
      console.info("[ACE] CATCH-UP BASELINE ADVANCED", {
        code: "D-CATCHUP-BASELINE-ADVANCED",
        trigger: "bind",
        projectId: project.id,
        manuscriptSurfaceId: surface.manuscriptSurfaceId,
        baselineWordCount,
        currentWordCount,
        netWordDelta: aceCalculateNetWordDelta(baselineWordCount, currentWordCount),
        result: result.reason || "zero-net"
      });
    }
    return { candidate: null, baselineUpdated: Boolean(currentSnapshot?.ok) };
  }

  async function acePostSession(session) {
    const payload = aceSessionSyncPayload(session);
    console.info("[ACE] PRE-SYNC", {
      measurementPath: aceMeasurementPathForSession(payload),
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

  async function aceGetExtensionIssues(surfaceOrDocumentId) {
    const surface = aceRecordSurfaceKey(surfaceOrDocumentId);
    const query = new URLSearchParams({
      documentId: surface.documentId,
      tabId: surface.tabId,
      manuscriptSurfaceId: surface.manuscriptSurfaceId
    });
    return aceApiFetch(
      `/api/extension/issues?${query.toString()}`
    );
  }

  function aceIssueSyncPayload(draft, projectId, note, snippet) {
    const surface = aceSurfaceFromParts({
      documentId: draft?.documentId || aceExtractDocumentId(),
      tabId: draft?.tabId,
      tabTitle: draft?.tabTitle,
      manuscriptSurfaceId: draft?.manuscriptSurfaceId,
      manuscriptSurfaceLabel: draft?.manuscriptSurfaceLabel
    });
    return {
      documentId: surface.documentId,
      tabId: surface.tabId,
      tabTitle: surface.tabTitle,
      manuscriptSurfaceId: surface.manuscriptSurfaceId,
      manuscriptSurfaceLabel: surface.manuscriptSurfaceLabel,
      projectId,
      extensionIssueId: draft?.extensionIssueId,
      note,
      snippet,
      documentUrl: draft?.documentUrl || aceDocumentUrl(),
      source: "catch-up",
      quoteLocator: {
        strategy: "quote-finder",
        quote: snippet,
        createdAt: new Date().toISOString()
      }
    };
  }

  function aceSessionSyncPayload(session) {
    const sessionType = session?.sessionType === "editing" ? "editing" : "writing";
    const surface = aceSessionManuscriptSurface(session);
    const startWordCount = Number(session?.startDocumentWordCount);
    const endWordCount = Number(session?.endDocumentWordCount);
    const startDocumentWordCount = session?.startDocumentWordCount !== null
      && session?.startDocumentWordCount !== undefined
      && Number.isFinite(startWordCount)
      ? startWordCount
      : null;
    const endDocumentWordCount = session?.endDocumentWordCount !== null
      && session?.endDocumentWordCount !== undefined
      && Number.isFinite(endWordCount)
      ? endWordCount
      : null;
    const measuredNetWordsChanged = aceSessionNetWordsChanged({
      ...session,
      startDocumentWordCount,
      endDocumentWordCount
    });

    return {
      documentId: surface.documentId,
      tabId: surface.tabId,
      tabTitle: surface.tabTitle,
      manuscriptSurfaceId: surface.manuscriptSurfaceId,
      manuscriptSurfaceLabel: surface.manuscriptSurfaceLabel,
      chromeTabId: session?.chromeTabId || session?.sessionScope?.chromeTabId || "",
      sessionScope: session?.sessionScope || aceSessionScope(session),
      projectId: session?.projectId || "",
      sessionType,
      startedAt: session?.startedAt || "",
      endedAt: session?.endedAt || "",
      durationMinutes: Number(session?.durationMinutes) || 1,
      source: session?.source || "chrome-extension",
      documentUrl: session?.documentUrl || "",
      notes: session?.notes || "",
      extensionSessionId: session?.extensionSessionId || "",
      wordsWritten: sessionType === "writing"
        ? Math.max(0, Number(session?.wordsWritten) || measuredNetWordsChanged)
        : 0,
      wordsAdded: 0,
      wordsRemoved: 0,
      wordsEdited: 0,
      netWordsChanged: measuredNetWordsChanged,
      startDocumentWordCount,
      startDocumentRevisionId: session?.startDocumentRevisionId || "",
      endDocumentWordCount,
      endDocumentRevisionId: session?.endDocumentRevisionId || "",
      wordCountTokenizerVersion: session?.wordCountTokenizerVersion || "",
      wordCountMethod: session?.wordCountMethod || "google-docs-api",
      wordCountError: session?.wordCountError || "",
      wordCountDiagnostic: session?.wordCountDiagnostic || "",
      hadDocumentActivity: Boolean(session?.hadDocumentActivity),
      measurementPending: Boolean(session?.measurementPending)
    };
  }

  function aceMeasurementPathForSession(session) {
    if (session?.measurementPending) {
      return "measurement-unavailable";
    }
    if (session?.wordCountMethod === "visible-total-fallback") {
      return "visible-total-fallback";
    }
    if (session?.wordCountMethod === "visible-total-baseline") {
      return "saved-total-baseline";
    }
    if (session?.wordCountMethod === "stable-visible-count") {
      return "stable-visible-count";
    }
    return session?.wordCountMethod === "google-docs-api"
      ? "google-docs-net-count"
      : "measurement-unavailable";
  }

  function aceNowMs() {
    return Date.now();
  }

  function aceTimingElapsedMs(startedAt) {
    return Math.max(0, aceNowMs() - startedAt);
  }

  function aceCreateWordCountTiming(kind, trigger = "") {
    return {
      kind,
      trigger,
      startedAt: aceNowMs(),
      totalElapsedMs: 0,
      settleDelayMs: 0,
      visibleReadElapsedMs: 0,
      visibleReadCount: 0,
      frameScanElapsedMs: 0,
      stableVisibleElapsedMs: 0,
      apiElapsedMs: 0,
      apiAttempts: 0,
      apiPendingAtDecision: false,
      compareElapsedMs: 0,
      backendSyncElapsedMs: 0,
      finalSelectedCountSource: "",
      trustedReason: "",
      action: ""
    };
  }

  function aceApplyVisibleTiming(timing, stableVisible) {
    if (!timing || !stableVisible) {
      return;
    }

    const reads = Array.isArray(stableVisible.reads) ? stableVisible.reads : [];
    timing.visibleReadCount += reads.length;
    timing.visibleReadElapsedMs += reads.reduce(function (total, read) {
      return total + Math.max(0, Number(read?.elapsedMs) || 0);
    }, 0);
    timing.frameScanElapsedMs += reads.reduce(function (total, read) {
      return total + Math.max(0, Number(read?.frameScanElapsedMs) || 0);
    }, 0);
    timing.stableVisibleElapsedMs += Math.max(0, Number(stableVisible.elapsedMs) || 0);
  }

  function aceCompleteWordCountTiming(timing, updates = {}) {
    if (!timing) {
      return null;
    }

    Object.assign(timing, updates);
    timing.totalElapsedMs = aceTimingElapsedMs(timing.startedAt);
    console.info("[ACE] WORD COUNT TIMING", { ...timing });
    return timing;
  }

  function aceApiTimeoutResponse(timeoutMs) {
    return {
      ok: false,
      wordCount: null,
      apiWordCount: null,
      error: `E-API-TIMEOUT: Google Docs API did not return within ${timeoutMs}ms.`
    };
  }

  function aceWithTimeout(promise, timeoutMs, timeoutValue) {
    let timeoutId = null;
    const timeoutPromise = new Promise(function (resolve) {
      timeoutId = window.setTimeout(function () {
        resolve(timeoutValue);
      }, timeoutMs);
    });
    return Promise.race([promise, timeoutPromise]).then(function (result) {
      if (timeoutId) {
        window.clearTimeout(timeoutId);
      }
      return result;
    });
  }

  function aceStartTimedApiWordCount(apiCall, timing) {
    const startedAt = aceNowMs();
    if (timing) {
      timing.apiAttempts += 1;
    }
    return Promise.resolve()
      .then(apiCall)
      .then(function (response) {
        if (timing) {
          timing.apiElapsedMs += aceTimingElapsedMs(startedAt);
        }
        return response;
      })
      .catch(function (error) {
        if (timing) {
          timing.apiElapsedMs += aceTimingElapsedMs(startedAt);
        }
        return {
          ok: false,
          wordCount: null,
          apiWordCount: null,
          error: error?.message || "Google Docs API word count failed."
        };
      });
  }

  async function aceGoogleDocMessage(aceType, payload) {
    const apiStartedAt = aceNowMs();
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

    const apiElapsedMs = aceTimingElapsedMs(apiStartedAt);
    const visibleStartedAt = aceNowMs();
    const visibleWordCount = await aceVisibleGoogleDocWordCount();
    const visibleReadElapsedMs = aceTimingElapsedMs(visibleStartedAt);
    const responseWordCount = Number(response.wordCount);
    const netWordsChanged = Number(response.netWordsChanged);
    return {
      ...response,
      apiElapsedMs,
      revisionId: response.revisionId || "",
      startRevisionId: response.startRevisionId || "",
      wordCount: Number.isFinite(responseWordCount) ? Math.max(0, responseWordCount) : null,
      apiWordCount: Number.isFinite(responseWordCount) ? Math.max(0, responseWordCount) : null,
      visibleWordCount: Number.isFinite(visibleWordCount) ? Math.max(0, visibleWordCount) : null,
      visibleReadElapsedMs,
      wordCounts: response.wordCounts || null,
      wordTokens: response.wordTokens || null,
      endWordCounts: response.endWordCounts || null,
      endWordTokens: response.endWordTokens || null,
      wordCountTokenizerVersion: response.wordCountTokenizerVersion || "",
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged: Number.isFinite(netWordsChanged) ? netWordsChanged : 0
    };
  }

  async function aceVisibleGoogleDocWordCount() {
    const read = await aceVisibleGoogleDocWordCountRead();
    const best = read.selectedCandidate;
    return Number.isFinite(best?.count) ? best.count : null;
  }

  async function aceVisibleGoogleDocWordCountRead() {
    const startedAt = aceNowMs();
    const localStartedAt = aceNowMs();
    const localCandidates = aceVisibleGoogleDocWordCountCandidatesLocal();
    const localScanElapsedMs = aceTimingElapsedMs(localStartedAt);
    const frameStartedAt = aceNowMs();
    const frameCandidate = ACE_IS_TOP_FRAME
      ? await aceVisibleGoogleDocWordCountCandidateFromFrames()
      : null;
    const frameScanElapsedMs = ACE_IS_TOP_FRAME ? aceTimingElapsedMs(frameStartedAt) : 0;
    const candidates = [
      ...localCandidates,
      ...(frameCandidate ? [{ ...frameCandidate, source: frameCandidate.source || "frame-message" }] : [])
    ];
    const selectedCandidate = aceBestVisibleWordCountCandidate(candidates);
    return {
      count: Number.isFinite(selectedCandidate?.count) ? selectedCandidate.count : null,
      candidates,
      selectedCandidate,
      readAt: Date.now(),
      elapsedMs: aceTimingElapsedMs(startedAt),
      localScanElapsedMs,
      frameScanElapsedMs
    };
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
            ...candidate,
            count: Math.max(0, count),
            score: Number(candidate?.score) || 0,
            bottom: Number(candidate?.bottom) || 0,
            left: Number(candidate?.left) || 0,
            source: "frame-message"
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
    return aceBestVisibleWordCountCandidate(aceVisibleGoogleDocWordCountCandidatesLocal());
  }

  function aceVisibleGoogleDocWordCountCandidatesLocal() {
    const candidates = [];
    const visitedWindows = new Set();

    function collectFromWindow(targetWindow, depth = 0) {
      if (!targetWindow || visitedWindows.has(targetWindow) || depth > 4) {
        return;
      }

      visitedWindows.add(targetWindow);
      try {
        aceVisibleWordCountCandidatesInDocument(targetWindow.document, targetWindow).forEach(function (candidate) {
          candidates.push({
            ...candidate,
            frameDepth: depth,
            source: depth === 0 ? "active-frame" : "child-frame"
          });
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

    return candidates;
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

    const viewportWidth = targetWindow.innerWidth || targetDocument.documentElement?.clientWidth || 0;
    const viewportHeight = targetWindow.innerHeight || targetDocument.documentElement?.clientHeight || 0;
    const elements = new Map();
    Array.from(targetDocument.querySelectorAll("div, span, button, [aria-label], [role='button']"))
      .forEach(function (element) {
        elements.set(element, false);
      });

    aceBottomLeftViewportElements(targetDocument, viewportWidth, viewportHeight)
      .forEach(function (element) {
        elements.set(element, true);
      });

    return Array.from(elements.entries())
      .map(function (element) {
        const fromViewportProbe = Boolean(element[1]);
        element = element[0];
        if (element.closest?.("#ace-widget")) {
          return null;
        }

        const count = aceVisibleWordCountFromElement(element);
        if (!Number.isFinite(count)) {
          return null;
        }

        const rect = element.getBoundingClientRect();
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

        const snippet = aceVisibleWordCountSnippet(element);
        return {
          count,
          bottom: rect.bottom,
          left: rect.left,
          snippet,
          fromViewportProbe,
          score: (fromViewportProbe ? 2 : 0)
            + (viewportHeight ? rect.bottom / viewportHeight : 0)
            + (viewportWidth ? (1 - rect.left / viewportWidth) : 0)
        };
      })
      .filter(function (candidate) {
        return candidate && Number.isFinite(candidate.count);
      });
  }

  function aceBottomLeftViewportElements(targetDocument, viewportWidth, viewportHeight) {
    if (!targetDocument.elementsFromPoint || !viewportWidth || !viewportHeight) {
      return [];
    }

    const elements = new Set();
    const xPositions = [24, 80, 160, Math.min(280, viewportWidth * 0.36)];
    const yPositions = [
      Math.max(0, viewportHeight - 96),
      Math.max(0, viewportHeight - 56),
      Math.max(0, viewportHeight - 24)
    ];
    xPositions.forEach(function (x) {
      yPositions.forEach(function (y) {
        targetDocument.elementsFromPoint(x, y).forEach(function (element) {
          let current = element;
          let depth = 0;
          while (current && depth < 5) {
            elements.add(current);
            current = current.parentElement;
            depth += 1;
          }
        });
      });
    });
    return Array.from(elements);
  }

  function aceVisibleWordCountFromElement(element) {
    const texts = [
      element.textContent || "",
      element.innerText || "",
      element.getAttribute?.("aria-label") || "",
      element.getAttribute?.("title") || ""
    ];

    for (const rawText of texts) {
      const text = aceNormalizeIssueNoteText(rawText);
      if (!text || text.length > 120) {
        continue;
      }

      const match = text.match(/^([\d,]+)\s+words?$/iu);
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

  function aceVisibleWordCountSnippet(element) {
    const rawText = element.innerText
      || element.textContent
      || element.getAttribute?.("aria-label")
      || element.getAttribute?.("title")
      || "";
    return aceShortDiagnostic(aceNormalizeIssueNoteText(rawText), 80);
  }

  function aceVisibleCountClasses(result, context = {}) {
    const classes = [];
    const reads = Array.isArray(result?.reads) ? result.reads : [];
    const selectedCounts = reads
      .map(function (read) {
        return Number(read?.count);
      })
      .filter(Number.isFinite);
    const uniqueSelectedCounts = Array.from(new Set(selectedCounts));
    const candidateCounts = Array.from(new Set((result?.candidates || [])
      .map(function (candidate) {
        return Number(candidate?.count);
      })
      .filter(Number.isFinite)));
    const apiWordCount = Number(context.apiWordCount);
    const startWordCount = Number(context.startWordCount);
    const endWordCount = Number(result?.count);
    const durationMinutes = Number(context.durationMinutes);
    const netWordsChanged = Number.isFinite(startWordCount) && Number.isFinite(endWordCount)
      ? aceCalculateNetWordDelta(startWordCount, endWordCount)
      : Number(context.netWordsChanged);

    if (!Number.isFinite(endWordCount)) {
      classes.push("E-NO-VISIBLE-COUNT");
    } else if (result?.stable) {
      classes.push("D-STABLE-VISIBLE-COUNT");
    }
    if (uniqueSelectedCounts.length > 1) {
      classes.push("W-VISIBLE-COUNT-CHANGED");
    }
    if (candidateCounts.length > 1) {
      classes.push("W-VISIBLE-CANDIDATE-CONFLICT");
    }
    if (Number.isFinite(apiWordCount) && apiWordCount === 0 && endWordCount > 0) {
      classes.push("W-API-ZERO");
    }
    if (context.startSource && context.endSource && context.startSource !== context.endSource) {
      classes.push("W-SOURCE-MISMATCH");
    }
    if (
      Number.isFinite(durationMinutes)
      && durationMinutes > 0
      && Number.isFinite(netWordsChanged)
      && Math.abs(netWordsChanged) / Math.max(1, durationMinutes) > 250
    ) {
      classes.push("W-IMPLAUSIBLE-NET");
    }

    return classes;
  }

  function aceVisibleCountDiagnostic(result, context = {}) {
    if (!result) {
      return "E-NO-VISIBLE-COUNT: no visible word-count read.";
    }
    const classes = aceVisibleCountClasses(result, context);
    const readCounts = (result.reads || [])
      .map(function (read) {
        return Number(read?.count);
      })
      .filter(Number.isFinite);
    const uniqueCandidateCounts = Array.from(new Set((result.candidates || [])
      .map(function (candidate) {
        return Number(candidate?.count);
      })
      .filter(Number.isFinite)));
    const selected = Number.isFinite(Number(result.count)) ? Number(result.count) : "none";
    const changedCopy = readCounts.length > 1
      ? `visible reads: ${readCounts.join(" -> ")}`
      : `visible reads: ${readCounts.join("") || "none"}`;
    const candidatesCopy = uniqueCandidateCounts.length
      ? `visible candidates: ${uniqueCandidateCounts.join(", ")}`
      : "visible candidates: none";
    return `${classes.join(", ") || "D-VISIBLE-COUNT"}: ${changedCopy}; ${candidatesCopy}; selected ${selected}.`;
  }

  async function aceStableVisibleGoogleDocWordCount(options = {}) {
    const readVisibleCount = options.readVisibleCount || aceVisibleGoogleDocWordCountRead;
    const timeoutMs = Number.isFinite(Number(options.timeoutMs)) ? Number(options.timeoutMs) : ACE_VISIBLE_STABLE_TIMEOUT_MS;
    const delayMs = Number.isFinite(Number(options.delayMs)) ? Number(options.delayMs) : ACE_VISIBLE_STABLE_DELAY_MS;
    const minStableReads = Math.max(2, Number(options.minStableReads) || 2);
    const ignoreZero = Boolean(options.ignoreZero);
    const startedAt = aceNowMs();
    const reads = [];
    let stableCount = null;
    let stableRun = 0;

    while (aceNowMs() - startedAt <= timeoutMs || reads.length === 0) {
      const rawRead = await readVisibleCount();
      const read = typeof rawRead === "number"
        ? { count: rawRead, candidates: [{ count: rawRead }], selectedCandidate: { count: rawRead }, readAt: Date.now(), elapsedMs: 0, frameScanElapsedMs: 0 }
        : {
            ...rawRead,
            count: Number.isFinite(Number(rawRead?.count)) ? Math.max(0, Number(rawRead.count)) : null,
            readAt: rawRead?.readAt || Date.now(),
            elapsedMs: Math.max(0, Number(rawRead?.elapsedMs) || 0),
            frameScanElapsedMs: Math.max(0, Number(rawRead?.frameScanElapsedMs) || 0)
          };
      if (ignoreZero && read.count === 0) {
        read.count = null;
      }
      reads.push(read);

      if (Number.isFinite(read.count) && read.count === stableCount) {
        stableRun += 1;
      } else if (Number.isFinite(read.count)) {
        stableCount = read.count;
        stableRun = 1;
      }

      if (stableRun >= minStableReads) {
        break;
      }
      if (aceNowMs() - startedAt >= timeoutMs) {
        break;
      }
      await aceDelay(delayMs);
    }

    const latestValidRead = [...reads].reverse().find(function (read) {
      return Number.isFinite(read.count);
    }) || null;
    const candidates = reads.flatMap(function (read) {
      return Array.isArray(read.candidates) ? read.candidates : [];
    });
    const selectedCandidate = latestValidRead?.selectedCandidate
      || aceBestVisibleWordCountCandidate(candidates);
    const result = {
      count: Number.isFinite(latestValidRead?.count) ? latestValidRead.count : null,
      stable: stableRun >= minStableReads,
      reads,
      candidates,
      selectedCandidate,
      reason: stableRun >= minStableReads ? "stable" : "timeout",
      elapsedMs: aceTimingElapsedMs(startedAt),
      readCount: reads.length,
      frameScanElapsedMs: reads.reduce(function (total, read) {
        return total + Math.max(0, Number(read?.frameScanElapsedMs) || 0);
      }, 0)
    };
    result.diagnostic = aceVisibleCountDiagnostic(result, options.context || {});
    return result;
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
      ...aceCurrentManuscriptSurface(documentId),
      interactive: Boolean(interactive)
    });
  }

  async function aceVisibleStartSnapshot(documentId, reason = "visible-start-fallback") {
    const stableVisible = await aceStableVisibleGoogleDocWordCount();
    const visibleWordCount = Number(stableVisible.count);
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
      wordCountTokenizerVersion: "",
      visibleCountDiagnostic: stableVisible,
      wordCountDiagnostic: `W-START-VISIBLE-FALLBACK: started from visible Google Docs count ${Math.max(0, visibleWordCount)} because ${reason}. ${aceShortDiagnostic(stableVisible.diagnostic, 220)}`
    };
  }

  async function aceStoreVisibleStartSnapshot(documentId, extensionSessionId, visibleSnapshot) {
    await aceStorageSet({
      [aceSnapshotStorageKey(extensionSessionId)]: {
        documentId,
        revisionId: "",
        wordCount: visibleSnapshot.wordCount,
        wordCountTokenizerVersion: "",
        createdAt: new Date().toISOString(),
        source: "visible-total-baseline"
      }
    });
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
          : `E-START-COUNT-UNAVAILABLE: ${startSnapshot.error || "Google Docs API before snapshot was unavailable."}`
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
    if (!documentId || !extensionSessionId || !Number.isFinite(wordCount)) {
      return null;
    }

    await aceStorageSet({
      [aceSnapshotStorageKey(extensionSessionId)]: {
        documentId,
        revisionId: baseline.revisionId || "",
        wordCount,
        wordCountTokenizerVersion: baseline?.endDocumentWordCountTokenizerVersion || baseline?.wordCountTokenizerVersion || "",
        createdAt: new Date().toISOString(),
        source: "visible-total-baseline"
      }
    });

    return {
      ok: true,
      status: 200,
      method: "visible-total-baseline",
      revisionId: baseline.revisionId || "",
      wordCount,
      wordCountTokenizerVersion: baseline?.endDocumentWordCountTokenizerVersion || baseline?.wordCountTokenizerVersion || "",
      wordCountDiagnostic: `W-START-SAVED-TOTAL-BASELINE: started from saved total ${wordCount}.`
    };
  }

  async function aceGoogleDocWordCount(documentId, interactive, surfaceOverride = null) {
    if (!documentId) {
      return {
        ok: false,
        wordCount: null,
        error: "Google Docs document ID is missing."
      };
    }
    const surface = surfaceOverride
      ? aceSurfaceFromParts({ documentId, ...surfaceOverride })
      : aceCurrentManuscriptSurface(documentId);

    return aceGoogleDocMessage(ACE_GOOGLE_DOC_WORD_COUNT_MESSAGE, {
      documentId,
      ...surface,
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

  function aceGoogleDocNetDiagnostic({
    code = "D-NET-WORD-COUNT",
    attempt = 0,
    startWordCount = null,
    apiEndWordCount = null,
    visibleEndWordCount = null,
    netWordsChanged = 0,
    revisionChanged = false,
    startRevisionId = "",
    endRevisionId = "",
    startSource = "unknown",
    endSource = "unknown",
    details = ""
  } = {}) {
    const visibleCopy = visibleEndWordCount !== null
      && visibleEndWordCount !== undefined
      && visibleEndWordCount !== ""
      && Number.isFinite(Number(visibleEndWordCount))
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

    const detailsCopy = details ? ` ${aceShortDiagnostic(details, 260)}` : "";
    return `${code}: start ${aceDiagnosticWordCountCopy(startWordCount)}; API end ${aceDiagnosticWordCountCopy(apiEndWordCount)}; visible end ${visibleCopy}; net ${aceFormatSignedNumber(netWordsChanged)}; start source ${startSource}; end source ${endSource}; revision ${revisionCopy}; attempt ${attemptCopy}.${detailsCopy}`;
  }

  function aceVisibleWordCountFallbackResponse(response, {
    attempt = 0,
    startWordCount = null,
    startRevisionId = "",
    revisionChanged = false,
    visibleCountResult = null
  } = {}) {
    const visibleWordCount = Number(response?.visibleWordCount);
    const apiWordCount = Number(response?.apiWordCount ?? response?.wordCount);
    const netWordsChanged = Number.isFinite(startWordCount) && Number.isFinite(visibleWordCount)
      ? aceCalculateNetWordDelta(startWordCount, visibleWordCount)
      : 0;

    return {
      ...response,
      ok: true,
      wordCount: Math.max(0, visibleWordCount),
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged,
      wordCounts: null,
      endWordCounts: null,
      wordCountDiagnostic: aceGoogleDocNetDiagnostic({
        code: "W-VISIBLE-FALLBACK",
        attempt,
        startWordCount,
        apiEndWordCount: apiWordCount,
        visibleEndWordCount: visibleWordCount,
        netWordsChanged,
        revisionChanged,
        startRevisionId,
        endRevisionId: response?.revisionId || "",
        startSource: "stored-start-count",
        endSource: "visible-total-fallback",
        details: visibleCountResult?.diagnostic || response?.visibleCountDiagnostic || ""
      }),
      visibleCountDiagnostic: visibleCountResult || response?.visibleCountDiagnostic || null,
      fallbackAttempt: attempt,
      fallbackRevisionChanged: revisionChanged
    };
  }

  async function aceStabilizeVisibleFallbackResponse(response, {
    attempt = 0,
    startWordCount = null,
    startRevisionId = "",
    revisionChanged = false,
    durationMinutes = null
  } = {}) {
    const apiWordCount = Number(response?.apiWordCount ?? response?.wordCount);
    const stableVisible = await aceStableVisibleGoogleDocWordCount({
      context: {
        apiWordCount,
        startWordCount,
        startSource: "stored-start-count",
        endSource: "visible-total-fallback",
        durationMinutes
      }
    });
    const visibleWordCount = Number(stableVisible.count);
    const sourceResponse = Number.isFinite(visibleWordCount)
      ? { ...response, visibleWordCount }
      : response;
    return aceVisibleWordCountFallbackResponse(sourceResponse, {
      attempt,
      startWordCount,
      startRevisionId,
      revisionChanged,
      visibleCountResult: stableVisible
    });
  }

  async function aceGoogleDocNetAfterSave(
    documentId,
    extensionSessionId,
    startWordCount,
    startRevisionId,
    hadDocumentActivity,
    options = {}
  ) {
    const timing = options.timing || aceCreateWordCountTiming("session-end", options.trigger || "end-session");
    const settleStartedAt = aceNowMs();
    const settleDelayMs = Number.isFinite(Number(options.settleDelayMs))
      ? Math.max(0, Number(options.settleDelayMs))
      : ACE_GOOGLE_DOC_SETTLE_DELAY_MS;
    if (settleDelayMs > 0) {
      await aceDelay(settleDelayMs);
    }
    timing.settleDelayMs += aceTimingElapsedMs(settleStartedAt);

    const baselineWordCount = Math.max(0, Number(startWordCount) || 0);
    const apiCall = options.apiCall || function () {
      return aceGoogleDocMessage(ACE_GOOGLE_DOC_NET_COUNT_MESSAGE, {
        documentId,
        extensionSessionId,
        ...aceSessionManuscriptSurface({ documentId, tabId: options.tabId, tabTitle: options.tabTitle, manuscriptSurfaceId: options.manuscriptSurfaceId }),
        interactive: false,
        clearSnapshot: false
      });
    };
    const apiPromise = aceStartTimedApiWordCount(apiCall, timing);
    let apiResponse = null;
    apiPromise.then(function (response) {
      apiResponse = response;
      return response;
    });

    const stableVisible = await aceStableVisibleGoogleDocWordCount({
      timeoutMs: Number.isFinite(Number(options.visibleTimeoutMs)) ? Number(options.visibleTimeoutMs) : ACE_VISIBLE_STABLE_TIMEOUT_MS,
      delayMs: Number.isFinite(Number(options.visibleDelayMs)) ? Number(options.visibleDelayMs) : ACE_VISIBLE_STABLE_DELAY_MS,
      minStableReads: 2,
      ignoreZero: Boolean(options.ignoreZero),
      readVisibleCount: options.readVisibleCount,
      context: {
        startWordCount: baselineWordCount,
        startSource: "stored-start-count",
        endSource: "stable-visible"
      }
    });
    aceApplyVisibleTiming(timing, stableVisible);

    const visibleWordCount = Number(stableVisible.count);
    const hasStableVisible = Boolean(stableVisible.stable && Number.isFinite(visibleWordCount));
    const apiTimeoutMs = Number.isFinite(Number(options.apiTimeoutMs))
      ? Math.max(0, Number(options.apiTimeoutMs))
      : ACE_GOOGLE_DOC_API_TIMEOUT_MS;
    const compareStartedAt = aceNowMs();
    let selectedResponse = null;

    if (hasStableVisible) {
      const selectedEndWordCount = Math.max(0, visibleWordCount);
      const netWordsChanged = aceCalculateNetWordDelta(baselineWordCount, selectedEndWordCount);
      const immediateApiWordCount = Number(apiResponse?.apiWordCount ?? apiResponse?.wordCount);
      const immediateApiAvailable = Boolean(apiResponse?.ok && Number.isFinite(immediateApiWordCount));
      const immediateApiMismatch = immediateApiAvailable
        && Math.max(0, immediateApiWordCount) !== selectedEndWordCount;
      const wordCountDiagnostic = aceGoogleDocNetDiagnostic({
        code: immediateApiMismatch ? "W-API-VISIBLE-MISMATCH" : "D-STABLE-VISIBLE-NET",
        attempt: 1,
        startWordCount: baselineWordCount,
        apiEndWordCount: immediateApiAvailable ? Math.max(0, immediateApiWordCount) : null,
        visibleEndWordCount: selectedEndWordCount,
        netWordsChanged,
        revisionChanged: false,
        startRevisionId,
        endRevisionId: apiResponse?.revisionId || "",
        startSource: "stored-start-count",
        endSource: "stable-visible",
        details: immediateApiMismatch
          ? "stable active-tab visible count selected; Google Docs API may be document-wide or stale"
          : stableVisible.diagnostic
      });

      if (!apiResponse) {
        apiPromise.then(function (lateApiResponse) {
          const lateApiWordCount = Number(lateApiResponse?.apiWordCount ?? lateApiResponse?.wordCount);
          if (
            lateApiResponse?.ok
            && Number.isFinite(lateApiWordCount)
            && Math.max(0, lateApiWordCount) !== selectedEndWordCount
          ) {
            console.info("[ACE] WORD COUNT API VERIFY", {
              code: "W-API-VISIBLE-MISMATCH",
              documentId,
              startWordCount: baselineWordCount,
              apiEndWordCount: Math.max(0, lateApiWordCount),
              visibleEndWordCount: selectedEndWordCount,
              netWordsChanged
            });
          }
        });
      }

      selectedResponse = {
        ...(apiResponse || {}),
        ok: true,
        status: apiResponse?.status || 200,
        method: "stable-visible-count",
        wordCount: selectedEndWordCount,
        apiWordCount: immediateApiAvailable ? Math.max(0, immediateApiWordCount) : null,
        visibleWordCount: selectedEndWordCount,
        visibleCountDiagnostic: stableVisible,
        wordsAdded: 0,
        wordsRemoved: 0,
        netWordsChanged,
        wordCounts: null,
        endWordCounts: null,
        wordCountTokenizerVersion: apiResponse?.wordCountTokenizerVersion || "",
        wordCountDiagnostic,
        error: "",
        timing
      };
      timing.compareElapsedMs += aceTimingElapsedMs(compareStartedAt);
      aceCompleteWordCountTiming(timing, {
        apiPendingAtDecision: !apiResponse,
        finalSelectedCountSource: "stable-visible",
        trustedReason: "stable active-tab visible count matched twice",
        action: "sync-session"
      });
      return selectedResponse;
    }

    const boundedApiResponse = apiResponse || await aceWithTimeout(
      apiPromise,
      apiTimeoutMs,
      aceApiTimeoutResponse(apiTimeoutMs)
    );
    const apiWordCount = Number(boundedApiResponse?.apiWordCount ?? boundedApiResponse?.wordCount);
    const apiAvailable = Boolean(boundedApiResponse?.ok && Number.isFinite(apiWordCount));
    const apiSurfaceTrusted = Boolean(options.apiSurfaceTrusted);
    if (apiAvailable && apiSurfaceTrusted) {
      const selectedEndWordCount = Math.max(0, apiWordCount);
      const netWordsChanged = aceCalculateNetWordDelta(baselineWordCount, selectedEndWordCount);
      selectedResponse = {
        ...boundedApiResponse,
        ok: true,
        wordCount: selectedEndWordCount,
        apiWordCount: selectedEndWordCount,
        visibleWordCount: null,
        visibleCountDiagnostic: stableVisible,
        wordsAdded: 0,
        wordsRemoved: 0,
        netWordsChanged,
        wordCountDiagnostic: aceGoogleDocNetDiagnostic({
          code: "D-API-FALLBACK-NET",
          attempt: 1,
          startWordCount: baselineWordCount,
          apiEndWordCount: selectedEndWordCount,
          visibleEndWordCount: null,
          netWordsChanged,
          revisionChanged: false,
          startRevisionId,
          endRevisionId: boundedApiResponse.revisionId || "",
          startSource: "stored-start-count",
          endSource: "google-docs-api",
          details: "visible count unavailable; API accepted because surface was explicitly trusted"
        }),
        timing
      };
      timing.compareElapsedMs += aceTimingElapsedMs(compareStartedAt);
      aceCompleteWordCountTiming(timing, {
        apiPendingAtDecision: false,
        finalSelectedCountSource: "google-docs-api",
        trustedReason: "visible unavailable; API surface explicitly trusted",
        action: "sync-session"
      });
      return selectedResponse;
    }

    timing.compareElapsedMs += aceTimingElapsedMs(compareStartedAt);
    selectedResponse = {
      ...boundedApiResponse,
      ok: false,
      wordCount: null,
      apiWordCount: apiAvailable ? Math.max(0, apiWordCount) : null,
      visibleWordCount: null,
      visibleCountDiagnostic: stableVisible,
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged: 0,
      wordCountDiagnostic: aceGoogleDocNetDiagnostic({
        code: "E-NO-TRUSTED-END-COUNT",
        attempt: 1,
        startWordCount: baselineWordCount,
        apiEndWordCount: apiAvailable ? Math.max(0, apiWordCount) : null,
        visibleEndWordCount: null,
        netWordsChanged: 0,
        revisionChanged: false,
        startRevisionId,
        endRevisionId: boundedApiResponse?.revisionId || "",
        startSource: "stored-start-count",
        endSource: apiAvailable ? "google-docs-api-untrusted" : "none",
        details: stableVisible.diagnostic || boundedApiResponse?.error || "stable visible count unavailable"
      }),
      error: "E-NO-TRUSTED-END-COUNT: Stable active-tab visible word count was unavailable.",
      timing
    };
    aceCompleteWordCountTiming(timing, {
      apiPendingAtDecision: false,
      finalSelectedCountSource: apiAvailable ? "google-docs-api-untrusted" : "none",
      trustedReason: apiAvailable
        ? "API is document-wide unless surface is explicitly trusted"
        : "no stable visible count and no API count",
      action: "measurement-pending"
    });
    return selectedResponse;
  }

  function aceCatchUpVisibleCandidateSummary(stableVisible) {
    return (stableVisible?.candidates || []).slice(0, 6).map(function (candidate) {
      return {
        count: Number.isFinite(Number(candidate?.count)) ? Math.max(0, Number(candidate.count)) : null,
        source: candidate?.source || "",
        snippet: candidate?.snippet || ""
      };
    });
  }

  function aceCatchUpBaselineSource(baseline) {
    if (!baseline) {
      return "";
    }
    if (baseline.wordCountMethod) {
      return baseline.wordCountMethod;
    }
    return baseline.endDocumentWordCountTokenizerVersion
      ? "google-docs-api"
      : "saved-total-baseline";
  }

  function aceFiniteNumberOrNull(value) {
    return value !== null && value !== undefined && Number.isFinite(Number(value))
      ? Number(value)
      : null;
  }

  function aceCatchUpTriggerCanPrompt(trigger) {
    return ACE_CATCH_UP_BOUNDARY_TRIGGERS.has(String(trigger || ""));
  }

  function aceIsTypingSuppressionActive(now = Date.now()) {
    return Boolean(
      aceLastTypingTimestamp
      && Math.max(0, Number(now) - Number(aceLastTypingTimestamp)) < ACE_TYPING_CATCH_UP_SUPPRESSION_MS
    );
  }

  function aceCatchUpSuppressionReason(trigger) {
    if (aceCatchUpTriggerCanPrompt(trigger)) {
      return "";
    }
    if (aceIsTypingSuppressionActive()) {
      return "typing-suppressed";
    }
    return "trigger-not-boundary";
  }

  function aceCatchUpSnapshotDiagnostic({
    code,
    baselineWordCount,
    apiWordCount = null,
    visibleWordCount = null,
    netWordsChanged = 0,
    apiResponse = null,
    stableVisible = null,
    endSource = "unknown",
    details = ""
  }) {
    return aceGoogleDocNetDiagnostic({
      code,
      attempt: 1,
      startWordCount: baselineWordCount,
      apiEndWordCount: apiWordCount,
      visibleEndWordCount: visibleWordCount,
      netWordsChanged,
      revisionChanged: false,
      startRevisionId: "",
      endRevisionId: apiResponse?.revisionId || "",
      startSource: "saved-total-baseline",
      endSource,
      details: details || stableVisible?.diagnostic || ""
    });
  }

  async function aceGoogleDocWordCountAfterSettle(documentId, baseline, options = {}) {
    const timing = options.timing || aceCreateWordCountTiming("catch-up-check", options.trigger || "catch-up");
    const settleStartedAt = aceNowMs();
    const settleDelayMs = Number.isFinite(Number(options.settleDelayMs))
      ? Math.max(0, Number(options.settleDelayMs))
      : ACE_GOOGLE_DOC_SETTLE_DELAY_MS;
    if (settleDelayMs > 0) {
      await aceDelay(settleDelayMs);
    }
    timing.settleDelayMs += aceTimingElapsedMs(settleStartedAt);

    const baselineWordCount = Math.max(0, Number(baseline?.endDocumentWordCount) || 0);
    const apiCall = options.apiCall || function () {
      return aceGoogleDocWordCount(documentId, true);
    };
    const apiPromise = options.apiResponse
      ? Promise.resolve(options.apiResponse)
      : aceStartTimedApiWordCount(apiCall, timing);
    let apiResponse = options.apiResponse || null;
    apiPromise.then(function (response) {
      apiResponse = response;
      return response;
    });

    const stableVisible = await aceStableVisibleGoogleDocWordCount({
      timeoutMs: Number.isFinite(Number(options.visibleTimeoutMs)) ? Number(options.visibleTimeoutMs) : ACE_REFRESH_VISIBLE_STABLE_TIMEOUT_MS,
      delayMs: Number.isFinite(Number(options.visibleDelayMs)) ? Number(options.visibleDelayMs) : ACE_VISIBLE_STABLE_DELAY_MS,
      minStableReads: 2,
      ignoreZero: Boolean(options.ignoreZero),
      readVisibleCount: options.readVisibleCount,
      context: {
        startWordCount: baselineWordCount,
        startSource: "saved-total-baseline",
        endSource: "stable-visible"
      }
    });
    aceApplyVisibleTiming(timing, stableVisible);

    const visibleWordCount = Number(stableVisible.count);
    const hasStableVisible = Boolean(stableVisible.stable && Number.isFinite(visibleWordCount));
    const apiTimeoutMs = Number.isFinite(Number(options.apiTimeoutMs))
      ? Math.max(0, Number(options.apiTimeoutMs))
      : ACE_GOOGLE_DOC_API_TIMEOUT_MS;
    const compareStartedAt = aceNowMs();
    const boundedApiResponse = hasStableVisible
      ? apiResponse
      : apiResponse || await aceWithTimeout(apiPromise, apiTimeoutMs, aceApiTimeoutResponse(apiTimeoutMs));
    const apiWordCount = Number(boundedApiResponse?.wordCount ?? boundedApiResponse?.apiWordCount);
    const apiAvailable = Boolean(boundedApiResponse?.ok && Number.isFinite(apiWordCount));
    const visibleApiMismatch = hasStableVisible
      && apiAvailable
      && Math.max(0, visibleWordCount) !== Math.max(0, apiWordCount);
    const selectedWordCount = hasStableVisible ? Math.max(0, visibleWordCount) : null;
    const netWordsChanged = Number.isFinite(selectedWordCount)
      ? aceCalculateNetWordDelta(baselineWordCount, selectedWordCount)
      : 0;

    if (!hasStableVisible) {
      const apiFallbackAllowed = Boolean(options.apiSurfaceTrusted);
      if (apiAvailable && apiFallbackAllowed) {
        const selectedApiCount = Math.max(0, apiWordCount);
        const apiNetWordsChanged = aceCalculateNetWordDelta(baselineWordCount, selectedApiCount);
        timing.compareElapsedMs += aceTimingElapsedMs(compareStartedAt);
        aceCompleteWordCountTiming(timing, {
          apiPendingAtDecision: false,
          finalSelectedCountSource: "google-docs-api",
          trustedReason: "visible unavailable; API surface explicitly trusted",
          action: apiNetWordsChanged === 0 ? "no-catch-up" : "show-catch-up"
        });
        return {
          ...boundedApiResponse,
          ok: true,
          status: boundedApiResponse.status || 200,
          method: "google-docs-api",
          wordCount: selectedApiCount,
          apiWordCount: selectedApiCount,
          visibleWordCount: null,
          visibleCountDiagnostic: stableVisible,
          visibleCandidates: aceCatchUpVisibleCandidateSummary(stableVisible),
          netWordsChanged: apiNetWordsChanged,
          currentCountSource: "google-docs-api",
          currentCountTrusted: true,
          wordCountDiagnostic: aceCatchUpSnapshotDiagnostic({
            code: "D-CATCHUP-API-FALLBACK",
            baselineWordCount,
            apiWordCount: selectedApiCount,
            visibleWordCount: null,
            netWordsChanged: apiNetWordsChanged,
            apiResponse: boundedApiResponse,
            stableVisible,
            endSource: "google-docs-api",
            details: "visible count unavailable; API accepted because surface was explicitly trusted"
          }),
          error: "",
          timing
        };
      }

      timing.compareElapsedMs += aceTimingElapsedMs(compareStartedAt);
      aceCompleteWordCountTiming(timing, {
        apiPendingAtDecision: false,
        finalSelectedCountSource: apiAvailable ? "google-docs-api-untrusted" : "none",
        trustedReason: apiAvailable
          ? "API is document-wide unless surface is explicitly trusted"
          : "stable visible count unavailable",
        action: "diagnostic-only"
      });
      return {
        ...boundedApiResponse,
        ok: false,
        skipCatchUp: true,
        wordCount: null,
        apiWordCount: apiAvailable ? Math.max(0, apiWordCount) : null,
        visibleWordCount: null,
        visibleCountDiagnostic: stableVisible,
        visibleCandidates: aceCatchUpVisibleCandidateSummary(stableVisible),
        netWordsChanged: 0,
        currentCountSource: "none",
        currentCountTrusted: false,
        wordCountDiagnostic: aceCatchUpSnapshotDiagnostic({
          code: "D-CATCHUP-NO-STABLE-VISIBLE",
          baselineWordCount,
          apiWordCount: apiAvailable ? Math.max(0, apiWordCount) : null,
          visibleWordCount: null,
          stableVisible,
          endSource: apiAvailable ? "google-docs-api" : "none",
          details: "stable visible count unavailable; diagnostic only"
        }),
        error: "",
        timing
      };
    }

    if (!apiResponse) {
      apiPromise.then(function (lateApiResponse) {
        const lateApiWordCount = Number(lateApiResponse?.wordCount ?? lateApiResponse?.apiWordCount);
        if (
          lateApiResponse?.ok
          && Number.isFinite(lateApiWordCount)
          && Math.max(0, lateApiWordCount) !== selectedWordCount
        ) {
          console.info("[ACE] CATCH-UP API VERIFY", {
            code: "W-API-VISIBLE-MISMATCH",
            documentId,
            baselineWordCount,
            apiEndWordCount: Math.max(0, lateApiWordCount),
            visibleEndWordCount: selectedWordCount,
            netWordsChanged
          });
        }
      });
    }

    timing.compareElapsedMs += aceTimingElapsedMs(compareStartedAt);
    aceCompleteWordCountTiming(timing, {
      apiPendingAtDecision: !apiResponse,
      finalSelectedCountSource: "stable-visible",
      trustedReason: "stable active-tab visible count matched twice",
      action: netWordsChanged === 0 ? "no-catch-up" : "show-catch-up"
    });
    return {
      ...(apiResponse || {}),
      ok: true,
      status: apiResponse?.status || 200,
      method: "stable-visible-count",
      wordCount: selectedWordCount,
      apiWordCount: apiAvailable ? Math.max(0, apiWordCount) : null,
      visibleWordCount: selectedWordCount,
      visibleCountDiagnostic: stableVisible,
      visibleCandidates: aceCatchUpVisibleCandidateSummary(stableVisible),
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged,
      currentCountSource: "stable-visible",
      currentCountTrusted: true,
      apiVisibleMismatch: visibleApiMismatch,
      wordCountDiagnostic: aceCatchUpSnapshotDiagnostic({
        code: visibleApiMismatch
          ? "D-CATCHUP-STABLE-VISIBLE-API-MISMATCH"
          : apiAvailable ? "D-CATCHUP-STABLE-VISIBLE" : "D-CATCHUP-VISIBLE-ONLY",
        baselineWordCount,
        apiWordCount: apiAvailable ? Math.max(0, apiWordCount) : null,
        visibleWordCount: selectedWordCount,
        netWordsChanged,
        apiResponse: apiResponse || null,
        stableVisible,
        endSource: "stable-visible",
        details: visibleApiMismatch
          ? "stable active-tab visible count selected; API may be document-wide or stale"
          : ""
      }),
      error: "",
      timing
    };
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

    aceActiveSession = {
      ...aceActiveSession,
      lastPersistedAt: new Date().toISOString()
    };
    await aceStorageSet({ [ACE_LOCAL_STORAGE.activeSession]: aceActiveSession });
  }

  function aceAbandonedSessionRecord(session, reason = "abandoned") {
    const sessionScope = aceSessionScope(session);
    return {
      ...aceSessionManuscriptSurface(session),
      sessionScope,
      chromeTabId: session?.chromeTabId || sessionScope.chromeTabId || "",
      pageInstanceId: session?.pageInstanceId || sessionScope.pageInstanceId || "",
      projectId: session?.projectId || sessionScope.projectId || "",
      project: session?.project || null,
      documentUrl: session?.documentUrl || aceDocumentUrl(),
      extensionSessionId: session?.extensionSessionId || "",
      sessionType: aceNormalizeSessionType(session?.sessionType),
      startedAt: session?.startedAt || new Date().toISOString(),
      abandonedAt: session?.abandonedAt || new Date().toISOString(),
      lastPersistedAt: session?.lastPersistedAt || new Date().toISOString(),
      abandonReason: reason,
      status: session?.status || "abandoned",
      wordsWritten: 0,
      wordsEdited: 0,
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged: 0,
      startDocumentWordCount: Number.isFinite(Number(session?.startDocumentWordCount))
        ? Number(session.startDocumentWordCount)
        : null,
      startDocumentRevisionId: session?.startDocumentRevisionId || "",
      startDocumentWordCountTokenizerVersion: session?.startDocumentWordCountTokenizerVersion || "",
      wordCountMethod: session?.wordCountMethod || "google-docs-api",
      wordCountError: session?.wordCountError || "",
      wordCountDiagnostic: session?.wordCountDiagnostic || "",
      hadDocumentActivity: Boolean(session?.hadDocumentActivity)
    };
  }

  async function aceAbandonedSessions() {
    const stored = await aceStorageGet(ACE_LOCAL_STORAGE.abandonedSessions);
    return Array.isArray(stored[ACE_LOCAL_STORAGE.abandonedSessions])
      ? stored[ACE_LOCAL_STORAGE.abandonedSessions]
      : [];
  }

  async function aceStoreAbandonedSession(session, reason = "abandoned") {
    const record = aceAbandonedSessionRecord(session, reason);
    if (!record.extensionSessionId || !record.manuscriptSurfaceId) {
      return null;
    }
    const abandoned = await aceAbandonedSessions();
    const withoutDuplicate = abandoned.filter(function (item) {
      return item.extensionSessionId !== record.extensionSessionId;
    });
    await aceStorageSet({
      [ACE_LOCAL_STORAGE.abandonedSessions]: [...withoutDuplicate, record]
    });
    return record;
  }

  async function aceRemoveAbandonedSession(extensionSessionId) {
    const abandoned = await aceAbandonedSessions();
    await aceStorageSet({
      [ACE_LOCAL_STORAGE.abandonedSessions]: abandoned.filter(function (item) {
        return item.extensionSessionId !== extensionSessionId;
      })
    });
  }

  async function aceRemoveStoredActiveSessionIfMatches(extensionSessionId) {
    const stored = await aceStorageGet(ACE_LOCAL_STORAGE.activeSession);
    const activeSession = stored[ACE_LOCAL_STORAGE.activeSession];
    if (activeSession?.extensionSessionId === extensionSessionId) {
      await aceStorageRemove(ACE_LOCAL_STORAGE.activeSession);
    }
  }

  async function aceAbandonedSessionForCurrentSurface(surface = aceCurrentManuscriptSurface()) {
    const abandoned = await aceAbandonedSessions();
    const pending = await acePendingSessions();
    return abandoned.find(function (session) {
      if (!aceRecordMatchesSurface(session, surface)) {
        return false;
      }
      return !pending.some(function (pendingSession) {
        return pendingSession.extensionSessionId === session.extensionSessionId;
      });
    }) || null;
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
    const surface = aceCurrentSurface || aceCurrentManuscriptSurface();
    const binding = aceCurrentBinding;
    const project = binding?.project || null;
    const projectTitle = project?.bookTitle || "Project";
    const tabLabel = surface?.manuscriptSurfaceLabel || surface?.tabTitle || "Current manuscript";
    const errorCopy = acePromptError
      ? `<div class="ace-sync-copy ace-sync-copy--pending">${aceEscapeHtml(aceShortDiagnostic(acePromptError))}</div>`
      : "";
    const body = binding?.projectId
      ? `
        <div class="ace-field-readout"><span>Project</span><strong>${aceEscapeHtml(projectTitle)}</strong></div>
        <div class="ace-field-readout"><span>Current tab</span><strong>${aceEscapeHtml(tabLabel)}</strong></div>
        ${errorCopy}
        <div class="ace-actions">
          <button class="ace-button ace-button--primary" type="button" data-ace-action="start-writing">Start writing</button>
          <button class="ace-button" type="button" data-ace-action="start-editing">Start editing</button>
          <button class="ace-button" type="button" data-ace-action="manual-sync">Sync document changes</button>
          <button class="ace-button" type="button" data-ace-action="show-unbind">Unbind</button>
          <button class="ace-button" type="button" data-ace-action="open">Open project</button>
        </div>
      `
      : `
        <div class="ace-field-readout"><span>Project</span><strong>Not bound</strong></div>
        <div class="ace-field-readout"><span>Current tab</span><strong>${aceEscapeHtml(tabLabel)}</strong></div>
        ${errorCopy}
        <div class="ace-actions">
          <button class="ace-button ace-button--primary" type="button" data-ace-action="bind-project">Bind project</button>
          <button class="ace-button" type="button" data-ace-action="create-project">Create project</button>
        </div>
      `;
    aceWidget.className = "ace-widget ace-widget--prompt";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      ${acePanelHeaderHtml("Scriptor", "Google Docs")}
      ${body}
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderUnbindConfirm() {
    const project = aceCurrentBinding?.project || {};
    aceWidget.className = "ace-widget ace-widget--prompt";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      ${acePanelHeaderHtml("Unbind project?", "Scriptor")}
      <div class="ace-prompt-copy">${aceEscapeHtml(project.bookTitle || "Project")}</div>
      <div class="ace-project-copy">Future sessions only.<br>History stays saved.</div>
      <div class="ace-actions">
        <button class="ace-button" type="button" data-ace-action="cancel-unbind">Cancel</button>
        <button class="ace-button ace-button--end" type="button" data-ace-action="confirm-unbind">Unbind</button>
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
      ? "Enter the current Google Docs word count to log catch-up."
      : `${activity.activityCopy} since last session.`;
    const changeCopy = needsWordCountConfirmation
      ? "Detected change: <span data-ace-catch-up-preview>enter a count to calculate changes</span>."
      : `Detected change: <span data-ace-catch-up-preview>${activity.deltaCopy}</span>.`;
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
      <div class="ace-project-copy">From ${aceFormatNumber(startWordCount)} &rarr; ${Number.isFinite(endWordCount) ? aceFormatNumber(endWordCount) : "current"} words. ${changeCopy}</div>
      <label class="ace-field ace-field--compact">
        <span>Current word count</span>
        <input type="text" inputmode="numeric" autocomplete="off" data-ace-catch-up-word-count value="${Number.isFinite(endWordCount) ? aceEscapeHtml(aceFormatNumber(endWordCount)) : ""}">
      </label>
      <div class="ace-project-copy">This creates a 1 min ${sessionTypeCopy} session.</div>
      ${diagnosticCopy}
      ${statusCopy}
      <div class="ace-actions">
        <button class="ace-button ace-button--primary" type="button" data-ace-action="add-catch-up">Log catch-up</button>
        <button class="ace-button" type="button" data-ace-action="skip-catch-up">Skip</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderRecoveryModal() {
    const abandonedSession = aceRecoveryCandidate?.abandonedSession || {};
    const projectName = abandonedSession.project?.bookTitle
      || aceSelectedProject?.bookTitle
      || "this project";
    const durationMinutes = aceDurationMinutes(aceRecoveryCandidate?.elapsedMs || aceElapsedMsForSession(abandonedSession));
    const netWordsChanged = Number(aceRecoveryCandidate?.netWordsChanged);
    const wordDeltaCopy = Number.isFinite(netWordsChanged)
      ? `${aceFormatSignedNumber(netWordsChanged)} ${Math.abs(netWordsChanged) === 1 ? "word" : "words"}`
      : "pending";
    const statusCopy = aceSyncStatus
      ? `<div class="ace-sync-copy ace-sync-copy--${aceSyncStatus}">${aceEscapeHtml(aceSyncMessage)}</div>`
      : "";

    aceWidget.className = "ace-widget ace-widget--recovery";
    aceWidget.innerHTML = `
      ${acePanelHeaderHtml("Session recovery", "Scriptor")}
      <div class="ace-prompt-copy">You closed your last session without saving.</div>
      <div class="ace-field-readout"><span>Let's do Project</span><strong>${aceEscapeHtml(projectName)}</strong></div>
      <div class="ace-field-readout"><span>Duration</span><strong>${aceEscapeHtml(String(durationMinutes))} min</strong></div>
      <div class="ace-field-readout"><span>Word change</span><strong>${aceEscapeHtml(wordDeltaCopy)}</strong></div>
      ${statusCopy}
      <div class="ace-actions">
        <button class="ace-button ace-button--primary" type="button" data-ace-action="recover-session">Recover Session</button>
        <button class="ace-button" type="button" data-ace-action="discard-recovery">Discard</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderTabChangedBlock(currentSurface = aceCurrentManuscriptSurface()) {
    const startedSurface = aceSessionManuscriptSurface(aceActiveSession);
    const originalTab = startedSurface.tabTitle || startedSurface.manuscriptSurfaceLabel || "original tab";
    const currentTab = currentSurface?.tabTitle || currentSurface?.manuscriptSurfaceLabel || "current tab";
    const unknownCopy = aceSurfaceConfidence(currentSurface)
      ? ""
      : '<div class="ace-sync-copy ace-sync-copy--pending">Couldn’t identify this tab. Refresh Google Docs.</div>';
    aceWidget.className = "ace-widget ace-widget--prompt";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      ${acePanelHeaderHtml("Tab changed", "Session paused")}
      <div class="ace-field-readout"><span>Project</span><strong>${aceEscapeHtml(aceActiveSession?.project?.bookTitle || aceSelectedProject?.bookTitle || "Project")}</strong></div>
      <div class="ace-field-readout"><span>Current tab</span><strong>${aceEscapeHtml(currentTab)}</strong></div>
      <div class="ace-prompt-copy">Return to "${aceEscapeHtml(originalTab)}" or end this session.</div>
      ${unknownCopy}
      <div class="ace-actions">
        <button class="ace-button ace-button--primary" type="button" data-ace-action="go-back-tab">Go back</button>
        <button class="ace-button ace-button--end" type="button" data-ace-action="end">End session</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceCatchUpActivity(candidate) {
    const netWordsChanged = Number(candidate?.netWordsChanged) || 0;
    const totalWords = Math.abs(netWordsChanged);
    const deltaCopy = `${aceFormatSignedNumber(netWordsChanged)} ${totalWords === 1 ? "word" : "words"}`;
    const activityCopy = `Net: ${deltaCopy}`;

    return {
      wordsAdded: 0,
      wordsRemoved: 0,
      totalWords,
      isEditingCatchUp: candidate?.sessionType === "editing",
      activityCopy,
      deltaCopy,
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
    const netWordsChanged = aceCalculateNetWordDelta(startWordCount, currentWordCount);
    const sessionType = netWordsChanged < 0 ? "editing" : "writing";

    return {
      ...candidate,
      currentSnapshot: {
        ...(candidate?.currentSnapshot || {}),
        wordCount: currentWordCount,
        wordsAdded: 0,
        wordsRemoved: 0,
        netWordsChanged
      },
      endDocumentWordCount: currentWordCount,
      endDocumentWordCounts: null,
      wordsWritten: Math.max(0, netWordsChanged),
      wordsAdded: 0,
      wordsRemoved: 0,
      wordsEdited: 0,
      netWordsChanged,
      sessionType,
      needsWordCountConfirmation: false
    };
  }

  function aceBuildCatchUpSession(catchUpCandidate, projectId, now = Date.now()) {
    const endedAt = new Date(now).toISOString();
    const startedAt = new Date(now - 60000).toISOString();
    const sessionType = catchUpCandidate.sessionType === "editing" ? "editing" : "writing";
    const netWordsChanged = Number(catchUpCandidate.netWordsChanged) || 0;
    const wordsWritten = sessionType === "writing"
      ? Math.max(0, Number(catchUpCandidate.wordsWritten) || netWordsChanged)
      : 0;

    return {
      ...aceSessionManuscriptSurface(catchUpCandidate),
      projectId,
      sessionType,
      startedAt,
      endedAt,
      durationMinutes: 1,
      source: "catch-up",
      documentUrl: catchUpCandidate.documentUrl || aceDocumentUrl(),
      notes: sessionType === "editing"
        ? "Catch-up: net words changed outside a tracked session."
        : "Catch-up: net word progress outside a tracked session.",
      extensionSessionId: aceCreateExtensionSessionId(catchUpCandidate.documentId),
      wordsWritten,
      wordsEdited: 0,
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged,
      startDocumentWordCount: catchUpCandidate.startDocumentWordCount,
      startDocumentRevisionId: catchUpCandidate.baseline?.revisionId || "",
      endDocumentWordCount: catchUpCandidate.endDocumentWordCount,
      endDocumentWordCounts: null,
      endDocumentRevisionId: catchUpCandidate.currentSnapshot?.revisionId || "",
      wordCountTokenizerVersion: catchUpCandidate.currentSnapshot?.wordCountTokenizerVersion || "",
      wordCountMethod: catchUpCandidate.currentSnapshot?.currentCountSource || "catch-up",
      wordCountError: "",
      hadDocumentActivity: true,
      measurementPending: false
    };
  }

  async function aceSaveSkippedCatchUpBaseline(catchUpCandidate) {
    if (!catchUpCandidate || catchUpCandidate.endDocumentWordCount === null || catchUpCandidate.endDocumentWordCount === undefined) {
      return;
    }

    await aceSaveDocumentBaseline({
      ...aceSessionManuscriptSurface(catchUpCandidate),
      projectId: catchUpCandidate.baseline?.projectId || "",
      extensionSessionId: "",
      endDocumentWordCount: catchUpCandidate.endDocumentWordCount,
      endDocumentRevisionId: catchUpCandidate.currentSnapshot?.revisionId || "",
      wordCountTokenizerVersion: catchUpCandidate.currentSnapshot?.wordCountTokenizerVersion || "",
      wordCountMethod: catchUpCandidate.currentSnapshot?.currentCountSource || "catch-up-skip"
    }, catchUpCandidate.baseline?.project || null);
    console.info("[ACE] D-CATCHUP-SKIPPED-BASELINE-UPDATED", {
      documentId: catchUpCandidate.documentId || "",
      tabId: catchUpCandidate.tabId || "",
      manuscriptSurfaceId: catchUpCandidate.manuscriptSurfaceId || "",
      baselineWordCount: catchUpCandidate.startDocumentWordCount,
      currentWordCount: catchUpCandidate.endDocumentWordCount,
      netWordsChanged: catchUpCandidate.netWordsChanged
    });
  }

  function aceUnsafeCatchUpMessage(candidate) {
    return "";
  }

  function aceSessionWordCount(session = aceActiveSession) {
    if (!session) {
      return 0;
    }

    return Math.abs(aceSessionNetWordsChanged(session));
  }

  function aceSessionWordsCopy(session) {
    if (!session) {
      return "";
    }

    const netWordsChanged = aceSessionNetWordsChanged(session);
    return ` · Net: ${aceFormatSignedNumber(netWordsChanged)} ${Math.abs(netWordsChanged) === 1 ? "word" : "words"}`;
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

    const currentSurface = aceCurrentManuscriptSurface();
    if (aceSessionSurfaceMismatch(aceActiveSession, currentSurface) || !aceSurfaceConfidence(currentSurface)) {
      aceState = "tab-blocked";
      aceRenderTabChangedBlock(currentSurface);
      return;
    }

    const label = aceCapitalize(aceActiveSession?.sessionType || "writing");
    const projectTitle = aceSelectedProject?.bookTitle
      || aceCurrentBinding?.project?.bookTitle
      || aceActiveSession?.project?.bookTitle
      || "";
    const projectCopy = projectTitle
      ? `<div class="ace-project-copy">Project: ${aceEscapeHtml(projectTitle)}</div>`
      : "";
    const tabCopy = `<div class="ace-project-copy">Current tab: ${aceEscapeHtml(aceActiveSession.tabTitle || aceActiveSession.manuscriptSurfaceLabel || "Current manuscript")}</div>`;

    aceWidget.className = "ace-widget ace-widget--active";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      ${acePanelHeaderHtml("Session", label)}
      <div class="ace-session-metric">
        <span class="ace-session-type">${label}</span>
        <strong class="ace-time">${aceFormatTimer(aceElapsedMs())}</strong>
        <small>Tracking from Google Docs</small>
      </div>
      ${projectCopy}
      ${tabCopy}
      <div class="ace-actions">
        <button class="ace-button ace-button--end" type="button" data-ace-action="end">End</button>
        <button class="ace-button" type="button" data-ace-action="show-issue-form">Issue</button>
        <button class="ace-button" type="button" data-ace-action="open">Open project</button>
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
    const completedSurface = aceSessionManuscriptSurface(aceCompletedSession);
    const tabCopy = `<div class="ace-project-copy">Tab: ${aceEscapeHtml(completedSurface.tabTitle || completedSurface.manuscriptSurfaceLabel || "Current manuscript")}</div>`;
    const netCopy = aceCompletedSession?.measurementPending
      ? '<div class="ace-project-copy">Net: pending</div>'
      : `<div class="ace-project-copy">Net: ${aceEscapeHtml(aceFormatSignedNumber(aceSessionNetWordsChanged(aceCompletedSession)))} words</div>`;
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
      netWordsChanged: aceCompletedSession?.netWordsChanged,
      display: wordsCopy
    });

    aceWidget.className = "ace-widget ace-widget--completed";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      ${acePanelHeaderHtml("Session saved", label)}
      <div class="ace-completed-copy">${label} session tracked: ${aceFormatCompletedMinutes(aceCompletedSession?.durationMinutes || 1)}</div>
      ${projectCopy}
      ${tabCopy}
      ${netCopy}
      ${wordCountDiagnosticCopy}
      ${diagnosticCopy}
      ${statusCopy}
      <div class="ace-actions">
        <button class="ace-button" type="button" data-ace-action="open">Open project</button>
        ${aceSyncStatus && aceSyncStatus !== "synced" ? `<button class="ace-button" type="button" data-ace-action="retry-sync" ${retryDisabled}>Retry sync</button>` : ""}
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
          : aceProjectPickerMode === "bind"
            ? "Bind project"
            : "Choose project";
    const rows = aceProjects.length
      ? aceProjects.map(function (item) {
          const project = item.project || item;
          const isBound = Boolean(item.isBound);
          const isStale = aceIsStaleBinding(item);
          const deletedBinding = aceDeletedBindingForProjectItem(item);
          const statusLabel = aceProjectPickerStatusLabel(item);
          if (isStale && deletedBinding && !isBound) {
            return `
              <button class="ace-project-option ace-project-option--stale" type="button" data-ace-project-id="${aceEscapeHtml(project.id)}">
                <span>${aceEscapeHtml(project.bookTitle)}</span>
                <small>${aceEscapeHtml(statusLabel)}</small>
              </button>
            `;
          }
          if (isStale) {
            return `
              <div class="ace-project-option ace-project-option--stale">
                <span>${aceEscapeHtml(project.bookTitle)}</span>
                <small>${aceEscapeHtml(statusLabel)}</small>
                <button class="ace-button" type="button" data-ace-action="clear-stale-binding" data-ace-project-id="${aceEscapeHtml(project.id)}">Clear</button>
              </div>
            `;
          }
          return `
            <button class="ace-project-option" type="button" ${isBound ? "disabled" : `data-ace-project-id="${aceEscapeHtml(project.id)}"`}>
              <span>${aceEscapeHtml(project.bookTitle)}</span>
              <small>${aceEscapeHtml(statusLabel)}</small>
            </button>
          `;
        }).join("")
      : '<div class="ace-empty">No active projects found.</div>';

    aceWidget.className = "ace-widget ace-widget--picker";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      ${acePanelHeaderHtml(copy, "Projects")}
      <div class="ace-project-list">${rows}</div>
      <div class="ace-actions">
        <button class="ace-button" type="button" data-ace-action="open">Open app</button>
        <button class="ace-button" type="button" data-ace-action="show-controls">Back</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderClearStaleBindingConfirm() {
    const item = acePendingClearStaleBinding;
    const project = item?.project || item || {};
    const label = aceProjectPickerStatusLabel(item || {});
    aceWidget.className = "ace-widget ace-widget--confirm";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      ${acePanelHeaderHtml("Clear binding?", project.bookTitle || "Project")}
      <div class="ace-prompt-copy">This document is missing. History stays saved.</div>
      <div class="ace-prompt-copy">${aceEscapeHtml(label)}</div>
      <div class="ace-actions">
        <button class="ace-button" type="button" data-ace-action="cancel-clear-stale-binding">Cancel</button>
        <button class="ace-button ace-button--end" type="button" data-ace-action="confirm-clear-stale-binding">Clear</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderDeletedBindingRebindConfirm() {
    aceWidget.className = "ace-widget ace-widget--confirm";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      ${acePanelHeaderHtml("Update binding?", "Scriptor")}
      <div class="ace-prompt-copy">This project was bound to a now-deleted file. Update this project to your current tab?</div>
      <div class="ace-actions">
        <button class="ace-button ace-button--primary" type="button" data-ace-action="confirm-deleted-binding-rebind">Yes</button>
        <button class="ace-button" type="button" data-ace-action="cancel-deleted-binding-rebind">No</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  const ACE_CREATE_PROJECT_STEPS = [
    { key: "title", label: "Project title", placeholder: "Example: The Hollow Orchard", type: "text" },
    { key: "manuscriptType", label: "Manuscript type", placeholder: "Novel", type: "text" },
    { key: "structureUnit", label: "Structure unit", placeholder: "Chapter", type: "text" },
    { key: "targetWordCount", label: "Target word count", placeholder: "80000", type: "number" },
    { key: "wordsWrittenSoFar", label: "Words written so far", placeholder: "0", type: "number" },
    { key: "deadline", label: "Deadline", placeholder: "", type: "date" }
  ];

  function aceDefaultCreateProjectDraft(currentWordCount = 0) {
    return {
      title: "",
      manuscriptType: "Novel",
      structureUnit: "Chapter",
      targetWordCount: 80000,
      wordsWrittenSoFar: Math.max(0, Number(currentWordCount) || 0),
      deadline: ""
    };
  }

  function aceValidateCreateProjectDraft(draft) {
    if (!String(draft?.title || "").trim()) {
      return "Title required.";
    }
    if (!String(draft?.manuscriptType || "").trim()) {
      return "Type required.";
    }
    if (!String(draft?.structureUnit || "").trim()) {
      return "Unit required.";
    }
    if (!(Number(draft?.targetWordCount) > 0)) {
      return "Target must be positive.";
    }
    if (!(Number(draft?.wordsWrittenSoFar) >= 0)) {
      return "Words must be zero or more.";
    }
    return "";
  }

  function aceReadCreateProjectStep() {
    const input = aceWidget.querySelector("[data-ace-create-project-input]");
    const step = ACE_CREATE_PROJECT_STEPS[aceCreateProjectStep];
    if (!input || !step || !aceCreateProjectDraft) {
      return;
    }
    const value = input.value;
    aceCreateProjectDraft[step.key] = step.type === "number"
      ? Math.max(0, Number(String(value).replace(/,/g, "")) || 0)
      : String(value || "").trim();
  }

  function aceRenderCreateProject() {
    const step = ACE_CREATE_PROJECT_STEPS[aceCreateProjectStep] || ACE_CREATE_PROJECT_STEPS[0];
    const value = aceCreateProjectDraft?.[step.key] ?? "";
    const dots = ACE_CREATE_PROJECT_STEPS.map(function (_item, index) {
      return `<span class="ace-dot${index === aceCreateProjectStep ? " ace-dot--active" : ""}"></span>`;
    }).join("");
    const errorCopy = aceCreateProjectError
      ? `<div class="ace-sync-copy ace-sync-copy--pending">${aceEscapeHtml(aceCreateProjectError)}</div>`
      : "";
    const isLast = aceCreateProjectStep === ACE_CREATE_PROJECT_STEPS.length - 1;
    aceWidget.className = "ace-widget ace-widget--prompt";
    aceWidget.innerHTML = `
      ${aceCloseButtonHtml()}
      ${acePanelHeaderHtml("Create project", "Scriptor")}
      <label class="ace-field">
        <span>${aceEscapeHtml(step.label)}</span>
        <input data-ace-create-project-input type="${step.type}" placeholder="${aceEscapeHtml(step.placeholder)}" value="${aceEscapeHtml(value)}">
      </label>
      ${errorCopy}
      <div class="ace-pagination">${dots}</div>
      <div class="ace-actions">
        <button class="ace-button" type="button" data-ace-action="${aceCreateProjectStep === 0 ? "show-controls" : "create-project-back"}">Back</button>
        <button class="ace-button ace-button--primary" type="button" data-ace-action="${isLast ? "create-project-submit" : "create-project-next"}">${isLast ? "Create" : "Next"}</button>
      </div>
    `;
    aceApplyWidgetPosition();
    aceWidget.querySelector("[data-ace-create-project-input]")?.focus();
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

  function aceResetTransientSurfaceState(options = {}) {
    const preserveCompletedSession = Boolean(options.preserveCompletedSession);
    aceProjectPickerMode = "completed";
    aceProjects = [];
    acePendingClearStaleBinding = null;
    acePromptError = "";
    aceSyncStatus = "";
    aceSyncMessage = "";
    aceCatchUpCandidate = null;
    aceRecoveryCandidate = null;
    aceIssueDraft = null;
    aceIssueReturnState = "idle";
    aceCurrentIssues = [];
    aceIssueStatus = "";
    aceCreateProjectDraft = null;
    aceCreateProjectStep = 0;
    aceCreateProjectError = "";
    if (!preserveCompletedSession) {
      aceCompletedSession = null;
    }
  }

  async function aceRefreshTabTitleMetadata(surface) {
    if (!surface?.manuscriptSurfaceId) {
      return;
    }

    if (aceCurrentBinding?.manuscriptSurfaceId === surface.manuscriptSurfaceId && surface.tabTitle) {
      aceCurrentBinding = aceSurfaceRecordWithMetadata(aceCurrentBinding, surface);
      await aceSaveLocalDocumentBinding(surface, aceCurrentBinding.project || null);
    }
    if (aceActiveSession?.manuscriptSurfaceId === surface.manuscriptSurfaceId && surface.tabTitle && aceActiveSession.tabTitle !== surface.tabTitle) {
      aceActiveSession = {
        ...aceActiveSession,
        tabTitle: surface.tabTitle,
        manuscriptSurfaceLabel: surface.manuscriptSurfaceLabel
      };
      await acePersistActiveSession();
    }
  }

  function aceCompletedSessionBelongsToSurface(surface) {
    return Boolean(
      aceCompletedSession
      && aceSessionManuscriptSurface(aceCompletedSession).manuscriptSurfaceId === surface?.manuscriptSurfaceId
    );
  }

  async function aceRefreshStateForSurfaceSwitch(surface, trigger = "tab-switch") {
    if (aceSurfaceRefreshInFlight) {
      return;
    }
    aceSurfaceRefreshInFlight = true;
    try {
      const preserveCompletedSession = Boolean(aceCompletedSession && !aceCompletedSessionBelongsToSurface(surface));
      aceResetTransientSurfaceState({ preserveCompletedSession });
      aceCurrentSurface = surface;
      aceCurrentBinding = null;
      aceSelectedProject = null;
      aceLogTabDiagnostic("D-TAB-RESOLVED", aceSurfaceDiagnostic(surface));

      if (!aceSurfaceConfidence(surface)) {
        acePromptError = "Couldn’t identify this tab. Refresh Google Docs.";
        aceState = "prompt";
        aceRenderPrompt();
        aceLogTabDiagnostic("E-TAB-UNKNOWN", aceSurfaceDiagnostic(surface));
        return;
      }

      let pendingSession = null;
      const pending = await acePendingSessions();
      pendingSession = pending.find(function (session) {
        return aceSessionManuscriptSurface(session).manuscriptSurfaceId === surface.manuscriptSurfaceId;
      }) || null;

      const binding = await aceGetBoundProjectForDocument(surface);
      aceCurrentBinding = binding?.projectId ? binding : null;
      aceSelectedProject = binding?.project || null;
      if (pendingSession) {
        aceCompletedSession = pendingSession;
        aceSyncStatus = "pending";
        aceSyncMessage = "Not synced yet.";
        aceState = "completed";
        aceRenderCompleted();
      } else {
        aceState = "prompt";
        aceRenderPrompt();
      }

      aceLogTabDiagnostic("D-TAB-STATE-REFRESH", {
        trigger,
        surface: aceSurfaceDiagnostic(surface),
        bound: Boolean(binding?.projectId),
        projectId: binding?.projectId || "",
        pendingSession: Boolean(pendingSession)
      });
    } finally {
      aceSurfaceRefreshInFlight = false;
    }
  }

  async function aceBlockActiveSessionForTabSwitch(currentSurface) {
    if (!aceActiveSession) {
      return;
    }
    if (!aceActiveSession.tabBlockedAt) {
      aceActiveSession = {
        ...aceActiveSession,
        tabBlockedAt: new Date().toISOString()
      };
      await acePersistActiveSession();
    }
    aceClearTimer();
    aceState = "tab-blocked";
    aceRenderTabChangedBlock(currentSurface);
    aceLogTabDiagnostic("W-TAB-CHANGED-DURING-SESSION", {
      sessionSurface: aceSurfaceDiagnostic(aceSessionManuscriptSurface(aceActiveSession)),
      currentSurface: aceSurfaceDiagnostic(currentSurface)
    });
  }

  async function aceResumeActiveSessionForSurface(surface) {
    if (!aceActiveSession) {
      return;
    }
    if (aceActiveSession.tabBlockedAt) {
      const blockedMs = Math.max(0, Date.now() - new Date(aceActiveSession.tabBlockedAt).getTime());
      aceActiveSession = {
        ...aceActiveSession,
        tabBlockedAt: "",
        pausedDurationMs: Math.max(0, Number(aceActiveSession.pausedDurationMs) || 0) + blockedMs,
        tabTitle: surface.tabTitle || aceActiveSession.tabTitle,
        manuscriptSurfaceLabel: surface.manuscriptSurfaceLabel || aceActiveSession.manuscriptSurfaceLabel
      };
      await acePersistActiveSession();
    }
    aceState = "active";
    aceStartTimer();
  }

  async function aceHandleSurfaceLifecycleChange(trigger = "surface-check") {
    const nextSurface = aceCurrentManuscriptSurface();
    const previousSurfaceId = aceLastResolvedSurfaceId;
    const nextSurfaceId = nextSurface.manuscriptSurfaceId || "";
    if (!previousSurfaceId) {
      aceLastResolvedSurfaceId = nextSurfaceId;
      aceCurrentSurface = nextSurface;
      aceLogTabDiagnostic("D-TAB-RESOLVED", aceSurfaceDiagnostic(nextSurface));
      return;
    }
    if (previousSurfaceId === nextSurfaceId) {
      await aceRefreshTabTitleMetadata(nextSurface);
      if (
        aceActiveSession
        && !aceSessionSurfaceMismatch(aceActiveSession, nextSurface)
        && (aceActiveSession.tabBlockedAt || aceState === "tab-blocked")
      ) {
        await aceResumeActiveSessionForSurface(nextSurface);
      }
      return;
    }

    aceLogTabDiagnostic("D-TAB-SWITCH", {
      from: previousSurfaceId,
      to: nextSurfaceId,
      trigger
    });
    aceLastResolvedSurfaceId = nextSurfaceId;
    aceCurrentSurface = nextSurface;

    if (aceActiveSession) {
      if (aceSessionSurfaceMismatch(aceActiveSession, nextSurface) || !aceSurfaceConfidence(nextSurface)) {
        await aceBlockActiveSessionForTabSwitch(nextSurface);
      } else {
        await aceResumeActiveSessionForSurface(nextSurface);
      }
      return;
    }

    await aceRefreshStateForSurfaceSwitch(nextSurface, trigger);
  }

  function aceStartSurfaceMonitor() {
    if (aceSurfaceMonitorId) {
      return;
    }
    const scheduleCheck = function (trigger) {
      aceRunAsync(aceHandleSurfaceLifecycleChange(trigger), "handle tab lifecycle change");
    };
    window.addEventListener("hashchange", function () {
      scheduleCheck("hashchange");
    });
    window.addEventListener("popstate", function () {
      scheduleCheck("popstate");
    });
    aceSurfaceMonitorId = window.setInterval(function () {
      scheduleCheck("poll");
    }, 750);
  }

  function aceStartTimer() {
    aceClearTimer();
    if (!aceActiveSession) {
      aceRenderActive();
      return;
    }

    aceRenderActive();
    if (aceActiveSession) {
      aceTimerId = window.setInterval(function () {
        aceRenderActive();
        acePersistActiveSessionHeartbeat();
      }, ACE_TIMER_INTERVAL_MS);
    }
  }

  function acePersistActiveSessionHeartbeat() {
    if (!aceActiveSession) {
      return;
    }
    const now = Date.now();
    if (aceLastActiveHeartbeatAt && now - aceLastActiveHeartbeatAt < ACE_ACTIVE_SESSION_HEARTBEAT_MS) {
      return;
    }
    aceLastActiveHeartbeatAt = now;
    aceActiveSession = {
      ...aceActiveSession,
      lastHeartbeatAt: new Date(now).toISOString(),
      elapsedMsAtLastHeartbeat: aceElapsedMs()
    };
    aceRunAsync(acePersistActiveSession(), "persist active session heartbeat");
  }

  function aceShowStartPrompt() {
    if (aceState !== "idle") {
      return;
    }

    aceState = "binding-loading";
    aceClearActivityTimers();
    aceRenderLoading("Checking project...", "Looking up this manuscript.");
    aceRunAsync(aceLoadPromptBinding(), "load prompt binding");
  }

  async function aceLoadPromptBinding() {
    await aceRefreshCurrentBinding();
    aceState = "prompt";
    aceRenderPrompt();
  }

  async function aceRefreshCurrentBinding() {
    const surface = aceCurrentManuscriptSurface();
    aceCurrentSurface = surface;
    aceCurrentBinding = null;
    if (!surface.documentId || !surface.manuscriptSurfaceId) {
      acePromptError = "Couldn’t identify this tab. Refresh Google Docs.";
      return null;
    }
    try {
      aceCurrentBinding = await aceGetBoundProjectForDocument(surface);
      aceSelectedProject = aceCurrentBinding?.project || null;
    } catch (error) {
      acePromptError = error.message || "Project lookup failed.";
    }
    return aceCurrentBinding;
  }

  async function aceShowControls() {
    if (aceRecoveryCandidate) {
      aceState = "recovery";
      aceRenderRecoveryModal();
      return;
    }

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

    aceState = "binding-loading";
    acePromptError = "";
    aceRenderLoading("Checking project...", "Looking up this manuscript.");
    const binding = await aceRefreshCurrentBinding();
    if (!binding?.projectId) {
      aceState = "prompt";
      aceRenderPrompt();
      return;
    }
    aceState = "prompt";
    aceRenderPrompt();
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

  async function aceOpenIssueForm() {
    aceRememberIssueReturnState();
    aceClearTimer();
    const surface = aceCurrentManuscriptSurface();
    if (aceActiveSession && (aceSessionSurfaceMismatch(aceActiveSession, surface) || !aceSurfaceConfidence(surface))) {
      await aceBlockActiveSessionForTabSwitch(surface);
      return;
    }
    const binding = aceCurrentBinding?.projectId && aceRecordMatchesSurface(aceCurrentBinding, surface)
      ? aceCurrentBinding
      : await aceGetBoundProjectForDocument(surface);
    if (!binding?.projectId) {
      aceCurrentSurface = surface;
      aceCurrentBinding = null;
      acePromptError = "Bind this manuscript first.";
      aceState = "prompt";
      aceRenderPrompt();
      return;
    }
    aceCurrentBinding = binding;
    aceSelectedProject = binding.project || aceSelectedProject;
    const selectedText = aceGetSelectedText();
    aceIssueDraft = {
      ...surface,
      documentUrl: aceDocumentUrl(),
      extensionIssueId: aceCreateExtensionIssueId(surface.documentId),
      projectId: binding.projectId,
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
    const surface = aceSurfaceFromParts({
      documentId: aceIssueDraft?.documentId || aceExtractDocumentId(),
      tabId: aceIssueDraft?.tabId,
      tabTitle: aceIssueDraft?.tabTitle,
      manuscriptSurfaceId: aceIssueDraft?.manuscriptSurfaceId,
      manuscriptSurfaceLabel: aceIssueDraft?.manuscriptSurfaceLabel
    });
    aceIssueDraft = {
      ...(aceIssueDraft || {}),
      ...surface,
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
      const payload = await aceGetExtensionIssues(aceCurrentManuscriptSurface());
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
    if (!projectId && draft.projectId) {
      projectId = String(draft.projectId);
    }
    if (!projectId) {
      const localBinding = await aceGetLocalDocumentBinding(draft);
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
      aceIssueDraft = null;
      acePromptError = "Bind this manuscript first.";
      aceState = "prompt";
      aceRenderPrompt();
      return;
    }

    const snippet = aceNormalizeIssueNoteText(draft.snippet).slice(0, 500);
    const payload = aceIssueSyncPayload(draft, projectId, note, snippet);

    try {
      const result = await acePostIssue(payload);
      const selectedProject = result.project || project;
      if (selectedProject) {
        await aceSaveLocalDocumentBinding(draft, selectedProject);
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
        aceIssueDraft = null;
        acePromptError = "Bind this manuscript first.";
        aceState = "prompt";
        aceRenderPrompt();
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
      const binding = await aceSaveBinding(draft, project.id);
      const selectedProject = binding.project || project;
      await aceSaveLocalDocumentBinding(draft, selectedProject);
      await aceRefreshDocumentBaselineFromCurrentCount(draft, selectedProject);
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

  async function aceShowBindProjectPicker() {
    const surface = aceCurrentSurface || aceCurrentManuscriptSurface();
    aceCurrentSurface = surface;
    if (!surface.documentId || !surface.manuscriptSurfaceId) {
      acePromptError = "Couldn’t identify this tab. Refresh Google Docs.";
      aceState = "prompt";
      aceRenderPrompt();
      return;
    }

    aceProjectPickerMode = "bind";
    aceState = "project-loading";
    aceRenderLoading("Loading projects...", "Checking bindings.");
    await aceNextFrame();
    try {
      aceProjects = await aceReconcileProjectPickerBindings(await aceGetExtensionProjects());
      aceState = "project-picker";
      aceRenderProjectPicker();
    } catch (error) {
      acePromptError = `Projects unavailable. ${error.message}`;
      aceState = "prompt";
      aceRenderPrompt();
    }
  }

  async function aceConfirmUnbind() {
    const surface = aceCurrentSurface || aceCurrentManuscriptSurface();
    aceState = "binding-saving";
    aceRenderLoading("Unbinding...", "Future sessions only.");
    await aceNextFrame();
    try {
      await aceDeleteBinding(surface);
      await aceRemoveLocalDocumentBinding(surface);
      aceCurrentSurface = surface;
      aceCurrentBinding = null;
      aceSelectedProject = null;
      acePromptError = "";
      aceState = "prompt";
      aceRenderPrompt();
    } catch (error) {
      acePromptError = error.message || "Could not unbind.";
      aceState = "prompt";
      aceRenderPrompt();
    }
  }

  function aceShowClearStaleBindingConfirm(projectId) {
    acePendingClearStaleBinding = aceProjects.find(function (item) {
      const project = item.project || item;
      return String(project.id) === String(projectId);
    }) || null;
    if (!acePendingClearStaleBinding) {
      return;
    }
    aceState = "clear-stale-binding-confirm";
    aceRenderClearStaleBindingConfirm();
  }

  function aceShowDeletedBindingRebindConfirm(item) {
    const project = item?.project || item || null;
    const deletedBinding = aceDeletedBindingForProjectItem(item);
    if (!project?.id || !deletedBinding) {
      return false;
    }
    acePendingDeletedBindingRebind = {
      project,
      deletedBinding
    };
    aceState = "deleted-binding-rebind-confirm";
    aceRenderDeletedBindingRebindConfirm();
    return true;
  }

  async function aceConfirmDeletedBindingRebind() {
    const pending = acePendingDeletedBindingRebind;
    acePendingDeletedBindingRebind = null;
    if (!pending?.project?.id) {
      aceState = "project-picker";
      aceRenderProjectPicker();
      return;
    }
    console.info("[ACE] DELETED BINDING REBIND CONFIRMED", {
      projectId: pending.project.id,
      previousBoundDocumentId: pending.deletedBinding?.documentId || "",
      currentTabDocumentId: aceCurrentManuscriptSurface().documentId,
      validationStatus: "deleted",
      deletionReason: pending.deletedBinding?.staleReason || pending.deletedBinding?.status || ""
    });
    await aceBindCurrentSurfaceToProject(pending.project.id, { skipInitialCatchUp: true });
  }

  function aceCancelDeletedBindingRebind() {
    const pending = acePendingDeletedBindingRebind;
    console.info("[ACE] DELETED BINDING REBIND DECLINED", {
      projectId: pending?.project?.id || "",
      previousBoundDocumentId: pending?.deletedBinding?.documentId || "",
      currentTabDocumentId: aceCurrentManuscriptSurface().documentId,
      validationStatus: "deleted",
      baselineScanRan: false
    });
    acePendingDeletedBindingRebind = null;
    aceState = "project-picker";
    aceRenderProjectPicker();
  }

  async function aceConfirmClearStaleBinding() {
    const staleBinding = acePendingClearStaleBinding;
    if (!staleBinding) {
      aceState = "project-picker";
      aceRenderProjectPicker();
      return;
    }
    aceState = "binding-saving";
    aceRenderLoading("Clearing binding...", "History stays saved.");
    await aceNextFrame();
    try {
      await aceClearStaleProjectBinding(staleBinding);
      acePendingClearStaleBinding = null;
      aceProjects = await aceReconcileProjectPickerBindings(await aceGetExtensionProjects());
      aceState = "project-picker";
      aceRenderProjectPicker();
    } catch (error) {
      acePendingClearStaleBinding = null;
      acePromptError = error.message || "Could not clear binding.";
      aceState = "prompt";
      aceRenderPrompt();
    }
  }

  async function aceBeginCreateProject() {
    const surface = aceCurrentSurface || aceCurrentManuscriptSurface();
    aceCurrentSurface = surface;
    if (!surface.documentId || !surface.manuscriptSurfaceId) {
      acePromptError = "Couldn’t identify this tab. Refresh Google Docs.";
      aceState = "prompt";
      aceRenderPrompt();
      return;
    }
    let currentWordCount = await aceVisibleGoogleDocWordCount();
    if (!Number.isFinite(currentWordCount)) {
      currentWordCount = 0;
    }
    aceCreateProjectDraft = aceDefaultCreateProjectDraft(currentWordCount);
    aceCreateProjectStep = 0;
    aceCreateProjectError = "";
    aceState = "create-project";
    aceRenderCreateProject();
  }

  function aceCreateProjectNext() {
    aceReadCreateProjectStep();
    aceCreateProjectError = "";
    if (aceCreateProjectStep < ACE_CREATE_PROJECT_STEPS.length - 1) {
      aceCreateProjectStep += 1;
    }
    aceRenderCreateProject();
  }

  function aceCreateProjectBack() {
    aceReadCreateProjectStep();
    aceCreateProjectError = "";
    aceCreateProjectStep = Math.max(0, aceCreateProjectStep - 1);
    aceRenderCreateProject();
  }

  async function aceSubmitCreateProject() {
    aceReadCreateProjectStep();
    const error = aceValidateCreateProjectDraft(aceCreateProjectDraft);
    if (error) {
      aceCreateProjectError = error;
      aceRenderCreateProject();
      return;
    }

    const surface = aceCurrentSurface || aceCurrentManuscriptSurface();
    aceState = "project-saving";
    aceRenderLoading("Creating project...", "Binding this manuscript.");
    await aceNextFrame();
    try {
      const project = await aceCreateProject(aceCreateProjectDraft);
      if (!project?.id) {
        throw new Error("Project was not created.");
      }
      const binding = await aceSaveBinding(surface, project.id);
      const selectedProject = binding.project || project;
      await aceSaveLocalDocumentBinding(surface, selectedProject);
      aceCurrentBinding = {
        ...surface,
        projectId: selectedProject.id,
        project: selectedProject
      };
      aceSelectedProject = selectedProject;
      aceCreateProjectDraft = null;
      aceCreateProjectStep = 0;
      aceCreateProjectError = "";
      const reconciliation = await aceReconcileInitialBindBaseline(surface, selectedProject);
      if (reconciliation.candidate) {
        return;
      }
      acePromptError = `Bound to ${selectedProject.bookTitle}.`;
      aceState = "prompt";
      aceRenderPrompt();
    } catch (error) {
      aceCreateProjectError = error.message || "Could not create project.";
      aceState = "create-project";
      aceRenderCreateProject();
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

  function aceCatchUpDecisionTrace({
    trigger = "",
    surface = {},
    baseline = null,
    baselineKey = "",
    baselineIsLegacy = false,
    currentSnapshot = null,
    netWordsChanged = 0,
    eligible = false,
    action = "skip-catch-up",
    reason = ""
  } = {}) {
    const visibleResult = currentSnapshot?.visibleCountDiagnostic || null;
    return {
      trigger,
      surface: {
        documentId: surface.documentId || "",
        tabId: surface.tabId || "",
        tabTitle: surface.tabTitle || "",
        manuscriptSurfaceId: surface.manuscriptSurfaceId || "",
        confident: Boolean(surface.documentId && surface.manuscriptSurfaceId)
      },
      baseline: {
        key: baselineKey || "",
        count: Number.isFinite(Number(baseline?.endDocumentWordCount)) ? Number(baseline.endDocumentWordCount) : null,
        source: aceCatchUpBaselineSource(baseline),
        revisionId: baseline?.revisionId || "",
        syncedAt: baseline?.syncedAt || "",
        isLegacyDocumentId: Boolean(baselineIsLegacy)
      },
      currentCounts: {
        visibleCandidates: currentSnapshot?.visibleCandidates || aceCatchUpVisibleCandidateSummary(visibleResult),
        selectedVisibleCount: aceFiniteNumberOrNull(currentSnapshot?.visibleWordCount),
        stableVisible: visibleResult ? {
          count: aceFiniteNumberOrNull(visibleResult.count),
          stable: Boolean(visibleResult.stable),
          reason: visibleResult.reason || "",
          diagnostic: visibleResult.diagnostic || ""
        } : null,
        apiCount: aceFiniteNumberOrNull(currentSnapshot?.apiWordCount),
        apiRevisionId: currentSnapshot?.revisionId || "",
        selectedSource: currentSnapshot?.currentCountSource || "",
        sourceReason: currentSnapshot?.wordCountDiagnostic || ""
      },
      suppression: {
        typingSuppressionActive: aceIsTypingSuppressionActive(),
        lastTypingTimestamp: aceLastTypingTimestamp || 0,
        boundaryTrigger: aceCatchUpTriggerCanPrompt(trigger),
        activeSession: Boolean(aceActiveSession)
      },
      decision: {
        code: aceCatchUpDecisionCode(reason, netWordsChanged, eligible),
        eligibleCode: eligible ? "D-CATCHUP-ELIGIBLE" : "",
        netWordsChanged,
        netKind: netWordsChanged > 0 ? "positive" : netWordsChanged < 0 ? "negative" : "zero",
        eligible: Boolean(eligible),
        action,
        reason
      }
    };
  }

  function aceCatchUpDecisionCode(reason, netWordsChanged, eligible) {
    if (eligible) {
      return netWordsChanged < 0 ? "D-CATCHUP-SHOW-NEGATIVE" : "D-CATCHUP-SHOW-POSITIVE";
    }
    const codes = {
      missingBaseline: "D-CATCHUP-NO-BASELINE",
      "missing-baseline": "D-CATCHUP-NO-BASELINE",
      "no-binding": "D-CATCHUP-NO-BINDING",
      "active-session": "D-CATCHUP-SUPPRESSED-ACTIVE-SESSION",
      "uncertain-surface": "D-CATCHUP-SURFACE-MISMATCH",
      "baseline-surface-mismatch": "D-CATCHUP-SURFACE-MISMATCH",
      "pending-or-completed-session-exists": "D-CATCHUP-SUPPRESSED-PENDING",
      "current-count-diagnostic-only": "D-CATCHUP-NO-CURRENT-COUNT",
      "current-count-unavailable": "D-CATCHUP-NO-CURRENT-COUNT",
      "current-count-untrusted": "D-CATCHUP-NO-CURRENT-COUNT",
      "typing-suppressed": "D-CATCHUP-TYPING-SUPPRESSED",
      "trigger-not-boundary": "D-CATCHUP-NOT-BOUNDARY",
      "zero-net": "D-CATCHUP-NET-ZERO"
    };
    return codes[reason] || "D-CATCHUP-CHECK";
  }

  function aceLogCatchUpDecision(trace) {
    console.info("[ACE] CATCH-UP DECISION", trace);
  }

  function aceSessionCoversCatchUpChange(session, surface, startWordCount, endWordCount) {
    if (!session) {
      return false;
    }
    const sessionSurface = aceSessionManuscriptSurface(session);
    return sessionSurface.manuscriptSurfaceId === surface.manuscriptSurfaceId
      && Number(session.startDocumentWordCount) === Number(startWordCount)
      && Number(session.endDocumentWordCount) === Number(endWordCount);
  }

  async function acePendingSessionForCatchUpChange(surface, startWordCount, endWordCount) {
    const pending = await acePendingSessions();
    return pending.find(function (session) {
      return aceSessionCoversCatchUpChange(session, surface, startWordCount, endWordCount);
    }) || null;
  }

  function aceEvaluateCatchUpCandidate({
    trigger = "",
    surface = {},
    baseline = null,
    baselineKey = "",
    baselineIsLegacy = false,
    currentSnapshot = null,
    pendingSession = null,
    completedSession = null,
    binding = null,
    promptSuppressionReason = ""
  } = {}) {
    const baselineWordCount = Math.max(0, Number(baseline?.endDocumentWordCount) || 0);
    const currentWordCount = Math.max(0, Number(currentSnapshot?.wordCount) || 0);
    const netWordsChanged = aceCalculateNetWordDelta(baselineWordCount, currentWordCount);
    let reason = "";

    if (!surface.documentId || !surface.manuscriptSurfaceId) {
      reason = "uncertain-surface";
    } else if (!baseline || !Number.isFinite(Number(baseline.endDocumentWordCount))) {
      reason = "missing-baseline";
    } else if (baselineIsLegacy || baselineKey !== surface.manuscriptSurfaceId) {
      reason = "baseline-surface-mismatch";
    } else if (!binding?.projectId) {
      reason = "no-binding";
    } else if (aceActiveSession) {
      reason = "active-session";
    } else if (pendingSession || completedSession) {
      reason = "pending-or-completed-session-exists";
    } else if (currentSnapshot?.skipCatchUp) {
      reason = "current-count-diagnostic-only";
    } else if (!currentSnapshot?.ok || !Number.isFinite(Number(currentSnapshot.wordCount))) {
      reason = "current-count-unavailable";
    } else if (currentSnapshot.currentCountTrusted === false) {
      reason = "current-count-untrusted";
    } else if (netWordsChanged === 0) {
      reason = "zero-net";
    } else if (promptSuppressionReason) {
      reason = promptSuppressionReason;
    }

    const eligible = !reason;
    const trace = aceCatchUpDecisionTrace({
      trigger,
      surface,
      baseline,
      baselineKey,
      baselineIsLegacy,
      currentSnapshot,
      netWordsChanged,
      eligible,
      action: eligible ? "show-catch-up" : "skip-catch-up",
      reason: reason || (netWordsChanged < 0 ? "negative-stable-current-count" : "positive-stable-current-count")
    });

    if (!eligible) {
      return { candidate: null, trace, reason };
    }

    return {
      candidate: {
        ...surface,
        documentUrl: aceDocumentUrl(),
        baseline,
        currentSnapshot,
        startDocumentWordCount: baselineWordCount,
        endDocumentWordCount: currentWordCount,
        endDocumentWordCounts: null,
        wordsWritten: Math.max(0, netWordsChanged),
        wordsAdded: 0,
        wordsRemoved: 0,
        wordsEdited: 0,
        netWordsChanged,
        sessionType: netWordsChanged < 0 ? "editing" : "writing"
      },
      trace,
      reason: ""
    };
  }

  async function aceBuildCatchUpCandidate(documentId, trigger = "unknown") {
    const surface = aceCurrentManuscriptSurface(documentId);
    const baselineInfo = await aceGetDocumentBaselineForCatchUp(surface);
    const baseline = baselineInfo.baseline;
    if (!baseline || !Number.isFinite(Number(baseline.endDocumentWordCount))) {
      aceLogCatchUpDecision(aceCatchUpDecisionTrace({
        trigger,
        surface,
        baseline,
        baselineKey: baselineInfo.key,
        baselineIsLegacy: baselineInfo.isLegacy,
        reason: "missing-baseline",
        action: "skip-catch-up"
      }));
      return { candidate: null, error: "", reason: "missing-baseline" };
    }

    const binding = await aceGetBoundProjectForDocument(surface);
    if (!binding?.projectId) {
      const trace = aceCatchUpDecisionTrace({
        trigger,
        surface,
        baseline,
        baselineKey: baselineInfo.key,
        baselineIsLegacy: baselineInfo.isLegacy,
        reason: "no-binding",
        action: "skip-catch-up"
      });
      aceLogCatchUpDecision(trace);
      return { candidate: null, error: "", reason: "no-binding" };
    }
    const currentSnapshot = await aceGoogleDocWordCountAfterSettle(documentId, baseline, { trigger });
    const currentWordCount = Math.max(0, Number(currentSnapshot?.wordCount) || 0);
    const baselineWordCount = Math.max(0, Number(baseline?.endDocumentWordCount) || 0);
    const pendingSession = await acePendingSessionForCatchUpChange(surface, baselineWordCount, currentWordCount);
    const completedSession = aceSessionCoversCatchUpChange(aceCompletedSession, surface, baselineWordCount, currentWordCount)
      ? aceCompletedSession
      : null;
    const catchUpBaseline = {
      ...baseline,
      projectId: binding.projectId,
      project: binding.project || baseline.project || null
    };
    const result = aceEvaluateCatchUpCandidate({
      trigger,
      surface,
      baseline: catchUpBaseline,
      baselineKey: baselineInfo.key,
      baselineIsLegacy: baselineInfo.isLegacy,
      currentSnapshot,
      pendingSession,
      completedSession,
      binding,
      promptSuppressionReason: aceCatchUpSuppressionReason(trigger)
    });
    aceLogCatchUpDecision(result.trace);
    return {
      candidate: result.candidate,
      error: "",
      reason: result.reason,
      trace: result.trace,
      netWordsChanged: result.trace?.decision?.netWordsChanged || 0
    };
  }

  async function aceCheckCatchUpBeforeStartPrompt() {
    const startingSurfaceId = aceCurrentManuscriptSurface().manuscriptSurfaceId;
    aceState = "checking-catch-up";
    acePromptError = "";
    aceRenderLoading("Checking progress...", "Looking for missed words.");
    await aceNextFrame();

    const catchUpResult = await aceBuildCatchUpCandidate(aceExtractDocumentId(), "show-controls");
    if (aceCurrentManuscriptSurface().manuscriptSurfaceId !== startingSurfaceId) {
      return;
    }
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

  async function aceCheckCatchUpForCurrentSurface(trigger = "tab-switch") {
    const startingSurfaceId = aceCurrentManuscriptSurface().manuscriptSurfaceId;
    const catchUpResult = await aceBuildCatchUpCandidate(aceExtractDocumentId(), trigger);
    if (aceCurrentManuscriptSurface().manuscriptSurfaceId !== startingSurfaceId) {
      return;
    }
    if (!catchUpResult.candidate) {
      return;
    }
    aceCatchUpCandidate = catchUpResult.candidate;
    aceSyncStatus = "";
    aceSyncMessage = "";
    aceState = "catch-up";
    aceRenderCatchUpPrompt();
  }

  async function aceManualSyncDocumentChanges() {
    if (aceActiveSession) {
      aceState = "active";
      aceStartTimer();
      return;
    }

    const startingSurfaceId = aceCurrentManuscriptSurface().manuscriptSurfaceId;
    aceState = "checking-catch-up";
    acePromptError = "";
    aceRenderLoading("Syncing changes...", "Checking this document.");
    await aceNextFrame();

    const catchUpResult = await aceBuildCatchUpCandidate(aceExtractDocumentId(), "manual-sync");
    if (aceCurrentManuscriptSurface().manuscriptSurfaceId !== startingSurfaceId) {
      return;
    }
    if (catchUpResult.candidate) {
      aceCatchUpCandidate = catchUpResult.candidate;
      aceSyncStatus = "";
      aceSyncMessage = "";
      aceState = "catch-up";
      aceRenderCatchUpPrompt();
      return;
    }

    aceState = "prompt";
    acePromptError = catchUpResult.reason === "zero-net"
      ? "Document changes are already synced."
      : "Could not verify document changes right now.";
    aceRenderPrompt();
  }

  async function aceSkipCatchUp() {
    const catchUpCandidate = aceCatchUpCandidate;
    await aceSaveSkippedCatchUpBaseline(catchUpCandidate);
    if (aceCatchUpCandidate !== catchUpCandidate) {
      return;
    }
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
    surface = null,
    extensionSessionId,
    sessionType,
    startSnapshot,
    projectId = "",
    project = null,
    hadDocumentActivity = false
  }) {
    const now = new Date().toISOString();
    const manuscriptSurface = surface ? aceSurfaceFromParts(surface) : aceCurrentManuscriptSurface(documentId);
    const sessionScope = await aceGetCurrentDocumentScope(projectId || "");
    acePromptError = "";
    aceState = "active";
    aceActiveSession = {
      startedAt: now,
      sessionType: aceNormalizeSessionType(sessionType),
      ...manuscriptSurface,
      sessionScope: aceCreateSessionScope(manuscriptSurface, sessionScope),
      chromeTabId: sessionScope.chromeTabId,
      pageInstanceId: sessionScope.pageInstanceId,
      projectId: projectId || "",
      project,
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
      measurementPath: aceMeasurementPathForSession(aceActiveSession)
    });
    aceStartTimer();
    if (!projectId) {
      aceRunAsync(aceResolveBindingForActiveSession(), "resolve active session binding");
    }
  }

  async function aceStartSession(sessionTypeOverride = "") {
    if (aceState !== "prompt") {
      return;
    }

    aceState = "starting";
    acePromptError = "";
    aceRenderLoading("Starting...", "Checking Google Docs.");
    await aceNextFrame();

    const documentId = aceExtractDocumentId();
    const surface = aceCurrentManuscriptSurface(documentId);
    const binding = aceCurrentBinding?.projectId && aceRecordMatchesSurface(aceCurrentBinding, surface)
      ? aceCurrentBinding
      : await aceGetBoundProjectForDocument(surface);
    if (!binding?.projectId) {
      aceCurrentSurface = surface;
      aceCurrentBinding = null;
      aceState = "prompt";
      acePromptError = "Bind this manuscript first.";
      aceRenderPrompt();
      return;
    }
    const catchUpResult = await aceBuildCatchUpCandidate(documentId, "pre-session");
    if (catchUpResult.candidate) {
      aceCatchUpCandidate = catchUpResult.candidate;
      aceSyncStatus = "";
      aceSyncMessage = "";
      aceState = "catch-up";
      aceRenderCatchUpPrompt();
      return;
    }

    const extensionSessionId = aceCreateExtensionSessionId(documentId);
    const sessionType = aceNormalizeSessionType(sessionTypeOverride || await aceLastSessionType());
    const startSnapshot = await aceStartSnapshotWithVisibleFallback(
      documentId,
      extensionSessionId,
      true,
      "manual start",
      { allowVisibleFallback: true }
    );
    if (!startSnapshot.ok || !Number.isFinite(startSnapshot.wordCount)) {
      aceState = "prompt";
      acePromptError = `Google Docs word count is required. ${startSnapshot.error || "Check Google OAuth and try again."}`;
      aceRenderPrompt();
      return;
    }

    await aceActivateSessionFromSnapshot({
      documentId,
      surface,
      extensionSessionId,
      sessionType,
      startSnapshot,
      projectId: binding.projectId,
      project: binding.project
    });
    aceSelectedProject = binding.project || aceSelectedProject;
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
    if (aceIsTypingSuppressionActive()) {
      return;
    }

    const documentId = aceExtractDocumentId();
    if (!documentId) {
      return;
    }
    const surface = aceCurrentManuscriptSurface(documentId);

    aceAutoStartInFlight = true;
    try {
      const binding = await aceGetBoundProjectForDocument(surface);
      if (
        !binding?.projectId
        || aceState !== "idle"
        || aceActiveSession
        || aceCompletedSession
        || aceCatchUpCandidate
      ) {
        return;
      }

      const catchUpResult = await aceBuildCatchUpCandidate(documentId, "auto-start-attempt");
      if (
        aceState !== "idle"
        || aceActiveSession
        || aceCompletedSession
        || aceCatchUpCandidate
      ) {
        return;
      }
      if (catchUpResult.candidate) {
        aceCatchUpCandidate = catchUpResult.candidate;
        aceSyncStatus = "";
        aceSyncMessage = "";
        aceState = "catch-up";
        aceRenderCatchUpPrompt();
        return;
      }
      if (["typing-suppressed", "trigger-not-boundary"].includes(catchUpResult.reason)) {
        return;
      }
      if (["current-count-unavailable", "current-count-untrusted", "current-count-diagnostic-only"].includes(catchUpResult.reason)) {
        aceState = "prompt";
        acePromptError = "Could not auto-start session because catch-up could not verify the current word count.";
        aceRenderPrompt();
        return;
      }

      aceState = "starting";
      acePromptError = "";
      aceRenderLoading("Starting session...", "Connected Scriptor project found.");
      await aceNextFrame();

      const extensionSessionId = aceCreateExtensionSessionId(documentId);
      const sessionType = await aceLastSessionType();
      const baseline = await aceGetDocumentBaseline(surface);
      let startSnapshot = await aceSeedGoogleDocStartSnapshotFromBaseline(
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
          { allowVisibleFallback: true }
        );
      }

      if (!startSnapshot.ok || !Number.isFinite(startSnapshot.wordCount)) {
        aceState = "prompt";
        acePromptError = `Could not auto-start session. ${startSnapshot.error || "Check Google OAuth and try again."}`;
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
        project: binding.project,
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
    if (aceIsTypingSuppressionActive()) {
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
    const currentSurface = aceCurrentManuscriptSurface();
    const abandonedSession = await aceAbandonedSessionForCurrentSurface(currentSurface);
    if (abandonedSession) {
      await aceShowAbandonedSessionRecovery(abandonedSession, "abandoned-record");
      return;
    }

    if (!activeSession?.startedAt || !activeSession?.extensionSessionId) {
      await aceRestorePendingSessionForCurrentDocument();
      return;
    }

    const currentScope = await aceGetCurrentDocumentScope(activeSession.projectId || "");
    const scopeValidation = aceValidateSessionScope(activeSession, currentScope);
    if (!scopeValidation.ok) {
      aceLogTabDiagnostic("W-SESSION-SCOPE-MISMATCH-RESTORE", {
        reason: scopeValidation.reason,
        sessionScope: scopeValidation.sessionScope,
        currentScope: scopeValidation.currentScope
      });
      await aceRestorePendingSessionForCurrentDocument();
      return;
    }
    const pending = await acePendingSessions();
    if (pending.some(function (session) {
      return session.extensionSessionId === activeSession.extensionSessionId;
    })) {
      await aceRestorePendingSessionForCurrentDocument();
      return;
    }

    const storedAbandonedSession = await aceStoreAbandonedSession(activeSession, "restored-active-session");
    await aceShowAbandonedSessionRecovery(storedAbandonedSession || activeSession, "restored-active-session");
  }

  async function aceShowAbandonedSessionRecovery(abandonedSession, trigger = "recovery") {
    if (!abandonedSession?.extensionSessionId) {
      return false;
    }
    aceState = "recovery-loading";
    aceActiveSession = null;
    aceCompletedSession = null;
    aceCatchUpCandidate = null;
    aceSyncStatus = "";
    aceSyncMessage = "";
    aceRenderLoading("Recovering session...", "Checking this document.");
    await aceNextFrame();

    const candidate = await aceBuildRecoveryCandidate(abandonedSession, trigger);
    aceRecoveryCandidate = candidate;
    aceSelectedProject = candidate.abandonedSession.project || aceSelectedProject;
    aceState = "recovery";
    aceRenderRecoveryModal();
    return true;
  }

  async function aceBuildRecoveryCandidate(abandonedSession, trigger = "recovery") {
    const startDocumentWordCount = Number.isFinite(Number(abandonedSession.startDocumentWordCount))
      ? Math.max(0, Number(abandonedSession.startDocumentWordCount))
      : null;
    const baseline = {
      ...abandonedSession,
      endDocumentWordCount: Number.isFinite(startDocumentWordCount) ? startDocumentWordCount : 0
    };
    let currentSnapshot = null;
    if (Number.isFinite(startDocumentWordCount)) {
      currentSnapshot = await aceGoogleDocWordCountAfterSettle(abandonedSession.documentId, baseline, {
        trigger,
        apiSurfaceTrusted: true
      });
    }
    const endDocumentWordCount = Number.isFinite(Number(currentSnapshot?.wordCount))
      ? Math.max(0, Number(currentSnapshot.wordCount))
      : null;
    const netWordsChanged = Number.isFinite(endDocumentWordCount) && Number.isFinite(startDocumentWordCount)
      ? endDocumentWordCount - startDocumentWordCount
      : null;
    return {
      abandonedSession,
      currentSnapshot,
      startDocumentWordCount,
      endDocumentWordCount,
      netWordsChanged,
      elapsedMs: aceElapsedMsForSession(abandonedSession),
      measurementPending: !Number.isFinite(endDocumentWordCount)
    };
  }

  async function aceRestorePendingSessionForCurrentDocument() {
    const documentId = aceExtractDocumentId();
    const currentSurface = aceCurrentManuscriptSurface(documentId);
    const pending = await acePendingSessions();
    const pendingSession = pending.find(function (session) {
      const sessionSurface = aceSessionManuscriptSurface(session);
      if (sessionSurface.manuscriptSurfaceId) {
        return sessionSurface.manuscriptSurfaceId === currentSurface.manuscriptSurfaceId;
      }
      return session.documentId === documentId && currentSurface.tabId === "default";
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
    aceSelectedProject = pendingSession.project || aceSelectedProject;
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
      wordsEdited: 0,
      wordsAdded: 0,
      wordsRemoved: 0,
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
    if ((aceState !== "active" && aceState !== "active-minimized" && aceState !== "tab-blocked") || !activeSession) {
      return;
    }

    const endedAt = new Date().toISOString();
    const elapsedMs = aceElapsedMs();
    const currentScope = await aceGetCurrentDocumentScope(activeSession.projectId || "");
    const scopeValidation = aceValidateSessionScope(activeSession, currentScope);

    if (!scopeValidation.ok) {
      aceClearActivityTimers();
      aceClearTimer();
      aceState = "tab-blocked";
      aceSyncStatus = "pending";
      aceSyncMessage = scopeValidation.message;
      aceLogTabDiagnostic("W-SESSION-SCOPE-MISMATCH", {
        reason: scopeValidation.reason,
        sessionScope: scopeValidation.sessionScope,
        currentScope: scopeValidation.currentScope
      });
      if (!fromUnload) {
        aceRenderTabChangedBlock(aceSurfaceFromParts(currentScope));
      }
      return;
    }

    aceState = "ending";
    aceClearActivityTimers();
    aceClearTimer();

    const documentId = activeSession.documentId || currentScope.documentId;
    if (fromUnload) {
      await aceStoreAutoEndedSession(activeSession, endedAt, elapsedMs, documentId);
      return;
    }

    aceRenderLoading("Ending session...", "Measuring words from Google Docs.");
    await aceNextFrame();

    const wordCountResult = await aceGoogleDocNetAfterSave(
      documentId,
      activeSession.extensionSessionId,
      activeSession.startDocumentWordCount,
      activeSession.startDocumentRevisionId,
      activeSession.hadDocumentActivity,
      {
        allowVisibleFallback: true,
        ...aceSessionManuscriptSurface(activeSession)
      }
    );
    const endDocumentWordCount = Number.isFinite(wordCountResult.wordCount)
      ? wordCountResult.wordCount
      : null;
    const netWordsChanged = Number.isFinite(wordCountResult.netWordsChanged)
      ? wordCountResult.netWordsChanged
      : Number.isFinite(activeSession.startDocumentWordCount) && Number.isFinite(endDocumentWordCount)
        ? endDocumentWordCount - activeSession.startDocumentWordCount
        : 0;
    const measurementPending = !wordCountResult.ok || !Number.isFinite(endDocumentWordCount);
    const measuredWordsWritten = Math.max(0, netWordsChanged);
    const wordCountError = measurementPending
      ? wordCountResult.error || "Google Docs word count unavailable."
      : "";
    const wordCountDiagnostic = wordCountResult.wordCountDiagnostic || activeSession.wordCountDiagnostic || aceGoogleDocNetDiagnostic({
        code: measurementPending ? "E-API-UNAVAILABLE" : "D-NET-WORD-COUNT",
        startWordCount: activeSession.startDocumentWordCount,
        apiEndWordCount: endDocumentWordCount,
        visibleEndWordCount: wordCountResult.visibleWordCount,
        netWordsChanged,
        revisionChanged: Boolean(
          activeSession.startDocumentRevisionId
          && wordCountResult.revisionId
          && wordCountResult.revisionId !== activeSession.startDocumentRevisionId
        ),
        startRevisionId: activeSession.startDocumentRevisionId || "",
        endRevisionId: wordCountResult.revisionId || "",
        startSource: activeSession.wordCountMethod || "stored-start-count",
        endSource: wordCountResult.wordCountMethod || wordCountResult.method || "google-docs-api"
      });
    console.info("[ACE] SESSION END", {
      sessionType: activeSession.sessionType,
      visibleWordCount: wordCountResult.visibleWordCount ?? null,
      apiWordCount: wordCountResult.apiWordCount ?? wordCountResult.wordCount ?? null,
      revisionId: wordCountResult.revisionId || "",
      measurementPath: measurementPending ? "measurement-unavailable" : "google-docs-net-count"
    });
    console.info("[ACE] NET WORD COUNT", {
      measurementPath: measurementPending ? "measurement-unavailable" : "google-docs-net-count",
      startDocumentWordCount: activeSession.startDocumentWordCount,
      endDocumentWordCount,
      netWordsChanged
    });
    aceCompletedSession = {
      ...aceSessionManuscriptSurface(activeSession),
      sessionScope: activeSession.sessionScope || aceSessionScope(activeSession),
      chromeTabId: activeSession.chromeTabId || activeSession.sessionScope?.chromeTabId || "",
      pageInstanceId: activeSession.pageInstanceId || activeSession.sessionScope?.pageInstanceId || "",
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
      wordsEdited: 0,
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged: measurementPending ? 0 : netWordsChanged,
      startDocumentWordCount: Number.isFinite(activeSession.startDocumentWordCount)
        ? activeSession.startDocumentWordCount
        : null,
      startDocumentRevisionId: activeSession.startDocumentRevisionId || "",
      endDocumentWordCount: Number.isFinite(endDocumentWordCount)
        ? endDocumentWordCount
        : null,
      endDocumentWordCounts: null,
      endDocumentRevisionId: wordCountResult.revisionId || "",
      wordCountTokenizerVersion: wordCountResult.wordCountTokenizerVersion || "",
      wordCountMethod: wordCountResult.wordCountMethod || wordCountResult.method || "google-docs-api",
      wordCountError,
      wordCountDiagnostic,
      hadDocumentActivity: Boolean(activeSession.hadDocumentActivity),
      measurementPending,
      timing: wordCountResult.timing || null
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
      ...aceSessionManuscriptSurface({
        ...activeSession,
        documentId: activeSession.documentId || documentId
      }),
      sessionScope: activeSession.sessionScope || aceSessionScope(activeSession),
      chromeTabId: activeSession.chromeTabId || activeSession.sessionScope?.chromeTabId || "",
      pageInstanceId: activeSession.pageInstanceId || activeSession.sessionScope?.pageInstanceId || "",
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
      wordCountMethod: activeSession.wordCountMethod || "abandoned-session",
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
    aceRecoveryCandidate = null;
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

  function aceBuildRecoveredCompletedSession(recoveryCandidate) {
    const abandonedSession = recoveryCandidate?.abandonedSession || {};
    const endDocumentWordCount = Number.isFinite(Number(recoveryCandidate?.endDocumentWordCount))
      ? Math.max(0, Number(recoveryCandidate.endDocumentWordCount))
      : null;
    const startDocumentWordCount = Number.isFinite(Number(recoveryCandidate?.startDocumentWordCount))
      ? Math.max(0, Number(recoveryCandidate.startDocumentWordCount))
      : null;
    const netWordsChanged = Number.isFinite(startDocumentWordCount) && Number.isFinite(endDocumentWordCount)
      ? endDocumentWordCount - startDocumentWordCount
      : 0;
    const measurementPending = !Number.isFinite(endDocumentWordCount);
    const currentSnapshot = recoveryCandidate?.currentSnapshot || {};

    return {
      ...aceSessionManuscriptSurface(abandonedSession),
      sessionScope: abandonedSession.sessionScope || aceSessionScope(abandonedSession),
      chromeTabId: abandonedSession.chromeTabId || abandonedSession.sessionScope?.chromeTabId || "",
      pageInstanceId: abandonedSession.pageInstanceId || abandonedSession.sessionScope?.pageInstanceId || "",
      projectId: abandonedSession.projectId || "",
      project: abandonedSession.project || null,
      sessionType: aceNormalizeSessionType(abandonedSession.sessionType),
      startedAt: abandonedSession.startedAt,
      endedAt: new Date().toISOString(),
      durationMinutes: aceDurationMinutes(recoveryCandidate?.elapsedMs || aceElapsedMsForSession(abandonedSession)),
      source: "chrome-extension",
      documentUrl: abandonedSession.documentUrl || aceDocumentUrl(),
      notes: "Recovered abandoned session.",
      extensionSessionId: abandonedSession.extensionSessionId,
      wordsWritten: !measurementPending && abandonedSession.sessionType === "writing" ? Math.max(0, netWordsChanged) : 0,
      wordsEdited: 0,
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged: measurementPending ? 0 : netWordsChanged,
      startDocumentWordCount,
      startDocumentRevisionId: abandonedSession.startDocumentRevisionId || "",
      endDocumentWordCount,
      endDocumentWordCounts: null,
      endDocumentRevisionId: currentSnapshot.revisionId || "",
      wordCountTokenizerVersion: currentSnapshot.wordCountTokenizerVersion || "",
      wordCountMethod: currentSnapshot.currentCountSource || currentSnapshot.method || abandonedSession.wordCountMethod || "recovery",
      wordCountError: measurementPending ? "Google Docs word count unavailable during recovery." : "",
      wordCountDiagnostic: currentSnapshot.wordCountDiagnostic || abandonedSession.wordCountDiagnostic || "",
      hadDocumentActivity: Boolean(abandonedSession.hadDocumentActivity),
      measurementPending,
      recoveredFromAbandonedSession: true
    };
  }

  async function aceRecoverAbandonedSession() {
    const recoveryCandidate = aceRecoveryCandidate;
    const abandonedSession = recoveryCandidate?.abandonedSession;
    if (!abandonedSession?.extensionSessionId) {
      return;
    }

    aceState = "recovery-syncing";
    aceSyncStatus = "";
    aceSyncMessage = "";
    aceRenderLoading("Recovering session...", "Saving your work.");
    await aceNextFrame();

    let candidate = recoveryCandidate;
    if (!Number.isFinite(Number(candidate.endDocumentWordCount))) {
      candidate = await aceBuildRecoveryCandidate(abandonedSession, "recover-session");
      if (aceRecoveryCandidate !== recoveryCandidate) {
        return;
      }
      aceRecoveryCandidate = candidate;
    }

    aceCompletedSession = aceBuildRecoveredCompletedSession(candidate);
    aceActiveSession = null;
    aceRecoveryCandidate = null;
    aceSelectedProject = aceCompletedSession.project || aceSelectedProject;
    aceState = "completed";
    aceSyncStatus = "syncing";
    aceSyncMessage = "Syncing...";
    await acePersistActiveSession();
    aceRenderCompleted();
    await aceSyncCompletedSession();
    await aceRemoveAbandonedSession(abandonedSession.extensionSessionId);
    await aceRemoveStoredActiveSessionIfMatches(abandonedSession.extensionSessionId);
  }

  async function aceDiscardAbandonedSession() {
    const abandonedSession = aceRecoveryCandidate?.abandonedSession;
    if (!abandonedSession?.extensionSessionId) {
      return;
    }
    await aceRemoveAbandonedSession(abandonedSession.extensionSessionId);
    await aceRemoveStoredActiveSessionIfMatches(abandonedSession.extensionSessionId);
    aceRecoveryCandidate = null;
    aceActiveSession = null;
    aceCompletedSession = null;
    aceSyncStatus = "";
    aceSyncMessage = "";
    aceState = "idle";
    aceShowStartPrompt();
  }

  async function aceResolveBindingForActiveSession() {
    const extensionSessionId = aceActiveSession?.extensionSessionId;
    const documentId = aceActiveSession?.documentId;
    if (!extensionSessionId || !documentId) {
      return;
    }

    try {
      let binding = await aceGetLocalDocumentBinding(aceActiveSession);
      if (!binding?.projectId) {
        const serverProject = await aceGetServerDocumentBinding(aceActiveSession);
        if (serverProject?.id) {
          await aceSaveLocalDocumentBinding(aceActiveSession, serverProject);
          binding = {
            ...aceSessionManuscriptSurface(aceActiveSession),
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

        let binding = await aceGetLocalDocumentBinding(aceCompletedSession);
        if (!binding?.projectId) {
          const serverProject = await aceGetServerDocumentBinding(aceCompletedSession);
          if (serverProject?.id) {
            await aceSaveLocalDocumentBinding(aceCompletedSession, serverProject);
            binding = {
              ...aceSessionManuscriptSurface(aceCompletedSession),
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
    if (aceProjectPickerMode === "bind") {
      const item = aceProjects.find(function (projectItem) {
        const project = projectItem.project || projectItem;
        return String(project.id) === String(projectId);
      });
      if (aceDeletedBindingForProjectItem(item)) {
        aceShowDeletedBindingRebindConfirm(item);
        return;
      }
      await aceBindCurrentSurfaceToProject(projectId);
      return;
    }

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

    const project = aceFindPickerProject(projectId);
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
      const binding = await aceSaveBinding(aceCompletedSession, project.id);
      if (!aceIsCompletedSessionCurrent(completedSessionId)) {
        return;
      }

      aceSelectedProject = binding.project || project;
      await aceSaveLocalDocumentBinding(aceCompletedSession, aceSelectedProject);
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

  function aceFindPickerProject(projectId) {
    const row = aceProjects.find(function (item) {
      const project = item.project || item;
      return String(project.id) === String(projectId);
    });
    return row?.project || row || null;
  }

  async function aceBindCurrentSurfaceToProject(projectId, options = {}) {
    const surface = aceCurrentSurface || aceCurrentManuscriptSurface();
    const project = aceFindPickerProject(projectId);
    if (!surface.documentId || !surface.manuscriptSurfaceId) {
      acePromptError = "Couldn’t identify this tab. Refresh Google Docs.";
      aceState = "prompt";
      aceRenderPrompt();
      return;
    }
    if (!project?.id) {
      return;
    }

    aceState = "binding-saving";
    aceRenderLoading("Binding project...", "Saving this manuscript.");
    await aceNextFrame();
    try {
      const result = await aceSaveBinding(surface, project.id);
      const selectedProject = result.project || project;
      await aceSaveLocalDocumentBinding(surface, selectedProject);
      aceCurrentSurface = surface;
      aceCurrentBinding = {
        ...surface,
        projectId: selectedProject.id,
        project: selectedProject
      };
      aceSelectedProject = selectedProject;
      if (options.skipInitialCatchUp) {
        await aceRefreshDocumentBaselineFromCurrentCount(surface, selectedProject);
      } else {
        const reconciliation = await aceReconcileInitialBindBaseline(surface, selectedProject);
        if (reconciliation.candidate) {
          return;
        }
      }
      acePromptError = `Bound to ${selectedProject.bookTitle}.`;
      aceState = "prompt";
      aceRenderPrompt();
    } catch (error) {
      acePromptError = error.message || "Could not bind project.";
      aceState = "prompt";
      aceRenderPrompt();
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
      const binding = await aceSaveBinding(aceActiveSession, project.id);
      if (!aceIsActiveSessionCurrent(extensionSessionId)) {
        return;
      }

      aceActiveSession = {
        ...aceActiveSession,
        projectId: project.id
      };
      aceSelectedProject = binding.project || project;
      await aceSaveLocalDocumentBinding(aceActiveSession, aceSelectedProject);
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
      const binding = await aceSaveBinding(catchUpCandidate, project.id);
      if (aceCatchUpCandidate !== catchUpCandidate) {
        return;
      }

      const selectedProject = binding.project || project;
      await aceSaveLocalDocumentBinding(catchUpCandidate, selectedProject);
      if (aceCatchUpCandidate !== catchUpCandidate) {
        return;
      }

      await aceSyncCatchUpSession(selectedProject);
    } catch (error) {
      if (aceCatchUpCandidate !== catchUpCandidate) {
        return;
      }

      await aceStorePendingSession(catchUpSession);
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

    if ((Number(adjustedCatchUpCandidate.netWordsChanged) || 0) === 0) {
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
      const binding = await aceGetLocalDocumentBinding(catchUpCandidate);
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

    const catchUpSession = aceBuildCatchUpSession(catchUpCandidate, projectId);

    try {
      const syncStartedAt = aceNowMs();
      const result = await acePostSession(catchUpSession);
      if (catchUpCandidate.currentSnapshot?.timing) {
        catchUpCandidate.currentSnapshot.timing.backendSyncElapsedMs += aceTimingElapsedMs(syncStartedAt);
        aceCompleteWordCountTiming(catchUpCandidate.currentSnapshot.timing, {
          action: "catch-up-synced"
        });
      }
      if (aceCatchUpCandidate !== catchUpCandidate) {
        return;
      }

      const syncedProject = result.project || project || catchUpCandidate.baseline?.project || null;
      await aceSaveDocumentBaseline(catchUpSession, syncedProject);
      console.info("[ACE] CATCH-UP LOGGED", {
        code: "D-CATCHUP-LOGGED",
        trigger: catchUpCandidate.currentSnapshot?.timing?.trigger || "catch-up",
        projectId,
        manuscriptSurfaceId: catchUpCandidate.manuscriptSurfaceId || "",
        baselineWordCount: catchUpCandidate.startDocumentWordCount,
        currentWordCount: catchUpCandidate.endDocumentWordCount,
        netWordDelta: catchUpCandidate.netWordsChanged,
        baselineUpdateSuccess: true,
        reconciliationResult: "logged"
      });
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

      await aceStorePendingSession(catchUpSession);
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
      const syncStartedAt = aceNowMs();
      const result = await acePostSession(sessionToSync);
      if (sessionToSync?.timing) {
        sessionToSync.timing.backendSyncElapsedMs += aceTimingElapsedMs(syncStartedAt);
        aceCompleteWordCountTiming(sessionToSync.timing, {
          action: "session-synced"
        });
      }
      if (!aceIsCompletedSessionCurrent(completedSessionId)) {
        return;
      }

      aceSelectedProject = result.project || aceSelectedProject;
      aceSyncStatus = "synced";
      aceSyncMessage = result.duplicate ? "Already synced." : "Synced.";
      console.info("[ACE] POST-SYNC RESPONSE", {
        measurementPath: aceMeasurementPathForSession(result.session || sessionToSync),
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

    const currentScope = await aceGetCurrentDocumentScope(completedSession.projectId || "");
    const scopeValidation = aceValidateSessionScope(completedSession, currentScope);
    if (!scopeValidation.ok) {
      const surfaceError = scopeValidation.message || "This session belongs to another Google Docs tab.";
      aceCompletedSession = {
        ...aceCompletedSession,
        wordCountError: surfaceError,
        wordCountDiagnostic: `E-SESSION-SCOPE-MISMATCH: reason ${scopeValidation.reason}; started ${scopeValidation.sessionScope.manuscriptSurfaceId}; current ${scopeValidation.currentScope.manuscriptSurfaceId}.`
      };
      await aceMarkSessionUnsynced(surfaceError);
      return false;
    }

    const wordCountResult = await aceGoogleDocNetAfterSave(
      completedSession.documentId,
      completedSession.extensionSessionId,
      completedSession.startDocumentWordCount,
      completedSession.startDocumentRevisionId,
      completedSession.hadDocumentActivity,
      {
        allowVisibleFallback: true,
        ...aceSessionManuscriptSurface(completedSession)
      }
    );
    if (!aceIsCompletedSessionCurrent(completedSessionId)) {
      return false;
    }

    if (!wordCountResult.ok || !Number.isFinite(wordCountResult.wordCount)) {
      const visibleFallback = await aceVisibleStartSnapshot(
        completedSession.documentId,
        wordCountResult.error || "retry Google API count was unavailable"
      );
      if (visibleFallback && Number.isFinite(completedSession.startDocumentWordCount)) {
        const netWordsChanged = visibleFallback.wordCount - completedSession.startDocumentWordCount;
        aceCompletedSession = {
          ...aceCompletedSession,
          wordsWritten: completedSession.sessionType === "writing" ? Math.max(0, netWordsChanged) : 0,
          wordsEdited: 0,
          wordsAdded: 0,
          wordsRemoved: 0,
          netWordsChanged,
          endDocumentWordCount: visibleFallback.wordCount,
          endDocumentWordCounts: null,
          endDocumentRevisionId: "",
          wordCountTokenizerVersion: "",
          wordCountMethod: "visible-total-fallback",
          wordCountError: "",
          wordCountDiagnostic: aceGoogleDocNetDiagnostic({
            code: "W-END-VISIBLE-FALLBACK",
            startWordCount: completedSession.startDocumentWordCount,
            apiEndWordCount: wordCountResult.wordCount,
            visibleEndWordCount: visibleFallback.wordCount,
            netWordsChanged,
            revisionChanged: false,
            startRevisionId: completedSession.startDocumentRevisionId || "",
            endRevisionId: "",
            startSource: completedSession.wordCountMethod || "stored-start-count",
            endSource: "visible-total-fallback"
          }),
          measurementPending: false
        };
        return true;
      }

      aceCompletedSession = {
        ...aceCompletedSession,
        wordCountError: wordCountResult.error || "Try again shortly.",
        wordCountDiagnostic: wordCountResult.wordCountDiagnostic || aceGoogleDocNetDiagnostic({
          code: "E-API-UNAVAILABLE",
          startWordCount: completedSession.startDocumentWordCount,
          apiEndWordCount: wordCountResult.wordCount,
          visibleEndWordCount: wordCountResult.visibleWordCount,
          netWordsChanged: wordCountResult.netWordsChanged,
          revisionChanged: Boolean(
            completedSession.startDocumentRevisionId
            && wordCountResult.revisionId
            && wordCountResult.revisionId !== completedSession.startDocumentRevisionId
          ),
          startRevisionId: completedSession.startDocumentRevisionId || "",
          endRevisionId: wordCountResult.revisionId || "",
          startSource: completedSession.wordCountMethod || "stored-start-count",
          endSource: wordCountResult.wordCountMethod || wordCountResult.method || "google-docs-api"
        })
      };
      await aceMarkSessionUnsynced(`Google Docs count unavailable. ${wordCountResult.error || "Try again shortly."}`);
      return false;
    }

    const netWordsChanged = Number.isFinite(wordCountResult.netWordsChanged)
      ? wordCountResult.netWordsChanged
      : Number(wordCountResult.wordCount) - Number(completedSession.startDocumentWordCount);

    aceCompletedSession = {
      ...aceCompletedSession,
      wordsWritten: completedSession.sessionType === "writing" ? Math.max(0, netWordsChanged) : 0,
      wordsEdited: 0,
      wordsAdded: 0,
      wordsRemoved: 0,
      netWordsChanged,
      endDocumentWordCount: wordCountResult.wordCount,
      endDocumentWordCounts: null,
      endDocumentRevisionId: wordCountResult.revisionId || "",
      wordCountTokenizerVersion: wordCountResult.wordCountTokenizerVersion || "",
      wordCountMethod: wordCountResult.wordCountMethod || wordCountResult.method || "google-docs-api",
      wordCountError: "",
      wordCountDiagnostic: wordCountResult.wordCountDiagnostic || aceGoogleDocNetDiagnostic({
          code: "D-NET-WORD-COUNT",
          startWordCount: completedSession.startDocumentWordCount,
          apiEndWordCount: wordCountResult.wordCount,
          visibleEndWordCount: wordCountResult.visibleWordCount,
          netWordsChanged,
          revisionChanged: Boolean(
            completedSession.startDocumentRevisionId
            && wordCountResult.revisionId
            && wordCountResult.revisionId !== completedSession.startDocumentRevisionId
          ),
          startRevisionId: completedSession.startDocumentRevisionId || "",
          endRevisionId: wordCountResult.revisionId || "",
          startSource: completedSession.wordCountMethod || "stored-start-count",
          endSource: wordCountResult.wordCountMethod || wordCountResult.method || "google-docs-api"
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
    aceLastTypingTimestamp = Date.now();
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
    } else if (action === "start-writing") {
      aceRunAsync(aceStartSession("writing"), "start writing session");
    } else if (action === "start-editing") {
      aceRunAsync(aceStartSession("editing"), "start editing session");
    } else if (action === "manual-sync") {
      aceRunAsync(aceManualSyncDocumentChanges(), "manual sync document changes");
    } else if (action === "bind-project") {
      aceRunAsync(aceShowBindProjectPicker(), "show bind project picker");
    } else if (action === "create-project") {
      aceRunAsync(aceBeginCreateProject(), "start create project");
    } else if (action === "create-project-next") {
      aceCreateProjectNext();
    } else if (action === "create-project-back") {
      aceCreateProjectBack();
    } else if (action === "create-project-submit") {
      aceRunAsync(aceSubmitCreateProject(), "create project");
    } else if (action === "show-unbind") {
      aceState = "unbind-confirm";
      aceRenderUnbindConfirm();
    } else if (action === "cancel-unbind") {
      aceState = "prompt";
      aceRenderPrompt();
    } else if (action === "confirm-unbind") {
      aceRunAsync(aceConfirmUnbind(), "unbind project");
    } else if (action === "clear-stale-binding") {
      aceShowClearStaleBindingConfirm(button.getAttribute("data-ace-project-id"));
    } else if (action === "cancel-clear-stale-binding") {
      acePendingClearStaleBinding = null;
      aceState = "project-picker";
      aceRenderProjectPicker();
    } else if (action === "confirm-clear-stale-binding") {
      aceRunAsync(aceConfirmClearStaleBinding(), "clear stale binding");
    } else if (action === "confirm-deleted-binding-rebind") {
      aceRunAsync(aceConfirmDeletedBindingRebind(), "confirm deleted binding rebind");
    } else if (action === "cancel-deleted-binding-rebind") {
      aceCancelDeletedBindingRebind();
    } else if (action === "show-controls") {
      aceRunAsync(aceShowControls(), "show controls");
    } else if (action === "show-issue-form") {
      aceRunAsync(aceOpenIssueForm(), "show issue form");
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
      aceRunAsync(aceSkipCatchUp(), "skip catch-up session");
    } else if (action === "recover-session") {
      aceRunAsync(aceRecoverAbandonedSession(), "recover abandoned session");
    } else if (action === "discard-recovery") {
      aceRunAsync(aceDiscardAbandonedSession(), "discard abandoned session");
    } else if (action === "decline-start") {
      aceDeclineStart();
    } else if (action === "end") {
      aceRunAsync(aceEndSession(), "end session");
    } else if (action === "go-back-tab") {
      aceRenderTabChangedBlock(aceCurrentManuscriptSurface());
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
          ? aceCatchUpActivity(aceRecalculateCatchUpCandidate(aceCatchUpCandidate, wordCount)).deltaCopy
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
      aceTestState: function () {
        return {
          state: aceState,
          activeSession: aceActiveSession,
          completedSession: aceCompletedSession,
          currentSurface: aceCurrentSurface,
          currentBinding: aceCurrentBinding,
          catchUpCandidate: aceCatchUpCandidate,
          recoveryCandidate: aceRecoveryCandidate,
          pendingDeletedBindingRebind: acePendingDeletedBindingRebind,
          currentIssues: aceCurrentIssues,
          selectedProject: aceSelectedProject,
          widgetHtml: aceWidget?.innerHTML || ""
        };
      },
      aceSetTestState: function (updates = {}) {
        if (Object.prototype.hasOwnProperty.call(updates, "state")) {
          aceState = updates.state;
        }
        if (Object.prototype.hasOwnProperty.call(updates, "activeSession")) {
          aceActiveSession = updates.activeSession;
        }
        if (Object.prototype.hasOwnProperty.call(updates, "completedSession")) {
          aceCompletedSession = updates.completedSession;
        }
        if (Object.prototype.hasOwnProperty.call(updates, "currentSurface")) {
          aceCurrentSurface = updates.currentSurface;
        }
        if (Object.prototype.hasOwnProperty.call(updates, "currentBinding")) {
          aceCurrentBinding = updates.currentBinding;
        }
        if (Object.prototype.hasOwnProperty.call(updates, "selectedProject")) {
          aceSelectedProject = updates.selectedProject;
        }
        if (Object.prototype.hasOwnProperty.call(updates, "projects")) {
          aceProjects = updates.projects;
        }
        if (Object.prototype.hasOwnProperty.call(updates, "projectPickerMode")) {
          aceProjectPickerMode = updates.projectPickerMode;
        }
        if (Object.prototype.hasOwnProperty.call(updates, "catchUpCandidate")) {
          aceCatchUpCandidate = updates.catchUpCandidate;
        }
        if (Object.prototype.hasOwnProperty.call(updates, "recoveryCandidate")) {
          aceRecoveryCandidate = updates.recoveryCandidate;
        }
      },
      aceNormalizeTabId,
      aceNormalizeTabTitle,
      aceCreateManuscriptSurfaceId,
      aceCurrentGoogleDocTabInfo,
      aceCurrentManuscriptSurface,
      aceSurfaceFromParts,
      aceSurfaceConfidence,
      aceCreateSessionScope,
      aceCurrentChromeTabScope,
      aceGetCurrentDocumentScope,
      aceSessionScope,
      aceValidateSessionScope,
      aceRecordMatchesSurface,
      aceFindSurfaceRecord,
      aceGetLocalDocumentBinding,
      aceSaveLocalDocumentBinding,
      aceRemoveLocalDocumentBinding,
      aceGetBoundProjectForDocument,
      aceProjectPickerStatusLabel,
      aceDeletedBindingForProjectItem,
      aceValidateBoundDocument,
      aceClassifyBindingValidation,
      aceReconcileProjectPickerBindings,
      aceClearStaleProjectBinding,
      aceGetDocumentBaseline,
      aceGetDocumentBaselineForCatchUp,
      aceSaveDocumentBaseline,
      aceSaveGoogleDocBaseline,
      aceEnsureDocumentBaseline,
      aceRefreshDocumentBaselineFromCurrentCount,
      aceDefaultCreateProjectDraft,
      aceValidateCreateProjectDraft,
      aceCalculateNetWordDelta,
      aceSessionSurfaceMismatch,
      aceIssueSyncPayload,
      aceSessionNetWordsChanged,
      aceSessionSyncPayload,
      aceMeasurementPathForSession,
      aceSessionWordsCopy,
      aceVisibleWordCountCandidatesInDocument,
      aceVisibleWordCountFromElement,
      aceBestVisibleWordCountCandidate,
      aceStableVisibleGoogleDocWordCount,
      aceVisibleCountClasses,
      aceVisibleCountDiagnostic,
      aceStartSnapshotWithVisibleFallback,
      aceSeedGoogleDocStartSnapshotFromBaseline,
      aceGoogleDocNetAfterSave,
      aceGoogleDocWordCountAfterSettle,
      aceEvaluateCatchUpCandidate,
      aceCatchUpDecisionTrace,
      aceBuildCatchUpSession,
      aceSaveSkippedCatchUpBaseline,
      aceGoogleDocNetDiagnostic,
      aceBuildCatchUpCandidate,
      aceResetTransientSurfaceState,
      aceRefreshStateForSurfaceSwitch,
      aceHandleSurfaceLifecycleChange,
      aceCheckCatchUpForCurrentSurface,
      aceBlockActiveSessionForTabSwitch,
      aceResumeActiveSessionForSurface,
      aceChooseProject,
      aceShowDeletedBindingRebindConfirm,
      aceConfirmDeletedBindingRebind,
      aceCancelDeletedBindingRebind,
      aceStartSession,
      aceEndSession,
      aceAutoStartBoundSessionFromActivity,
      aceManualSyncDocumentChanges,
      aceRegisterWritingActivity,
      aceIsTypingSuppressionActive,
      aceBindCurrentSurfaceToProject,
      acePendingSessions,
      aceStorePendingSession,
      aceRestorePendingSessionForCurrentDocument,
      aceAbandonedSessions,
      aceStoreAbandonedSession,
      aceRemoveAbandonedSession,
      aceBuildRecoveryCandidate,
      aceBuildRecoveredCompletedSession,
      aceRecoverAbandonedSession,
      aceDiscardAbandonedSession,
      aceShowAbandonedSessionRecovery
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
    if (aceState === "recovery" || aceState === "recovery-loading") {
      aceStartSurfaceMonitor();
      return;
    }
    await acePrimeBaselineForCurrentDocument();
    await aceHandleSurfaceLifecycleChange("initialize");
    aceStartSurfaceMonitor();
  }

  function aceHandleDocumentExit() {
    if (aceExitHandled || !aceActiveSession) {
      return;
    }

    aceExitHandled = true;
    aceRunAsync(aceStoreAbandonedSession(aceActiveSession, "document-exit"), "persist abandoned session on document exit");
  }
})();
