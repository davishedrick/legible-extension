(function () {
  "use strict";

  const ACE_API_BASE_URL = "https://davishedrick.pythonanywhere.com";
  const ACE_APP_URL = ACE_API_BASE_URL;
  const ACE_WIDGET_ID = "ace-widget";
  const ACE_START_DELAY_MS = 4000;
  const ACE_ACTIVITY_GAP_MS = 1500;
  const ACE_TIMER_INTERVAL_MS = 1000;
  const ACE_GOOGLE_DOC_SETTLE_DELAY_MS = 1200;
  const ACE_GOOGLE_DOC_POLL_DELAY_MS = 1000;
  const ACE_GOOGLE_DOC_POLL_ATTEMPTS = 5;
  const ACE_ACTIVITY_MESSAGE = "ace-writing-activity";
  const ACE_GOOGLE_DOC_START_SNAPSHOT_MESSAGE = "ace-google-doc-start-snapshot";
  const ACE_GOOGLE_DOC_DIFF_MESSAGE = "ace-google-doc-diff";
  const ACE_IS_TOP_FRAME = window.top === window;

  const ACE_SESSION_STORAGE = {
    widgetPosition: "ace-widget-position"
  };

  const ACE_LOCAL_STORAGE = {
    activeSession: "aceActiveSession",
    pendingSessions: "acePendingSessions"
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

  let aceState = "idle";
  let aceActiveSession = null;
  let aceCompletedSession = null;
  let aceTimerId = null;
  let aceStartTimerId = null;
  let aceActivityGapTimerId = null;
  let aceProjects = [];
  let aceSelectedProject = null;
  let aceProjectPickerMode = "completed";
  let aceSyncStatus = "";
  let aceSyncMessage = "";
  let aceWidget = null;
  let aceWidgetPosition = sessionStorage.getItem(ACE_SESSION_STORAGE.widgetPosition) || ACE_DEFAULT_POSITION;
  let aceDragState = null;
  let acePromptError = "";

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

  async function aceApiFetch(path, options) {
    const response = await new Promise(function (resolve, reject) {
      const runtime = aceChromeRuntime();
      if (!runtime?.runtime?.sendMessage) {
        reject(new Error("Extension context is not available. Refresh the Google Doc."));
        return;
      }

      try {
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
      } catch (error) {
        reject(error);
      }
    });

    if (!response.ok) {
      throw new Error(response.error || "Author Companion API request failed.");
    }

    return response.payload || {};
  }

  async function aceGetBinding(documentId) {
    return aceApiFetch(
      `/api/extension/document-binding?documentId=${encodeURIComponent(documentId)}`
    );
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

  async function acePostSession(session) {
    const payload = aceSessionSyncPayload(session);
    console.info("[ACE] Syncing extension session payload", payload);
    return aceApiFetch("/api/extension/sessions", {
      method: "POST",
      body: JSON.stringify(payload)
    });
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

  function aceDelay(milliseconds) {
    return new Promise(function (resolve) {
      window.setTimeout(resolve, milliseconds);
    });
  }

  async function aceGoogleDocDiffAfterSave(documentId, extensionSessionId, startWordCount) {
    await aceDelay(ACE_GOOGLE_DOC_SETTLE_DELAY_MS);

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
        const isBetterResponse = !bestResponse.ok
          || totalActivity > bestActivity
          || (totalActivity === bestActivity && netMagnitude > bestNetMagnitude);

        if (isBetterResponse) {
          bestResponse = response;
          bestActivity = totalActivity;
          bestNetMagnitude = netMagnitude;
          bestAttempt = attempt + 1;
        }

        if (totalActivity > 0 || delta !== 0) {
          console.info("[ACE] Selected Google Docs word diff", {
            attempt: attempt + 1,
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
    if (aceStartTimerId) {
      window.clearTimeout(aceStartTimerId);
      aceStartTimerId = null;
    }

    if (aceActivityGapTimerId) {
      window.clearTimeout(aceActivityGapTimerId);
      aceActivityGapTimerId = null;
    }
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

  function aceRenderMessage(message) {
    aceWidget.className = "ace-widget ace-widget--prompt";
    aceWidget.innerHTML = `<div class="ace-prompt-copy">${aceEscapeHtml(message)}</div>`;
    aceApplyWidgetPosition();
  }

  function aceRenderIdle() {
    aceWidget.className = "ace-widget ace-widget--idle";
    aceWidget.innerHTML = '<span class="ace-status-dot" aria-hidden="true"></span><span>Idle</span>';
    aceApplyWidgetPosition();
  }

  function aceRenderDetecting() {
    aceWidget.className = "ace-widget ace-widget--idle ace-widget--detecting";
    aceWidget.innerHTML = '<span class="ace-status-dot" aria-hidden="true"></span><span>Listening...</span>';
    aceApplyWidgetPosition();
  }

  function aceRenderPrompt() {
    const errorCopy = acePromptError
      ? `<div class="ace-sync-copy ace-sync-copy--pending">${aceEscapeHtml(aceShortDiagnostic(acePromptError))}</div>`
      : "";
    aceWidget.className = "ace-widget ace-widget--prompt";
    aceWidget.innerHTML = `
      <div class="ace-prompt-copy">Start session?</div>
      ${errorCopy}
      <div class="ace-actions">
        <button class="ace-button ace-button--primary" type="button" data-ace-action="confirm-start">Yes</button>
        <button class="ace-button" type="button" data-ace-action="decline-start">No</button>
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
    const label = aceCapitalize(aceActiveSession?.sessionType || "writing");
    const nextType = aceActiveSession?.sessionType === "writing" ? "Editing" : "Writing";

    aceWidget.className = "ace-widget ace-widget--active";
    aceWidget.innerHTML = `
      <div class="ace-session-line">
        <span class="ace-session-type">${label}</span>
        <span class="ace-time">${aceFormatTimer(aceElapsedMs())}</span>
      </div>
      <div class="ace-actions">
        <button class="ace-button ace-button--end" type="button" data-ace-action="end">End</button>
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
      <div class="ace-completed-copy">${label} session tracked: ${aceFormatCompletedMinutes(aceCompletedSession?.durationMinutes || 1)}${wordsCopy}</div>
      ${projectCopy}
      ${methodCopy}
      ${diagnosticCopy}
      ${statusCopy}
      <div class="ace-actions">
        <button class="ace-button" type="button" data-ace-action="open">Open app</button>
        <button class="ace-button" type="button" data-ace-action="retry-sync" ${retryDisabled}>Retry sync</button>
        <button class="ace-button" type="button" data-ace-action="change-project" ${changeProjectDisabled}>Change project</button>
        ${contextInvalid ? '<button class="ace-button ace-button--primary" type="button" data-ace-action="refresh-page">Refresh doc</button>' : ""}
        <button class="ace-button" type="button" data-ace-action="start-new">Start new</button>
      </div>
    `;
    aceApplyWidgetPosition();
  }

  function aceRenderProjectPicker() {
    const copy = aceProjectPickerMode === "active" ? "Change project" : "Choose project";
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
      <div class="ace-prompt-copy">${copy}</div>
      <div class="ace-project-list">${rows}</div>
      <div class="ace-actions">
        <button class="ace-button" type="button" data-ace-action="open">Open app</button>
        <button class="ace-button" type="button" data-ace-action="retry-sync">Retry</button>
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

  function aceStartTimer() {
    aceClearTimer();
    aceRenderActive();
    aceTimerId = window.setInterval(aceRenderActive, ACE_TIMER_INTERVAL_MS);
  }

  function aceShowStartPrompt() {
    if (aceState !== "idle") {
      return;
    }

    aceState = "prompt";
    aceClearActivityTimers();
    aceRenderPrompt();
  }

  async function aceStartSession() {
    if (aceState !== "prompt") {
      return;
    }

    aceState = "starting";
    acePromptError = "";
    aceRenderMessage("Starting...");

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
    aceResolveBindingForActiveSession();
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
    aceSyncStatus = "pending";
    aceSyncMessage = "Not synced yet.";
    aceSelectedProject = null;
    aceRenderCompleted();
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

  async function aceEndSession() {
    if (aceState !== "active" || !aceActiveSession) {
      return;
    }

    aceState = "ending";
    aceClearActivityTimers();
    aceClearTimer();
    aceRenderMessage("Ending...");

    const endedAt = new Date().toISOString();
    const elapsedMs = aceElapsedMs();
    const documentId = aceActiveSession.documentId || aceExtractDocumentId();
    const wordDiff = await aceGoogleDocDiffAfterSave(
      documentId,
      aceActiveSession.extensionSessionId,
      aceActiveSession.startDocumentWordCount
    );
    const endDocumentWordCount = Number.isFinite(wordDiff.wordCount)
      ? wordDiff.wordCount
      : null;
    const wordsAdded = Math.max(0, Number(wordDiff.wordsAdded) || 0);
    const wordsRemoved = Math.max(0, Number(wordDiff.wordsRemoved) || 0);
    const netWordsChanged = Number.isFinite(wordDiff.netWordsChanged)
      ? wordDiff.netWordsChanged
      : Number.isFinite(aceActiveSession.startDocumentWordCount) && Number.isFinite(endDocumentWordCount)
        ? endDocumentWordCount - aceActiveSession.startDocumentWordCount
        : 0;
    const measuredWordsWritten = Math.max(0, netWordsChanged);
    const measuredWordsEdited = wordsAdded + wordsRemoved;
    const measurementPending = !wordDiff.ok || !Number.isFinite(endDocumentWordCount);
    const wordCountError = measurementPending
      ? wordDiff.error || "Google Docs word count unavailable."
      : "";
    aceCompletedSession = {
      documentId: aceActiveSession.documentId || aceExtractDocumentId(),
      projectId: aceActiveSession.projectId || "",
      sessionType: aceActiveSession.sessionType || "writing",
      startedAt: aceActiveSession.startedAt,
      endedAt,
      durationMinutes: aceDurationMinutes(elapsedMs),
      source: "chrome-extension",
      documentUrl: aceActiveSession.documentUrl || aceDocumentUrl(),
      notes: "",
      extensionSessionId: aceActiveSession.extensionSessionId,
      wordsWritten: aceActiveSession.sessionType === "writing" && !measurementPending ? measuredWordsWritten : 0,
      wordsEdited: aceActiveSession.sessionType === "editing" && !measurementPending ? measuredWordsEdited : 0,
      wordsAdded: measurementPending ? 0 : wordsAdded,
      wordsRemoved: measurementPending ? 0 : wordsRemoved,
      netWordsChanged: measurementPending ? 0 : netWordsChanged,
      startDocumentWordCount: Number.isFinite(aceActiveSession.startDocumentWordCount)
        ? aceActiveSession.startDocumentWordCount
        : null,
      endDocumentWordCount: Number.isFinite(endDocumentWordCount)
        ? endDocumentWordCount
        : null,
      wordCountMethod: "google-docs-api",
      wordCountError,
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
      aceResolveAndSyncCompletedSession(false);
    }
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
    if (!aceActiveSession?.documentId) {
      return;
    }

    try {
      const binding = await aceGetBinding(aceActiveSession.documentId);
      if (binding.project?.id) {
        aceActiveSession.projectId = binding.project.id;
        await acePersistActiveSession();
      }
    } catch (_error) {
      // Binding is resolved again when the session ends.
    }
  }

  async function aceResolveAndSyncCompletedSession(forcePicker) {
    if (!aceCompletedSession) {
      return;
    }

    if (aceCompletedSession.measurementPending) {
      const measured = await aceMeasureCompletedSession();
      if (!measured) {
        return;
      }
    }

    aceSyncStatus = "syncing";
    aceSyncMessage = forcePicker ? "Choose the correct project." : "Syncing...";
    aceRenderCompleted();

    try {
      if (!forcePicker) {
        const binding = await aceGetBinding(aceCompletedSession.documentId);
        if (binding.project?.id) {
          aceCompletedSession.projectId = binding.project.id;
          aceSelectedProject = binding.project;
          await aceSyncCompletedSession();
          return;
        }
      }

      aceProjects = await aceGetProjects();
      aceState = "project-picker";
      aceSyncStatus = "";
      aceSyncMessage = "";
      aceRenderProjectPicker();
    } catch (error) {
      await aceMarkSessionUnsynced(error.message);
    }
  }

  async function aceChooseProject(projectId) {
    if (aceProjectPickerMode === "active") {
      await aceChooseProjectForActiveSession(projectId);
      return;
    }

    if (!aceCompletedSession) {
      return;
    }

    const project = aceProjects.find(function (item) {
      return item.id === projectId;
    });
    if (!project) {
      return;
    }

    aceSelectedProject = project;
    aceCompletedSession.projectId = project.id;
    aceState = "completed";
    aceSyncStatus = "syncing";
    aceSyncMessage = "Saving project...";
    aceRenderCompleted();

    try {
      const binding = await aceSaveBinding(aceCompletedSession.documentId, project.id);
      aceSelectedProject = binding.project || project;
      await aceSyncCompletedSession();
    } catch (error) {
      await aceMarkSessionUnsynced(error.message);
    }
  }

  async function aceChooseProjectForActiveSession(projectId) {
    if (!aceActiveSession) {
      return;
    }

    const project = aceProjects.find(function (item) {
      return item.id === projectId;
    });
    if (!project) {
      return;
    }

    try {
      const binding = await aceSaveBinding(aceActiveSession.documentId, project.id);
      aceActiveSession.projectId = project.id;
      aceSelectedProject = binding.project || project;
      await acePersistActiveSession();
      aceState = "active";
      aceProjectPickerMode = "completed";
      aceStartTimer();
    } catch (error) {
      aceState = "active";
      aceSyncStatus = "pending";
      aceSyncMessage = `Project not changed. ${error.message}`;
      aceProjectPickerMode = "completed";
      aceRenderActive();
    }
  }

  async function aceShowProjectPickerForActiveSession() {
    if (!aceActiveSession) {
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
      aceRenderProjectPicker();
    } catch (error) {
      aceState = "active";
      aceProjectPickerMode = "completed";
      aceStartTimer();
    }
  }

  async function aceSyncCompletedSession() {
    if (!aceCompletedSession?.projectId) {
      await aceResolveAndSyncCompletedSession(true);
      return;
    }

    if (aceCompletedSession.measurementPending) {
      const measured = await aceMeasureCompletedSession();
      if (!measured) {
        return;
      }
    }

    aceSyncStatus = "syncing";
    aceSyncMessage = "Syncing...";
    aceRenderCompleted();

    try {
      const result = await acePostSession(aceCompletedSession);
      aceSelectedProject = result.project || aceSelectedProject;
      aceSyncStatus = "synced";
      aceSyncMessage = result.duplicate ? "Already synced." : "Synced.";
      await aceRemovePendingSession(aceCompletedSession.extensionSessionId);
    } catch (error) {
      await aceMarkSessionUnsynced(error.message);
      return;
    }

    aceRenderCompleted();
  }

  async function aceMeasureCompletedSession() {
    if (!aceCompletedSession?.measurementPending) {
      return true;
    }

    aceSyncStatus = "syncing";
    aceSyncMessage = "Checking Google Docs...";
    aceRenderCompleted();

    const wordDiff = await aceGoogleDocDiffAfterSave(
      aceCompletedSession.documentId,
      aceCompletedSession.extensionSessionId,
      aceCompletedSession.startDocumentWordCount
    );
    if (!wordDiff.ok || !Number.isFinite(wordDiff.wordCount)) {
      await aceMarkSessionUnsynced(`Google Docs count unavailable. ${wordDiff.error || "Try again shortly."}`);
      return false;
    }

    const wordsAdded = Math.max(0, Number(wordDiff.wordsAdded) || 0);
    const wordsRemoved = Math.max(0, Number(wordDiff.wordsRemoved) || 0);
    const netWordsChanged = Number.isFinite(wordDiff.netWordsChanged)
      ? wordDiff.netWordsChanged
      : Number(wordDiff.wordCount) - Number(aceCompletedSession.startDocumentWordCount);

    aceCompletedSession = {
      ...aceCompletedSession,
      wordsWritten: aceCompletedSession.sessionType === "writing" ? Math.max(0, netWordsChanged) : 0,
      wordsEdited: aceCompletedSession.sessionType === "editing" ? wordsAdded + wordsRemoved : 0,
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
    if (ACE_IS_TOP_FRAME) {
      aceSchedulePrompt();
      return;
    }

    window.top.postMessage({ aceType: ACE_ACTIVITY_MESSAGE }, "*");
  }

  function aceSchedulePrompt() {
    if (aceState !== "idle") {
      return;
    }

    if (!aceStartTimerId) {
      aceRenderDetecting();
      aceStartTimerId = window.setTimeout(aceShowStartPrompt, ACE_START_DELAY_MS);
    }

    if (aceActivityGapTimerId) {
      window.clearTimeout(aceActivityGapTimerId);
    }

    aceActivityGapTimerId = window.setTimeout(function () {
      aceClearActivityTimers();
      aceResetPromptState();
      aceRenderIdle();
    }, ACE_ACTIVITY_GAP_MS);
  }

  function aceHandleKeydown(event) {
    if (!aceIsWritingActivity(event)) {
      return;
    }

    aceRegisterWritingActivity();
  }

  function aceHandleInputLikeActivity(event) {
    if (event.inputType && event.inputType.startsWith("format")) {
      return;
    }

    aceRegisterWritingActivity();
  }

  function aceHandleClipboardActivity() {
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
    if (event.button !== 0 || event.target.closest("button")) {
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

  document.addEventListener("beforeinput", aceHandleInputLikeActivity, true);
  document.addEventListener("input", aceHandleInputLikeActivity, true);
  document.addEventListener("keydown", aceHandleKeydown, true);
  document.addEventListener("paste", aceHandleClipboardActivity, true);
  document.addEventListener("cut", aceHandleClipboardActivity, true);

  if (ACE_IS_TOP_FRAME) {
    window.addEventListener("message", function (event) {
      if (event.source === window || !event.data || event.data.aceType !== ACE_ACTIVITY_MESSAGE) {
        return;
      }

      aceSchedulePrompt();
    });

    window.addEventListener("resize", aceApplyWidgetPosition);
    aceWidget.addEventListener("pointerdown", aceHandlePointerDown);
    aceWidget.addEventListener("pointermove", aceHandlePointerMove);
    aceWidget.addEventListener("pointerup", aceHandlePointerUp);
    aceWidget.addEventListener("pointercancel", aceHandlePointerUp);

    aceWidget.addEventListener("click", function (event) {
      const projectButton = event.target.closest("[data-ace-project-id]");
      if (projectButton && aceWidget.contains(projectButton)) {
        aceChooseProject(projectButton.getAttribute("data-ace-project-id"));
        return;
      }

      const button = event.target.closest("[data-ace-action]");
      if (!button || !aceWidget.contains(button)) {
        return;
      }
      if (button.disabled) {
        return;
      }

      const action = button.getAttribute("data-ace-action");

      if (action === "confirm-start") {
        aceStartSession();
      } else if (action === "decline-start") {
        aceDeclineStart();
      } else if (action === "end") {
        aceEndSession();
      } else if (action === "switch") {
        aceSwitchSessionType();
      } else if (action === "open") {
        aceOpenApp();
      } else if (action === "refresh-page") {
        window.location.reload();
      } else if (action === "retry-sync") {
        if (aceCompletedSession?.projectId) {
          aceSyncCompletedSession();
        } else {
          aceResolveAndSyncCompletedSession(false);
        }
      } else if (action === "change-project") {
        if (aceCompletedSession) {
          aceResolveAndSyncCompletedSession(true);
        } else if (aceActiveSession) {
          aceShowProjectPickerForActiveSession();
        }
      } else if (action === "start-new") {
        aceStartNew();
      }
    });

    aceRestoreSession();
  }
})();
