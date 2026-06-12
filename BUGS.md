# Scriptor Bug Registry

Use this file as the first stop for reproducible bugs and regressions. Keep entries short, factual, and linked to tests or QA notes when possible.

## Bug-Fix Workflow

1. Reproduce the bug in the smallest reliable path.
2. Write a failing regression test when the behavior can be covered without browser automation.
3. Fix the smallest root cause that preserves the current architecture.
4. Confirm the new test and the relevant existing tests pass.
5. Map the bug to `SCENARIOS.md`, or add a new scenario when the behavior is not represented.
6. Add or update a bug entry here, and update `QA_CHECKLIST.md` when the scenario needs manual coverage.

## Entry Template

```md
## BUG-YYYY-MM-DD-short-title

- Status: Open | Fixed | Watching
- Reported: YYYY-MM-DD
- Area: Project state | Writing session | Editing session | Binding | Catch-up | Issues | Import/export | Publishing | Persistence | Extension bridge
- Severity: Low | Medium | High | Critical
- Owner:
- Related files:
- Related tests:
- Related scenarios:

### Summary

What failed, in one or two sentences.

### Reproduction

1. Start from:
2. Do:
3. Expected:
4. Actual:

### Root Cause

What code path or state transition was responsible.

### Fix

What changed and why it is the smallest safe fix.

### Verification

- Automated:
- Manual:

### Follow-Up

Any additional QA, monitoring, docs, or future regression coverage.
```

## Regression Test Structure

- Add pure JavaScript state/session tests to `tests/sessionMeasurement.test.js` or `tests/extensionQaHarness.test.js`.
- Add Google Docs counting/tokenization regressions to `tests/tokenCount.test.js`.
- Use `tests/factories.js` for projects, manuscript surfaces, bindings, baselines, sessions, issues, and snapshots.
- Use `tests/helpers.js` to load `content.js` or `background.js` in a VM with mocked Chrome APIs and storage.
- Add fake Google Docs browser regressions to `tests/e2e/fake-google-doc.spec.js` when behavior depends on real browser tabs, extension loading, or widget UI. Keep real Google Docs behavior in `QA_CHECKLIST.md` until Phase 3 introduces an optional smoke layer.
- Use `SCENARIOS.md` to track whether each critical journey has happy-path, edge-path, and failure-path coverage. Do not mark a scenario automated unless the named test covers the exact behavior.

## Current Test Coverage Snapshot

- `tests/tokenCount.test.js`: Google Docs API extraction, tab-specific counts, tokenizer edge cases, suggested insertion exclusion, deleted/inaccessible document errors.
- `tests/sessionMeasurement.test.js`: session payloads, net word calculation, document binding and baselines, catch-up decisions, tab switching protection, abandoned-session recovery, visible word count fallback, backend compatibility.
- `tests/extensionQaHarness.test.js`: higher-level regression flows for project binding, stale/deleted bindings, catch-up suppression, manual sync, pending sessions, wrong-document protection, negative writing sessions, recovery, and API bridge auth fallback.

## Open Registry

Add new bug entries below this line.

## BUG-2026-06-12-zero-bind-off-session-catchup

- Status: Fixed
- Reported: 2026-06-12
- Area: Catch-up | Binding | Writing session
- Severity: High
- Owner:
- Related files: `content.js`, `tests/extensionQaHarness.test.js`, `BUGS.md`, `SCENARIOS.md`, `QA_CHECKLIST.md`
- Related tests: `opening controls after zero bind and off-session writing shows catch-up`, `starting after zero bind and off-session writing shows catch-up before session`, `manual session start shows positive catch-up before starting`
- Related scenarios: `CU-001`

### Summary

After binding an empty Google Doc at `0`, writing outside a tracked session could fail to surface catch-up when the user reopened extension controls, making the flow look like it was ignoring words written right after the initial bind.

### Reproduction

1. Start from: a new Google Doc bound to a project with a verified baseline of `0`.
2. Do: write about `1,000` words without a tracked session, then reopen extension controls or click `Start writing`.
3. Expected: Scriptor offers a `+1,000` catch-up session before starting the timed session.
4. Actual: reopening controls rendered the normal bound prompt because `show-controls` was not treated as an intentional catch-up boundary.

### Root Cause

Catch-up prompting was limited to `bind`, `pre-session`, and `manual-sync` triggers. The control-open path validated the binding and rendered the prompt directly, so a changed document could appear synced until the narrower start-session check ran.

### Fix

`show-controls` is now an intentional catch-up boundary. Opening controls on a valid bound document runs the existing catch-up candidate builder; if a candidate exists, the catch-up prompt appears. If no baseline/candidate exists, the normal bound prompt is preserved.

### Verification

- Automated: `node --test tests/extensionQaHarness.test.js` covers zero-bind/off-session catch-up on control open and explicit session start, plus existing deleted/inaccessible binding cleanup.
- Manual: `QA_CHECKLIST.md` includes the empty-bind, off-session-writing, reopen/start catch-up check.

## BUG-2026-06-11-deleted-doc-binding-stays-active

- Status: Fixed
- Reported: 2026-06-11
- Area: Binding | Persistence | Extension bridge
- Severity: High
- Owner:
- Related files: `content.js`, `tests/extensionQaHarness.test.js`, `tests/e2e/fake-google-doc.spec.js`, `BUGS.md`, `SCENARIOS.md`, `QA_CHECKLIST.md`
- Related tests: `prompt lookup invalidates deleted bound document before showing bound controls`, `prompt lookup marks forbidden bound document unavailable instead of staying valid`, `project picker treats bound row without document identity as rebind required`, `existing valid bound project remains available after prompt validation`, `website-created and extension-created projects use the same local binding schema`, `bound prompt lists project document and tab titles`, `deleted bound document reopens as rebind required instead of active`
- Related scenarios: `DB-007`, `PS-005`

### Summary

A project whose bound Google Doc was deleted or inaccessible could still appear actively bound in the main extension prompt because normal prompt lookup trusted backend/local binding metadata without validating the Google Docs file. A partial active binding row could also appear bound in the project picker without enough document identity to validate or rebind clearly.

### Reproduction

1. Start from: a project bound to a Google Docs document/tab, with binding metadata in Scriptor and `aceDocumentBindings`.
2. Do: delete the Google Doc or make it inaccessible, then reopen the extension controls on that document URL.
3. Expected: the project is unbound or marked unavailable, with a clear rebind-required prompt and no document sync/start actions.
4. Actual: the prompt could show the project as bound with `Start writing` and `Sync document changes` because the binding row still existed.

### Root Cause

Deleted/403/missing-tab detection existed in project-picker reconciliation, but `aceRefreshCurrentBinding` used `aceGetBoundProjectForDocument`, which reconciled only local/server binding records. It did not call the Google Docs word-count/API validation path before allowing the binding to drive the main prompt. When the server confirmed a binding by project only, the extension could also overwrite cached document-title metadata with an empty current-surface title during page load.

### Fix

Prompt binding refresh now validates an apparently bound current document through the Google Docs word-count/API path. A `404`, `403`, or missing tab response removes the local binding, patches the server binding status to stale/unavailable, clears `aceCurrentBinding`, and renders the project as unbound with a rebind-required message. Project-picker rows that claim to be bound but lack document identity are treated as `stale_missing_doc` instead of active. Valid bindings preserve cached document/tab titles, and the prompt displays Project, Document, and Tab readouts.

### Verification

- Automated: `tests/extensionQaHarness.test.js` / prompt invalidation, forbidden invalidation, partial binding identity invalidation, valid binding preservation, normalized schema; `tests/e2e/fake-google-doc.spec.js` / bound prompt lists project document and tab titles, deleted bound document reopens as rebind required.
- Manual: `QA_CHECKLIST.md` includes deleted/inaccessible bound-document checks using only dedicated test docs.

### Follow-Up

Confirm the hosted Scriptor API treats stale binding status consistently with the extension. Do not commit real deleted-document URLs, auth state, traces, screenshots, storage, cookies, or profiles.

## BUG-2026-06-10-empty-start-visible-zero-session-end

- Status: Fixed
- Reported: 2026-06-10
- Area: Writing session | Extension bridge
- Severity: High
- Owner:
- Related files: `content.js`, `tests/sessionMeasurement.test.js`, `tests/e2e/fake-google-doc.spec.js`, `BUGS.md`, `SCENARIOS.md`, `QA_CHECKLIST.md`
- Related tests: `session end uses API when empty-start session has positive API and false visible zero`, `session end still saves zero when empty document is verified by API and visible count`, `empty-start writing session uses positive API count when visible counter falsely stays zero`
- Related scenarios: `WS-001`, `PS-004`

### Summary

A writing session started from an empty Google Doc could save `Net: 0 words` after pasted text when the Google Docs API reported a positive ending count but the visible counter reader returned a false stable `0`.

### Reproduction

1. Start from: a bound Google Doc with verified start count `0`.
2. Do: start a Writing session, paste text so the Docs API reports about `2,024` words, force or encounter a visible counter read of `0`, then end the session.
3. Expected: the session logs about `+2,024` words, or at minimum does not sync a trusted zero.
4. Actual: the session synced with `Net: 0 words` and diagnostic `W-API-VISIBLE-MISMATCH: start 0; API end 2024; visible end 0; net 0; end source stable-visible`.

### Root Cause

`aceGoogleDocNetAfterSave` treated stable visible `0` as suspicious only when the session started from a positive word count. Empty-start sessions skipped that guard, so a false stable visible `0` won over a positive API count and the mismatch was recorded only as a warning.

### Fix

Session end now waits for the API whenever the stable visible result is `0`. If the session started at `0` and the API returns a positive ending count, the resolver rejects the visible zero and uses the session API count as the synced end count. Verified empty documents still save `0` when API and visible both report zero.

### Verification

- Automated: `tests/sessionMeasurement.test.js` / `session end uses API when empty-start session has positive API and false visible zero`; `session end still saves zero when empty document is verified by API and visible count`; `tests/e2e/fake-google-doc.spec.js` / `empty-start writing session uses positive API count when visible counter falsely stays zero`.
- Manual: `QA_CHECKLIST.md` includes the empty-doc paste/end-session check with a false or delayed visible counter.

### Follow-Up

On real Google Docs, use only a dedicated test document. Do not commit real document URLs, screenshots, traces, auth state, browser profiles, cookies, storage, or private manuscript content.

## BUG-2026-06-11-provisional-zero-bind-baseline

- Status: Fixed
- Reported: 2026-06-11
- Area: Binding | Project baseline | Activity
- Severity: Medium
- Owner:
- Related files: `content.js`, `tests/e2e/fake-google-doc.spec.js`, `tests/fixtures/fake-google-doc.html`, `BUGS.md`, `SCENARIOS.md`, `QA_CHECKLIST.md`
- Related tests: `verified bind replaces provisional zero project baseline`
- Related scenarios: `DB-009`

### Summary

A project could show `Current words: 320` and `Since tracking began: +320` after binding to a non-empty Google Doc, while the first timed writing session honestly saved `Net: 0 words`.

### Reproduction

1. Start from: a new project whose project state contains a provisional `startingWordCount` of `0`.
2. Do: bind it to a Google Doc whose verified current count is positive, such as `320`, then start and end a session without additional manuscript changes.
3. Expected: the verified bind count becomes both the project current count and the project starting baseline; the first unchanged session may save zero net, but the baseline words are not presented as new tracking progress.
4. Actual: the project current count became `320`, but the project starting count stayed at provisional `0`, making the project dashboard imply `+320` since tracking began while recent activity showed a zero-net session.

### Root Cause

Binding code treated any established `startingWordCount` as authoritative, including provisional or migration-created zero baselines on projects with no sessions. That caused the web project baseline and the extension document baseline to diverge.

### Fix

Verified binding now replaces a provisional zero project baseline when there are no sessions and the Google Docs count is positive. Existing real baselines and projects with sessions are preserved.

### Verification

- Automated: `tests/e2e/fake-google-doc.spec.js` / `verified bind replaces provisional zero project baseline`.
- Manual: `QA_CHECKLIST.md` includes binding a non-empty document to a fresh/provisional-zero project and confirming both current and starting project counts match the verified document count.

## BUG-2026-06-10-false-visible-zero-api-delta-catchup

- Status: Fixed
- Reported: 2026-06-10
- Area: Binding | Catch-up | Extension bridge
- Severity: High
- Owner:
- Related files: `content.js`, `tests/sessionMeasurement.test.js`, `tests/e2e/fake-google-doc.spec.js`, `BUGS.md`, `SCENARIOS.md`, `QA_CHECKLIST.md`
- Related tests: `catch-up check suppresses false visible zero when API returns changed positive count`, `manual sync suppresses catch-up when false visible zero conflicts with changed API count`
- Related scenarios: `DB-005`, `PS-003`, `PS-004`

### Summary

Opening a bound Google Doc after a while could show a catch-up prompt for a small negative delta, such as `60,265 -> 60,251`, even while the visible Google Docs counter still showed `60,265`.

### Reproduction

1. Start from: a bound document baseline of `60,265` words.
2. Do: trigger catch-up/manual sync when the visible count reader returns a false stable `0`, while the Docs API returns a positive but changed count such as `60,251`.
3. Expected: the extension treats this as diagnostic-only and preserves the baseline until the visible/current count confirms the change.
4. Actual: the extension trusted the changed positive API count and offered a `-14` catch-up editing session.

### Root Cause

`aceGoogleDocWordCountAfterSettle` treated any stable visible `0` plus positive trusted API count as a valid API replacement during catch-up. That was safe for first-bind baseline recovery and no-change API confirmation, but unsafe when the API count differed from the saved baseline because the visible counter had failed to confirm the changed count.

### Fix

For catch-up checks, a false visible zero plus changed positive API count is now diagnostic-only. First-bind baseline can still use the positive API count, and unchanged positive API confirmation still results in no catch-up.

### Verification

- Automated: `tests/sessionMeasurement.test.js` / `catch-up check suppresses false visible zero when API returns changed positive count`; `tests/e2e/fake-google-doc.spec.js` / `manual sync suppresses catch-up when false visible zero conflicts with changed API count`.
- Commands run: exact VM reproduction for `60,265 -> API 60,251 / visible 0` passed; focused Node and fake-doc Playwright tests passed.
- Manual: `QA_CHECKLIST.md` includes the large-document false-visible-zero changed-API check.

### Follow-Up

If live Google Docs still reports a changed API count while the visible counter remains unchanged, capture only diagnostics and sanitized counts. Do not commit real document URLs, screenshots, traces, auth state, storage, cookies, or browser profiles.

## BUG-2026-06-07-recovery-start-baseline-mismatch

- Status: Fixed
- Reported: 2026-06-07
- Area: Writing session | Persistence | Extension bridge
- Severity: High
- Owner:
- Related files: `content.js`, `tests/extensionQaHarness.test.js`, `BUGS.md`, `SCENARIOS.md`, `QA_CHECKLIST.md`
- Related tests: `recovery blocks stale start count that conflicts with saved baseline`
- Related scenarios: `WS-005`, `WS-006`, `PS-003`, `PS-004`

### Summary

An abandoned-session recovery could combine a stale stored session start count with the current Google Docs word count and offer a huge false writing-session delta.

### Reproduction

1. Start from: a project whose saved baseline/current count is 63,000 words.
2. Do: restore an abandoned active session whose stored `startDocumentWordCount` is stale at 1,081 words, while the current bound Google Doc count is 60,240 words.
3. Expected: recovery detects the stale start snapshot conflict and fails safely without logging a session.
4. Actual: recovery could calculate `60,240 - 1,081 = +59,159` and allow that false session to be recovered.

### Root Cause

`aceBuildRecoveryCandidate` trusted `aceActiveSession.startDocumentWordCount` during abandoned-session recovery without checking it against the saved document baseline or the project current word count. A stale stored active session could therefore be paired with the current document count from the correct document and still produce an impossible net change.

### Fix

Recovery now compares the stored session start count with the saved document baseline, falling back to the project current word count when no baseline is available. If the values conflict beyond a conservative tolerance, recovery produces a diagnostic `E-RECOVERY-START-BASELINE-MISMATCH`, keeps measurement pending, shows a warning, and blocks the Recover action from syncing the bogus session. Normal recovery still works when the stored start count agrees with the saved baseline.

### Verification

- Automated: `tests/extensionQaHarness.test.js` / `recovery blocks stale start count that conflicts with saved baseline`.
- Commands run: `node --test tests/extensionQaHarness.test.js` passed; `npm test` passed; `npm run test:e2e` passed.
- Manual: `QA_CHECKLIST.md` includes abandoned-session recovery checks for baseline/start-count mismatches.

### Follow-Up

On live Google Docs, use a dedicated test document to confirm stale recovery state fails safe and does not create large deltas. Do not commit real-doc screenshots, URLs, auth state, traces, videos, or browser profiles.

## BUG-2026-06-05-wrong-document-end-recovery

- Status: Fixed
- Reported: 2026-06-05
- Area: Writing session | Persistence | Extension bridge
- Severity: High
- Owner:
- Related files: `content.js`, `tests/e2e/fake-google-doc.spec.js`, `BUGS.md`, `SCENARIOS.md`, `QA_CHECKLIST.md`
- Related tests: `impossible wrong-document delta is blocked`, `active session remains tied to Document A when page identity switches`, `tab switching does not let Document B finalize Document A session`
- Related scenarios: `WS-003`, `WS-004`, `PS-004`

### Summary

The fake-doc Playwright suite exposed a timing path where ending a Document A session from Document B could navigate back to Document A and fall into abandoned-session recovery instead of staying blocked.

### Reproduction

1. Start from: a writing session started on Document A at 60,500 words.
2. Do: switch the same page context to Document B at 500 words, wait for the widget to enter the tab/document-blocked state, then click `End Session`.
3. Expected: ending remains blocked or requires returning to Document A; no huge wrong-document delta is logged or shown.
4. Actual: the blocked-state end path could navigate back to Document A, unload the page, and show abandoned-session recovery with a bogus large negative recovery delta.

### Root Cause

`aceEndBlockedSession` allowed the same Chrome tab to navigate back to the original session document and then continue ending. For a different document, that navigation unloads the current page before the normal end flow can safely measure the original surface, so recovery could measure an unrelated/default surface count.

### Fix

When the active session document ID and current document ID differ in the same Chrome tab, the blocked end action now stays blocked and tells the user to return to the original document before ending. Same-document manuscript-tab returns can still use the existing return-and-end path.

### Verification

- Automated: `npm run test:e2e` includes `impossible wrong-document delta is blocked`; `npm test` covers the lower-level session/document identity guardrails.
- Manual: `QA_CHECKLIST.md` includes wrong-document end-session protection for same-tab and separate-tab flows.

### Follow-Up

Keep real Google Docs checks manual/optional for this path. Do not turn the real-doc smoke layer into the default regression suite.

## BUG-2026-06-05-manual-sync-visible-zero

- Status: Fixed
- Reported: 2026-06-05
- Area: Binding | Catch-up | Extension bridge
- Severity: High
- Owner:
- Related files: `content.js`, `tests/sessionMeasurement.test.js`, `tests/e2e/fake-google-doc.spec.js`, `tests/fixtures/fake-google-doc.html`, `QA_CHECKLIST.md`
- Related tests: `catch-up decision does not coerce unknown current count to zero`, `catch-up check rejects false visible zero when API returns unchanged positive count`, `catch-up check rejects false visible zero when API returns changed positive count`, `catch-up check fails safely when visible zero cannot be verified`, `manual sync after binding at 1,114 reports no change when the document is unchanged`, `manual sync rejects a false visible zero when the bound document API still reports 1,114`, `manual sync rejects false visible zero when API reports a changed positive count`, `manual sync fails safely when the word-count read fails after binding`, `manual sync does not coerce missing word count to zero after binding`, `manual sync accepts a verified zero on the same bound document`, `manual sync from Document B does not apply a delta to bound Document A`
- Related scenarios: `DB-003`, `DB-005`, `DB-006`, `PS-003`, `PS-004`

### Summary

After binding a Google Doc at 1,114 or 60,284 words, `Sync document changes` could treat the current visible count as 0 and offer a false large negative catch-up even though the Docs API still reported a positive count.

### Reproduction

1. Start from: a Google Doc bound to a Scriptor project with a verified baseline of 1,114 or 60,284 words.
2. Do: click `Sync document changes` without changing the document.
3. Expected: Scriptor reports no change, or a small API-confirmed delta such as `60,284 -> 60,276`, and preserves the known baseline until the user logs catch-up.
4. Actual: the extension could accept a transient visible `0` as the current count and show a `-1,114` or `-60,284` catch-up.

### Root Cause

The binding flow used `aceGoogleDocWordCountAfterSettle` with `verifyZeroWithApi`, so visible zero was rejected unless the Google Docs API also confirmed zero. The manual sync/catch-up path called the same helper without that zero-verification guard, allowing a stable but false visible `0` to override a known-good API count or become a synthetic negative delta.

### Fix

Manual sync and catch-up now verify visible zero with the trusted current document API before accepting it. Unknown, failed, missing, or untrusted current counts remain diagnostic-only and no longer become `0` in catch-up candidate evaluation. The fake-doc harness can now simulate false visible zero, positive API counts that are unchanged or changed from baseline, missing word count, read failure, verified zero, and document mismatch.

### Verification

- Automated: `npm test` and `npm run test:e2e` cover unchanged 1,114 sync, false visible zero with API 1,114, false visible zero with API 60,276 after a 60,284 baseline, failed/missing counts, verified real zero, and sync from a different document.
- Commands run: `npm test` passed; `npm run test:e2e` passed.
- Manual: `QA_CHECKLIST.md` includes the real Google Docs bind-at-1,114 manual sync check.

### Follow-Up

Real Google Docs smoke tests remain optional. If this appears again on a live doc and the diagnostic still says `D-CATCHUP-STABLE-VISIBLE-API-MISMATCH` with `visible end 0` and a positive API end count, confirm Chrome has reloaded the current unpacked extension build before debugging product code further. Capture diagnostics without committing screenshots, private URLs, auth state, traces, videos, or browser profiles.

## BUG-2026-06-04-extension-document-session-identity

- Status: Watching
- Reported: 2026-06-04
- Area: Binding | Writing session | Catch-up | Persistence | Extension bridge
- Severity: High
- Owner:
- Related files: `content.js`, `background.js`, `tests/extensionQaHarness.test.js`, `tests/sessionMeasurement.test.js`, `tests/tokenCount.test.js`, `tests/e2e/fake-google-doc.spec.js`, `tests/fixtures/fake-google-doc.html`, `../Author-companion/writing_app/extension_bridge.py`, `../Author-companion/writing_app/tests/test_backend_contracts.py`
- Related tests: `initial bind with existing text establishes baseline without catch-up`, `manual session start shows positive catch-up before starting`, `typing suppression blocks ambient catch-up but explicit start still reconciles`, `different Google document blocks session completion before word count`, `same document and same Chrome tab can complete a negative writing session`, `active project binding mismatch requires confirmation instead of overwriting`, fake-doc Playwright tests for extension loading, word-count detection, first bind, catch-up, active typing suppression, wrong-document protection, negative sessions, impossible deltas, and tab switching
- Related scenarios: `DB-001`, `DB-002`, `DB-004`, `DB-006`, `WS-001`, `WS-002`, `WS-003`, `WS-004`, `WS-005`, `CU-001`, `CU-002`, `CU-003`, `CU-004`, `PS-001`, `PS-002`

### Summary

Extension-related regressions tend to happen when document identity, manuscript tab identity, Google Docs word counts, session scope, binding state, or catch-up baselines drift apart.

### Reproduction

1. Start from a bound Google Doc or an active extension session.
2. Switch documents, switch manuscript tabs, change word count outside a tracked session, or attempt to bind the same project from another document.
3. Expected: the extension distinguishes first bind, catch-up, valid negative deltas, and wrong-document context.
4. Actual: if this risk regresses, Scriptor could log words against the wrong document, overwrite a binding, or convert a wrong-document count into a huge false delta.

### Root Cause

No active defect is open. This is a tracked bug class for state/session identity regressions.

### Fix

Phase 1 identified existing pure/test-exported logic and added the missing bound-document mismatch coverage. Phase 2 added a narrow fake-doc browser harness that is gated to `docs.google.com` URLs with `scriptorFakeDocs=1` and an explicit test DOM marker, so production Google Docs detection is not weakened.

### Verification

- Automated: Node VM tests cover extension identity, binding, word-count, catch-up, negative delta, and wrong-document protection paths. Fake-doc Playwright tests cover browser extension loading, widget/project binding flows, tab switching, and wrong-document completion attempts without real Google Docs. Backend pytest coverage already verifies first verified binding establishes the project baseline without logging a session.
- Manual: Real Chrome/Google Docs OAuth and live Docs API flows remain in `QA_CHECKLIST.md`.

### Follow-Up

Phase 3 added an optional real-Google-Docs smoke layer for extension loading and document identity signals only. It is excluded from default test commands and must not require personal manuscripts, login automation, or committed auth artifacts.
