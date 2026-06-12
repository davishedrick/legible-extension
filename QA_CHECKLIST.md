# Scriptor QA Checklist

Use this checklist before release, after high-risk session/state changes, and when a bug cannot be fully covered by the Node regression suite. Record any failures in `BUGS.md`.

## Preparation

- Confirm the extension is loaded in Chrome and available on `https://docs.google.com/document/*`.
- Confirm Scriptor is reachable and the user is signed in.
- Use at least two Google Docs documents, and one document with two manuscript tabs when testing tab-scoped behavior.
- Keep Chrome DevTools open for extension console diagnostics when testing word counts, catch-up, or sync fallback.

## Project Creation

- Create a project from Scriptor and confirm it appears in the extension project picker.
- Create or bind a project from an unbound Google Doc and confirm the selected title/project is retained.
- Refresh the Google Doc and confirm the project remains available through local and remote state.

## Writing Session Start/End

- Start a writing session from an already bound document.
- Confirm the timer starts, the session type is Writing, and the starting word count is captured.
- Add words, end the session, and confirm the completed session shows the net word change.
- Confirm the session appears in Scriptor history and the project word count advances to the ending document word count.
- Bind an empty dedicated Google test doc, start Writing at `0`, paste enough text for a positive Docs API count, and end the session. If the visible counter is delayed or falsely reads `0`, confirm the saved session uses the positive API end count or fails safely, never `Net: 0 words` with a positive API diagnostic.

## Active Timer Minimize

- Start a writing session from an already bound document and click `Minimize`.
- Confirm the full timer collapses to the compact extension surface and shows a calm active-session indicator.
- Keep writing, restore with `Restore timer`, and confirm the elapsed timer is continuous and no second session is created.
- End the restored session and confirm the minimized indicator disappears.
- While minimized, switch to another document or manuscript tab and open the compact surface. Confirm the extension prompts return to the original session tab or end-session handling and does not attach the session to the wrong document.

## Negative Writing Sessions

- Start a Writing session, delete enough text to produce a negative net change, and end the session.
- Confirm the extension displays a negative net count rather than clamping to zero.
- Confirm Scriptor history records the negative net change and the project word count matches the ending document count.

## Document Binding and Starting Word Count

- Bind a non-empty document to a project and confirm the current Google Docs word count becomes the baseline.
- Bind an empty document and confirm a zero baseline is accepted only when the count is verified.
- Bind a non-empty document to a project with no prior sessions and a provisional zero starting count; confirm Scriptor shows the verified document count as both current words and starting baseline rather than `+N` words since tracking began.
- Bind one website-created project and one extension-created project to disposable Google Docs and confirm both persist the same real document/tab identity after refresh.
- Reopen the extension controls for a bound project and confirm it separately lists the project title, bound Google Doc title, and bound document tab title.
- Open the same project in Scriptor and confirm the project card shows the bound document title and tab title, or a clear unavailable binding line for stale/deleted bindings.
- Delete or restrict access to a disposable bound Google Doc, reopen the extension controls, and confirm it shows `Bound document unavailable. Rebind required.` or `Bound document inaccessible. Rebind required.` with no `Start writing` or `Sync document changes` action.
- Confirm a deleted/inaccessible bound document does not overwrite project word count, baseline, or history with stale or empty document data.
- Rebind a project from a stale/deleted document and confirm the new document baseline is reset to the new document count.
- Decline a rebind prompt and confirm no baseline or binding is changed.
- With Project A already bound to Document A, open Document B and try to bind Project A. Confirm the extension/app does not silently overwrite Document A and requires a separate confirmation, return, unbind, or rebind flow.

## Catch-Up Sessions

- With a prior baseline, edit outside a tracked session and trigger manual sync.
- Bind an empty dedicated Google test doc, write outside a tracked session, then reopen controls and start Writing. Confirm Scriptor offers catch-up from `0` to the current word count before starting the timed session.
- Confirm positive deltas become Writing catch-up sessions.
- Confirm negative deltas become Editing catch-up sessions.
- Bind a Google Doc at `1,114` words, click `Sync document changes` without editing, and confirm it reports no change rather than `-1,114`.
- Add a small real edit to the same bound document, trigger `Sync document changes`, and confirm the catch-up delta matches the actual change from the 1,114 baseline.
- Remove that small edit, trigger `Sync document changes`, and confirm the negative catch-up delta matches the real document change only when the document count is verified.
- With the same bound document, confirm unavailable or delayed word-count reads do not overwrite the last measured count with `0`.
- Bind a large real test document around `60,265` words, trigger `Sync document changes` when the visible counter briefly reads `0` but the API reports a changed positive count such as `60,251`, and confirm Scriptor fails safely rather than offering a `-14` catch-up.
- If possible, simulate a failed or unavailable word-count read by temporarily blocking the Docs API/OAuth path, reloading during Docs initialization, or testing while the word-count surface is unavailable. Confirm the extension shows a verification warning and preserves the last known count.
- Confirm no unavailable, missing, or failed read produces a false `-1,114` delta after binding at `1,114`.
- Skip catch-up and confirm the baseline advances so the same prompt does not repeat.
- Confirm catch-up does not appear during active typing or an active tracked session.

## Tab Switching and Wrong Document Protection

- Start a session in Google Docs tab A, switch to manuscript tab B, and confirm the extension blocks completion.
- Start a session on Document A, navigate the same Chrome tab to Document B, and confirm `End Session` stays blocked or requires returning to Document A rather than showing recovery or logging a huge delta.
- Use the return action and confirm the session resumes only on the original manuscript tab.
- Open the same Google Doc in another Chrome tab and confirm the other tab cannot end the original session.
- Switch back to the original tab and confirm completion measures the original document surface.

## Automated Fake-Docs Coverage

- Run `npm run test:e2e` and confirm the fake-doc Playwright suite passes without Google login or a real Google Doc.
- Confirm the fake harness covers extension load, fake document identity, fake word-count reading, first bind baseline correction, catch-up at session boundary, active typing suppression, wrong-document blocking, valid negative writing sessions, impossible wrong-document deltas, and Chrome tab switching.
- Treat failures in fake-doc e2e tests as regression candidates and record reproducible bug classes in `BUGS.md`.

## Optional Real Google Docs Smoke

- Use only a dedicated non-private Google test doc, never a private manuscript.
- Run `npm run test:e2e:google-docs` only with `ENABLE_GOOGLE_DOCS_SMOKE=true` and `GOOGLE_DOCS_SMOKE_URL` set.
- Confirm the extension widget appears on the live Google Doc and exposes a document identity/surface signal.
- Confirm the extension does not show an extension-context crash or reload error after opening the doc.
- Confirm any word-count OAuth prompt or read-only Docs API request is expected and does not require exposing private content.
- Open two dedicated real test docs and manually confirm switching between them does not bind, sync, or end sessions against the wrong document.
- Do not capture, commit, or share screenshots, traces, profiles, cookies, storage state, or private URLs from real-doc testing.

## Editing Session Start/End

- Start an Editing session from the widget.
- Make edits with positive, zero, and negative net word changes across separate runs.
- End each session and confirm the UI and Scriptor history show net word change only.
- Switch between Writing and Editing during a session and confirm the timer is not reset.

## Issue Creation and Resolution

- Select text and create an issue from the extension.
- Confirm the issue includes project, document, manuscript surface, quote/snippet, status, and priority.
- Resolve the issue in Scriptor and confirm the extension no longer treats it as open after refresh.
- Create issues on two manuscript tabs and confirm each issue stays scoped to the correct tab.

## Import/Export

- Import project/manuscript data through the app flow and confirm project state, word counts, sessions, and issues remain coherent.
- Export project data and confirm the export includes current sessions, issues, manuscript completion state, and publishing state.
- Re-import an export into a fresh state when supported and confirm no duplicate sessions are created.

## Manuscript Completion

- Mark a manuscript/project complete.
- Confirm completed projects are visually and behaviorally distinct from active drafts.
- Confirm new sessions, catch-up, issues, and export still behave as intended for completed manuscripts.

## Project Publishing and Reopening

- Publish a completed project and confirm publishing state is persisted.
- Reopen a published project and confirm writing/editing sessions can resume only when the app intends that workflow.
- Confirm project history and export retain the publish/reopen sequence.

## Local/Remote State Persistence Fallback

- Start and end a session while Scriptor is reachable; confirm it syncs and pending local state is cleared.
- Simulate Scriptor/API failure during sync; confirm the completed session remains in local pending state and the widget shows `Not synced yet`.
- Restore connectivity and retry; confirm the pending session syncs once and is removed locally.
- Refresh or close the Google Doc during an active session; confirm abandoned-session recovery appears and can recover or discard without corrupting the baseline.
- If abandoned-session recovery appears with a stored session start count that conflicts with the saved baseline or project count, confirm it shows a warning and does not log a huge false delta such as `+59,159` from `1,081` to `60,240`.
- Confirm local document bindings and baselines are reconciled with remote state after refresh.

## Remaining Manual / Phase 3 Checks

- Confirm real Google Docs OAuth consent and Docs API word-count reads still work with the installed extension.
- Confirm live Google Docs tab IDs, document titles, reload behavior, and deleted/inaccessible documents match the fake-doc assumptions.
- Confirm Scriptor hosted-app UI flows that are outside the extension widget, including import/export, issue resolution, publishing/reopening, and real remote fallback behavior.
- Keep real Google Docs automation optional for Phase 3 and do not make it the default regression layer.
