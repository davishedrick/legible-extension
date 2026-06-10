# Scriptor Chrome Extension

A minimal Chrome extension companion for Scriptor. It runs only in Google Docs, watches for writing activity, and asks before starting a local stopwatch session.

Hosted app:
https://davishedrick.pythonanywhere.com

## Folder structure

This folder is the extension folder to load in Chrome:

- `manifest.json`
- `background.js`
- `content.js`
- `style.css`
- `README.md`

## What it does

- Shows a small draggable `Idle` widget in Google Docs.
- Snaps to six window positions: top-left, top-right, middle-left, middle-right, bottom-left, and bottom-right.
- Detects writing-like activity without sending document contents to Scriptor.
- Auto-starts a visible session on the first edit in Google Docs already connected to a project.
- For unconnected docs, starts tracking only after clicking `Yes`.
- Starts new sessions as the last-used type, defaulting to Writing on first use.
- Lets you switch between Writing and Editing without restarting the timer.
- Measures session net word change from Google Docs start/end word counts.
- Binds the current manuscript surface to one Scriptor project before syncing. If Google Docs tab identity is unavailable, it falls back to the parent Google Doc.
- Stores active and unsynced completed sessions in `chrome.storage.local`.
- Does not automatically open Scriptor while you are writing.
- Keeps widget position in page `sessionStorage`.

## Install locally

1. Go to `chrome://extensions`.
2. Turn on `Developer mode`.
3. Click `Load unpacked`.
4. Select this folder.
5. Open a Google Doc.
6. Start typing for about 4 seconds.
7. Confirm the widget asks `Start session?`.
8. Click `Yes` and confirm the stopwatch starts.

## Google Docs word-count setup

Before/after word counting uses the Google Docs API through Chrome Identity. Sessions require this OAuth setup so the extension can avoid typing estimates.

1. In Google Cloud Console, enable the Google Docs API.
2. Create an OAuth client for a Chrome extension.
3. Use this unpacked extension's Chrome extension ID for the OAuth client.
4. Replace `REPLACE_WITH_GOOGLE_OAUTH_CLIENT_ID.apps.googleusercontent.com` in `manifest.json`.
5. Reload the extension in `chrome://extensions`.
6. Start a session in Google Docs and approve the read-only Google Docs permission if Chrome asks.

The extension requests `https://www.googleapis.com/auth/documents.readonly` only to compute start/end word counts. It sends net word change, final count, timing, project ID, manuscript surface metadata, and session metadata to Scriptor, not the document text.

## Automated tests

Install JavaScript test dependencies:

```sh
npm install
```

Run the pure extension logic suite:

```sh
npm test
```

Run the fake Google Docs browser suite:

```sh
npx playwright install chromium
npm run test:e2e
```

The Playwright suite loads this unpacked extension into Chromium and routes `https://docs.google.com/document/d/*` to `tests/fixtures/fake-google-doc.html`. The fake page simulates only the extension-facing pieces Scriptor needs: document ID, document title, tab ID, current word count, editable document activity, document switching, word-count changes, and multi-tab browser state.

These tests cover fake-doc extension loading, document identity detection, word-count reads, first bind baseline correction, catch-up at intentional session boundaries, active typing suppression, wrong-document session protection, valid negative writing sessions, impossible wrong-document deltas, and tab switching. They intentionally do not automate Google login, real Google Docs, real manuscript content, or Google Docs DOM internals.

On macOS, Chromium may need normal access to browser support/profile paths. If a sandbox blocks `~/Library/Application Support/Google/Chrome for Testing`, rerun `npm run test:e2e` outside that sandbox.

Run the optional real Google Docs smoke tests only when you intentionally want a live-doc check:

```sh
ENABLE_GOOGLE_DOCS_SMOKE=true \
GOOGLE_DOCS_SMOKE_URL="https://docs.google.com/document/d/DEDICATED_TEST_DOC_ID/edit" \
npm run test:e2e:google-docs
```

Use a dedicated throwaway Google test doc, preferably one that does not require account chooser automation. Do not use private manuscripts. Do not commit real document URLs, storage state, cookies, screenshots, traces, or browser profiles.

The real-doc smoke layer verifies only that the extension loads on a live Google Docs document and exposes an extension-owned document identity/surface signal. It does not automate Google login, bind projects, start or end sessions, assert exact word counts, cover the Phase 2 A-I scenarios, or replace fake-doc regression tests.

## Test checklist

- Open `https://docs.google.com/document/*` and confirm the small `Idle` widget appears.
- Drag the widget and release it near another part of the window; confirm it snaps into place.
- Type normal text for about 4 seconds and confirm the widget asks `Start session?`.
- Click `No` and confirm it returns to `Idle`.
- Type again, click `Yes`, and confirm the widget changes to `Writing 00:00`.
- Type during a Writing session and confirm the timer continues without showing live word counts.
- End the session and confirm the completed widget says `Words measured from Google Docs` when OAuth is configured.
- Switch to Editing and confirm the completed widget shows net word change only.
- Confirm Scriptor does not open automatically.
- Click `Switch to Editing` and confirm the label changes without resetting the timer.
- Click `End` and confirm the completed state shows the session type, tracked minutes, and net word change.
- If prompted, choose the correct Scriptor project and confirm the session syncs.
- Click `Open app` and confirm Scriptor opens at `https://davishedrick.pythonanywhere.com`.
- Start a session, reload the Google Docs tab, and confirm the timer/session type restore.
- Simulate an API failure and confirm the completed session shows `Not synced yet` with a retry option.

## API bridge checklist

- Bind Google Doc A to Project A, end a writing session, and confirm Project A receives it.
- If the Google Doc exposes a tab ID, bind two tabs in the same doc to different projects and confirm each session stays with its current manuscript.
- Confirm Project A History shows the synced net word change instead of `0`.
- Bind Google Doc B to Project B, end a session, and confirm Project B receives it.
- Switch active project in Scriptor, then end a session from Doc A and confirm it still goes to Project A.
- Reload the Google Doc mid-session, end it, and confirm it still syncs correctly.
- Simulate API failure, confirm the session is retained as unsynced, then retry after the app is reachable.

## Deleted binding QA checklist

- Test A: Bind Project A to Google Doc 1, delete or remove access to Google Doc 1, then open/refresh the extension. Expected: Project A is no longer actively bound, has a deleted/stale binding marker, Google Doc 1 is not treated as active, and the hosted app state is updated.
- Test B: With Project A marked as previously bound to deleted Google Doc 1, open Google Doc 2 and attempt to bind Project A. Expected: the prompt says `This project was bound to a now-deleted file. Update this project to your current tab?`; clicking `Yes` binds Project A to Google Doc 2, clears the deleted marker, updates the hosted app, saves the current word count as the new baseline, and shows no catch-up prompt.
- Test C: Repeat Test B and click `No`. Expected: Project A remains unbound, Google Doc 2 is not sent as the active binding, no baseline is saved for Google Doc 2, and no session or catch-up prompt is created.
- Test D: Bind Project A to accessible Google Doc 1 and refresh. Expected: Project A remains bound, no deleted marker appears, and no rebind prompt appears.
- Test E: Bind Project A to accessible Google Doc 1, then open Google Doc 2 and attempt to bind Project A. Expected: the extension does not silently overwrite the active binding; canceling leaves Project A bound to Google Doc 1.
- Test F: Rebind Project A from deleted Google Doc 1 with old baseline `1000` to Google Doc 2 with `397` words. Expected: the new baseline is `397`, no `+/-603` catch-up appears, and future catch-up compares against `397`.
- Test G: After every bind, unbind, or rebind, confirm active binding and deleted binding are not both set for the same deleted doc, unbound projects are not treated as bound, and the current tab is treated as bound only after confirming `Yes`.

## Catch-up QA checklist

- Test A: Bind a project to a document with `397` words and no prior baseline. Expected: catch-up appears immediately during bind, not later; log or skip advances the baseline to `397`.
- Test B: With baseline `0` and document count `500`, click `Start writing` or `Start editing`. Expected: catch-up appears before the tracked session starts; after log/skip, the session starts from `500`.
- Test C: Type continuously outside a tracked session. Expected: no catch-up prompt appears during active typing.
- Test D: Start a tracked session and keep writing. Expected: no catch-up prompt appears while the timer is running.
- Test E: Skip catch-up. Expected: baseline updates to the current count and the same prompt does not reappear.
- Test F: Click `Sync document changes` after editing outside a session. Expected: reconciliation runs on demand and baseline updates after log/skip.
- Test G: Keep typing within the 15 second suppression window. Expected: no automatic catch-up prompt appears.
- Test H: Confirm delta direction: baseline `100`, current `600` shows `+500`; baseline `600`, current `100` shows `-500`.

## Notes

This extension is intentionally a passive trigger layer. When Google OAuth is configured, it reads the Google Doc through the official read-only Docs API to calculate before/after word counts, then syncs only session metadata to Scriptor:

> You're already writing. I'll track it if you want.
