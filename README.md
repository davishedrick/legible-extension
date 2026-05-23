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

## Notes

This extension is intentionally a passive trigger layer. When Google OAuth is configured, it reads the Google Doc through the official read-only Docs API to calculate before/after word counts, then syncs only session metadata to Scriptor:

> You're already writing. I'll track it if you want.
