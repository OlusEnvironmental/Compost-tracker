# Compost Bay Tracker — Design Notes

Living source-of-truth documentation for this system, part of the Olus Management System consolidation project. Read this first at the start of any session touching this repo; update it after any material change.

Repo: `olusenvironmental/Compost-tracker` (public). Single file: `compost_tracker_onedrive.html` (~1,120 lines — markup, CSS and JS all in one page). No build step, no backend of its own — it talks directly to Microsoft Graph from the browser.

**Status: this system holds live production data in daily use.** Unlike General Management and Machine Management, findings here should not be batch-fixed on a general go-ahead — each change needs to be considered individually for its effect on data already in flight (bays mid-fill/turn/empty) before touching anything.

---

## What this app is

A single-page app for tracking compost bay lifecycle at the Olus site: 8 bays (A–H) cycling through fill → mature → turn → post-turn maturation → empty, plus a log of every event and some basic CSV/JSON export/analysis. It's designed to be used on a phone or tablet in the yard as well as at a desk.

There is no server component. The page loads directly from GitHub Pages (`https://olusenvironmental.github.io/Compost-tracker/compost_tracker_onedrive.html`), authenticates the user against Entra ID via MSAL.js, and reads/writes a single JSON file on a SharePoint site's document library through the Microsoft Graph API. The app and its own code/UI label this "OneDrive" throughout (variable/function names, sync-status text) — it is actually a SharePoint site drive (`sites/olusenvironmental.sharepoint.com/drive`), not a personal OneDrive. Harmless naming mismatch, not a functional bug, but worth knowing so nobody goes looking for the file in the wrong place.

## Auth & sync mechanism

- **MSAL config** (`MSAL_CONFIG`, line ~352): client ID `cdfb14bf-...`, tenant `6352d712-...`, redirect URI is the GitHub Pages URL itself, cache in `localStorage`.
- **Scopes**: `Files.ReadWrite`, `Sites.ReadWrite.All`, `User.Read`.
- **Target file**: `sites/olusenvironmental.sharepoint.com/drive/root:/Documents/Compost Bay Tracker/compost-tracker-data.json`.
- On load, `initAuth()` tries `handleRedirectPromise()` then falls back to any cached MSAL account; if signed in, `onSignedIn()` calls `loadFromOneDrive()` and overwrites the local `data` object with whatever comes back, then re-saves that into `localStorage`.
- `getToken()` tries `acquireTokenSilent`, and on failure falls back to `acquireTokenPopup` — a popup the user has to interact with.
- Every `save()` call (invoked after every single action — add bay, start/complete fill, start/complete turn, start/complete empty, delete bay) writes to `localStorage` synchronously, then fires `saveToOneDrive()` asynchronously (not awaited by any caller).
- There is no periodic re-sync and no `visibilitychange`/focus listener — the in-memory `data` object is loaded once at sign-in and never refreshed from the server again until the page is reloaded.
- Local fallback: if there's no token (offline, or auth failed), the app runs entirely off whatever is in `localStorage` (key `compost_v4`) and just marks the sync status "offline".

## Data model

```
data = {
  bays: [
    {
      id, name,
      filledDate, filledComplete,
      turnedDate, turnedComplete,
      emptyStartDate, emptyComplete,   // legacy fields, not actively used by current flow
      notes, batchNumber,
      inProgress: { type: 'filling'|'turning'|'turning-dest'|'emptying', since, notes, targetId?, sourceId? } | null
    },
    ...
  ],
  log: [ { id, date, bayId, bayName, type, notes }, ... ]
}
```

Status is never stored — it's derived on every render by `getStatus(b)` from the fields above: `empty | filling | maturing | turn-due | turning | turning-dest (rendering only — see Known issues) | post-turn | ready | emptying`. Maturation timer is 28 days from `filledDate`; post-turn timer is 21 days from `turnedDate`.

## Bay lifecycle (the five action pairs)

Each lifecycle stage is a "start X" / "confirm X complete" pair, each opening a modal, each ending in `save()` (localStorage write + fire-and-forget SharePoint push) and `renderAll()`:

1. **Fill**: `startFill` sets `inProgress:{type:'filling',...}` and the provisional `filledDate`; `completeFill` clears `inProgress` and sets `filledComplete:true`. The 28-day timer runs from `filledDate`, not from completion.
2. **Turn**: `startTurn` requires an empty bay to turn into (source and destination are different bays). It sets `inProgress:{type:'turning',targetId}` on the source and `inProgress:{type:'turning-dest',sourceId}` on the destination — both bays are locked while a turn is in progress. `confirmCompleteTurn` copies the source's fill/notes/batch data onto the destination, sets the destination's `turnedDate`/`turnedComplete`, and fully resets the source bay to empty.
3. **Empty**: `startEmpty` / `confirmCompleteEmpty` — clears the bay back to its empty state, keeping only the log entry as a record.
4. **Delete**: `confirmDelete` removes a bay from `data.bays` entirely (log entries referencing it are kept). See Known issues — this has no guard against deleting a bay that is mid-lifecycle.
5. **Add**: `confirmAddBay` appends a new bay object.

Batch numbers are auto-suggested (`B001`, `B002`, ...) by counting existing `fill-start` log entries, and carried through fill → turn onto the destination bay.

## Export / import

- `exportStatusCSV` / `exportLogCSV` — CSV snapshot of current bay state / full event log, browser download.
- `exportJSON` — full `data` object as a downloadable JSON backup.
- `restoreJSON` — reads a user-selected JSON file, checks only that it has `bays` and `log` keys present, asks a single `confirm()` ("Replace all current data?"), then overwrites `data` wholesale and calls `save()` — which immediately pushes the restored data to the shared SharePoint file. See Known issues.

## Known issues found during Phase 1 audit

Severity here means "how bad if it happens" combined with "how likely in normal day-to-day use" — flagged this way specifically because this system holds live data, per your instruction to treat any fix here with more caution than the other two systems. **Nothing below has been changed yet — this is the audit, not a fix pass.**

### HIGH

- **A failed or 404'd load from SharePoint silently wipes local data to empty.** `loadFromOneDrive()` treats a 404 (file genuinely not there yet) as "return `{bays:[],log:[]}`" — and `onSignedIn()` unconditionally assigns that into `data` and re-saves it to `localStorage`, overwriting whatever was there. In the normal case (file exists) this is harmless. But if the SharePoint file is ever briefly unreachable in a way that Graph reports as a 404 rather than a network error (e.g. mid-move, permissions hiccup, wrong path resolved transiently) — or if the file is genuinely deleted/renamed by anyone — every device that then signs in gets its local cache blown away to zero bays, and that empty state is what gets displayed and is what the *next* `save()` will push back up to SharePoint, clobbering the real data. A real network error (fetch throwing) is handled more safely — it returns `null` and the caller keeps the existing local copy — it's specifically the 404 path that's dangerous.
- **Deleting a bay that's mid-lifecycle silently discards or strands live in-flight data.** `confirmDelete` has no check on `inProgress` or on whether another bay references this one. Two concrete failure modes: (a) delete a bay that's mid-turn as the *source* — the destination bay is left permanently locked with `inProgress:{type:'turning-dest', sourceId: <now-deleted id>}` and no code path ever clears it, since only `confirmCompleteTurn` (looked up via the *source's* `targetId`) does that; (b) delete a bay that's mid-turn as the *destination* — the source bay's `confirmCompleteTurn` will run, look up a target that no longer exists, and silently no-op the "copy onto destination" step while still clearing the source, so the batch's contents/dates disappear entirely with no error shown. The delete confirmation dialog doesn't even mention that a bay is in progress, so a user has no warning they're about to do this.
- **`restoreJSON` can silently replace live production data in one click with no backup and weak validation.** Validation is just "does the parsed object have `.bays` and `.log` keys" — it doesn't check they're arrays, that bay objects have the expected shape, or that this isn't, say, last month's export being restored over today's data. There's a single `confirm()` and no automatic backup of the current state before the overwrite, and the overwrite is pushed to the shared SharePoint file immediately via the normal `save()` path — so a wrong file, or an old backup, chosen by accident replaces everyone's current data with no undo beyond whatever backup the user happens to have saved themselves beforehand.

### MEDIUM

- **No periodic re-sync — the "last write wins" conflict window is a full session, not a few seconds.** `data` is loaded from SharePoint once at sign-in and never refreshed again while the tab stays open (no interval, no `visibilitychange`/focus refetch). If two people have the tracker open at once — plausible for a yard tool used across a shift — whoever saves last simply overwrites the other's changes with their own stale-at-load-time copy of everything else, with no warning to either party. This isn't a corruption risk by itself but it's a realistic way for genuine actions (e.g. someone else starting a turn) to vanish.
- **`getStatus` has no branch for `'turning-dest'`, so a bay receiving a turn displays as "Empty / Available"** in its status badge and gets no progress bar (`progInfo` returns `null` for the `empty` status it falls through to). The card *does* separately show an "in-progress" banner ("Receiving turned compost…") and correctly hides its action buttons, and `startTurn`'s own "pick an empty bay" list correctly excludes any bay with `inProgress` set — so this doesn't cause a data-corrupting double-booking. But the status badge itself is actively misleading at a glance, which matters for a tool meant to be read quickly in the yard.
- **No date validation anywhere a date is entered.** Every "start X" / "complete X" modal is a bare `<input type="date">` with no `min`/`max` — a completion date before its own start date, or a future date, is accepted without complaint and will produce negative or nonsensical day-counts in the maturation timers and CSV exports. Low likelihood of malicious misuse; realistic risk is a mis-tap or a forgotten "actually this happened last Tuesday" backdating that quietly skews the 28-day/21-day countdown.
- **Bay-name collisions aren't prevented**, and `addLog` stores the bay's *name* (not just its id) at the time of the event. Renaming isn't currently exposed in the UI, but two bays sharing a name (nothing stops it at `confirmAddBay`) would make their log history and CSV exports indistinguishable from each other by name.

### LOW

- `emptyStartDate` / `emptyComplete` fields exist in the data model comment but aren't written or read anywhere in the current empty-bay flow (`confirmStartEmpty`/`confirmCompleteEmpty` use `inProgress` instead) — dead/legacy fields, harmless.
- `saveToOneDrive()` is fire-and-forget from every caller — a failed save only surfaces as a sync-status pill going red (`setSyncStatus('error', ...)`); there's no retry and no queue, so a save that fails while the tab is later closed is simply lost (though the `localStorage` copy is safe and will re-attempt sync next `save()`).
- The app's own naming ("OneDrive") throughout code, UI text and the folder path (`Documents/Compost Bay Tracker`) actually points at a SharePoint site drive, not a personal OneDrive — no functional impact, just a documentation/mental-model trap for future maintenance.
- `getToken()`'s fallback to `acquireTokenPopup()` on silent-token failure (e.g. expired refresh token) requires an interactive popup — on a shared yard tablet this can silently degrade to "no token" (if the popup is blocked or nobody's watching) and the app quietly drops into offline/local-only mode without a very prominent warning that OneDrive sync is off. The sync-status pill does say "offline" but it's a small UI element easy to miss on a busy screen.

## Left open, deliberately

Per your instruction, none of the above has been changed. Recommended priority if/when you want fixes, from least to most invasive:

1. Fix the `'turning-dest'` status-badge gap (pure display fix, zero data-model risk).
2. Add a guard on `confirmDelete` that blocks/warns on deleting a bay with `inProgress` set, or that is currently referenced as another bay's `targetId`/`sourceId`.
3. Change the 404 handling in `loadFromOneDrive()` to not silently return an empty dataset — e.g. treat a 404 as "not yet created" only the very first time, or require an explicit confirmation before an empty result is allowed to overwrite a non-empty local cache.
4. Add a lightweight periodic re-sync or at least a "someone else has saved since you loaded — reload before continuing?" check before `save()` pushes.
5. Strengthen `restoreJSON`: validate bay shape, show a diff/preview before committing, and auto-download a timestamped backup of the *current* data before overwriting.
6. Add `min`/basic sanity checks to the date inputs.

## Changelog

- **2026-08-20** — Completed the Phase 1 as-is audit and wrote this document (previously no documentation existed for this system). No code changes made — per your explicit instruction, this system carries live production data and findings are being surfaced for direction rather than auto-fixed the way General Management and Machine Management were.
