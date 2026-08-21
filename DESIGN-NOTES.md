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

- ~~A failed or 404'd load from SharePoint silently wipes local data to empty.~~ **Fixed** (commit `5922512`). `loadFromOneDrive()` treats a 404 (file genuinely not there yet) as "return `{bays:[],log:[]}`" — and `onSignedIn()` unconditionally assigns that into `data` and re-saves it to `localStorage`, overwriting whatever was there. In the normal case (file exists) this is harmless. But if the SharePoint file is ever briefly unreachable in a way that Graph reports as a 404 rather than a network error (e.g. mid-move, permissions hiccup, wrong path resolved transiently) — or if the file is genuinely deleted/renamed by anyone — every device that then signs in gets its local cache blown away to zero bays, and that empty state is what gets displayed and is what the *next* `save()` will push back up to SharePoint, clobbering the real data. A real network error (fetch throwing) is handled more safely — it returns `null` and the caller keeps the existing local copy — it's specifically the 404 path that's dangerous. Fix: on a 404, `loadFromOneDrive()` now checks whether this device's own local cache shows real usage (`looksLikeRealLocalData()` — any bay with a fill/turn/empty date or in-progress state, or any log entry beyond the one-time 8-bay seed entry). If it does, `data` is left untouched and the user gets a blocking warning instead of a silent wipe. A genuinely fresh device (no real local data yet) still treats a 404 as normal first-time setup, same as before.
- **Deleting a bay that's mid-lifecycle silently discards or strands live in-flight data.** `confirmDelete` has no check on `inProgress` or on whether another bay references this one. Two concrete failure modes: (a) delete a bay that's mid-turn as the *source* — the destination bay is left permanently locked with `inProgress:{type:'turning-dest', sourceId: <now-deleted id>}` and no code path ever clears it, since only `confirmCompleteTurn` (looked up via the *source's* `targetId`) does that; (b) delete a bay that's mid-turn as the *destination* — the source bay's `confirmCompleteTurn` will run, look up a target that no longer exists, and silently no-op the "copy onto destination" step while still clearing the source, so the batch's contents/dates disappear entirely with no error shown. **Note (2026-08-20):** on closer inspection, the "Remove" button in the UI only ever renders when a bay's status is `empty` and unlocked (`cardHTML`'s `s==='empty'&&!locked` guard) — so on a single device, in normal use, there is no click path that reaches `confirmDelete` on a bay that's mid-fill/turn/empty. The real exposure is the cross-device staleness window described in the next item below: Device A can still be *looking at* a stale render (Bay C shown as empty because Device A hasn't reloaded since Device B started using it) and click Remove on what is, on the server, no longer an empty bay. A guard inside `confirmDelete` itself can't catch that case either, since it would only be checking the same stale local data that rendered the button. Deliberately left open — see "Left open, deliberately" below.
- ~~`restoreJSON` can silently replace live production data in one click with no backup and weak validation.~~ **Fixed** (commit `e13621c`). Validation was just "does the parsed object have `.bays` and `.log` keys" — it didn't check they're arrays, that bay objects have the expected shape, or that this isn't, say, last month's export being restored over today's data. There was a single `confirm()` and no automatic backup of the current state before the overwrite, and the overwrite was pushed to the shared SharePoint file immediately via the normal `save()` path — so a wrong file, or an old backup, chosen by accident replaced everyone's current data with no undo beyond whatever backup the user happened to have saved themselves beforehand. Fix: `validateRestoreData()` now checks `bays`/`log` are actually arrays and every bay has an id/name before anything touches `data`; a confirmation modal shows current-vs-incoming bay/log counts and warns if any bay is currently mid-lifecycle and would be discarded; and confirming always downloads a timestamped backup of the *current* data first, before the overwrite — so a bad restore decision is recoverable from that file.

### MEDIUM

- ~~No periodic re-sync — the "last write wins" conflict window is a full session, not a few seconds.~~ **Narrowed** (commit `d5ba0c6`), not fully closed. `data` is loaded from SharePoint once at sign-in and never refreshed again while the tab stays open (no interval, no `visibilitychange`/focus refetch). If two people have the tracker open at once — plausible for a yard tool used across a shift — whoever saves last simply overwrites the other's changes with their own stale-at-load-time copy of everything else, with no warning to either party. Fix: `saveToOneDrive()` now does a lightweight metadata-only check of the file's current server-side eTag immediately before writing; if it's changed since our last load/save, the write is aborted and the user is warned to reload rather than silently overwriting. This is a best-effort pre-check, not a true atomic compare-and-swap on the write itself — Microsoft's docs for the plain content-upload endpoint this app uses don't document `If-Match` support (unlike the separate driveItem metadata PATCH endpoint), and risking every save failing outright on an unverified header wasn't an acceptable trade on a live system. A small race remains between the check and the write; there is still no periodic background re-sync while a tab sits open.
- ~~`getStatus` has no branch for `'turning-dest'`, so a bay receiving a turn displays as "Empty / Available"~~ **Fixed** (commit `bb9ad5f`). It had no progress bar either (`progInfo` returned `null` for the `empty` status it fell through to). The card's own "in-progress" banner ("Receiving turned compost…") and `startTurn`'s "pick an empty bay" list were already correct, so this was a status-badge/progress-bar display bug only — no data-model change needed. Fix added a `turning-dest` branch to `getStatus()`, `STATUS_LABEL`/`STATUS_ICON`/`STATUS_COLOR`, `progInfo()`, and the bay visualiser strip.
- **No date validation anywhere a date is entered.** Every "start X" / "complete X" modal is a bare `<input type="date">` with no `min`/`max` — a completion date before its own start date, or a future date, is accepted without complaint and will produce negative or nonsensical day-counts in the maturation timers and CSV exports. Low likelihood of malicious misuse; realistic risk is a mis-tap or a forgotten "actually this happened last Tuesday" backdating that quietly skews the 28-day/21-day countdown.
- **Bay-name collisions aren't prevented**, and `addLog` stores the bay's *name* (not just its id) at the time of the event. Renaming isn't currently exposed in the UI, but two bays sharing a name (nothing stops it at `confirmAddBay`) would make their log history and CSV exports indistinguishable from each other by name.

### LOW

- `emptyStartDate` / `emptyComplete` fields exist in the data model comment but aren't written or read anywhere in the current empty-bay flow (`confirmStartEmpty`/`confirmCompleteEmpty` use `inProgress` instead) — dead/legacy fields, harmless.
- `saveToOneDrive()` is fire-and-forget from every caller — a failed save only surfaces as a sync-status pill going red (`setSyncStatus('error', ...)`); there's no retry and no queue, so a save that fails while the tab is later closed is simply lost (though the `localStorage` copy is safe and will re-attempt sync next `save()`).
- The app's own naming ("OneDrive") throughout code, UI text and the folder path (`Documents/Compost Bay Tracker`) actually points at a SharePoint site drive, not a personal OneDrive — no functional impact, just a documentation/mental-model trap for future maintenance.
- `getToken()`'s fallback to `acquireTokenPopup()` on silent-token failure (e.g. expired refresh token) requires an interactive popup — on a shared yard tablet this can silently degrade to "no token" (if the popup is blocked or nobody's watching) and the app quietly drops into offline/local-only mode without a very prominent warning that OneDrive sync is off. The sync-status pill does say "offline" but it's a small UI element easy to miss on a busy screen.

## Left open, deliberately

Remaining items, from least to most invasive:

1. ~~Add `min`/basic sanity checks to the date inputs.~~ **Fixed 21 Aug 2026** (`24b74a8`) — a shared `readValidDate()` guard rejects blank/future dates and (for completing a turn) a completion date before the turn's start; date pickers also capped at today.
2. ~~Prevent bay-name collisions at `confirmAddBay`.~~ **Fixed 21 Aug 2026** (`24b74a8`) — duplicate names (case/space-insensitive) are now rejected.
3. Revisit the `confirmDelete` mid-lifecycle case — genuinely useful protection here likely means closing the cross-device staleness window further (more frequent/targeted re-sync, or a re-check immediately before the delete itself goes through) rather than a guard inside `confirmDelete` alone, which can't see past this device's own stale view of the world. Low urgency: the UI already blocks the straightforward single-device path to this.
4. Close the remaining "no periodic background re-sync" gap properly, if it keeps mattering in practice — the pre-save conflict check (commit `d5ba0c6`) narrows the window but a device can still sit on a stale view for a full session between loads.

## Changelog

- **2026-08-20** — Completed the Phase 1 as-is audit and wrote this document (previously no documentation existed for this system).
- **2026-08-21** — Non-urgent cleanup pass (`24b74a8`): added date-input sanity checks (`readValidDate()` + `max=today` on the pickers) and a duplicate-bay-name guard at `confirmAddBay`, both from the 'Left open' list. No change to the sync/storage path. Verified: app script passes `node --check`; date/collision logic unit-tested.
- **2026-08-20** — Fixed all four originally agreed findings, one at a time with a check-in on each given this system's live-data status: the `turning-dest` status-badge display bug (`bb9ad5f`), a pre-save conflict check against concurrent SharePoint edits (`d5ba0c6`), the 404-on-load silent data wipe — the audit's highest-severity finding (`5922512`), and `restoreJSON` hardening with validation/preview/auto-backup (`e13621c`). The mid-lifecycle delete finding was re-assessed during implementation and found to be already blocked at the UI level for the single-device case described in the original audit — the Remove button only ever renders on an empty, unlocked bay; left open pending a proper fix for the underlying cross-device staleness window instead (see "Left open, deliberately"). Deployment verified live via the GitHub Pages site after each push; full end-to-end sync verification against live SharePoint wasn't possible from automated tooling (silent-auth fallback doesn't complete in that context — a known, pre-existing low-severity gap, not a regression) so a manual spot-check by a signed-in user is still worthwhile.
