# Release Calendar

A static calendar page (`index.html`) kept in sync with a published Outlook
calendar via a scheduled GitHub Action — no server, no CORS problem, and
every change to the feed shows up as a git commit.

## One-time setup

1. Create a repo and push these files.
2. **Settings → Pages** → deploy from `main` / root. Your page will be at
   `https://<org>.github.io/<repo>/`.
3. **Settings → Secrets and variables → Actions → New repository secret**
   - Name: `ICS_URL`
   - Value: the ICS link from Outlook (Settings → Calendar → Shared
     calendars → Publish a calendar → copy the **ICS** link, not the HTML one)
4. **Actions tab** → run "Sync release calendar from Outlook" once manually
   (`workflow_dispatch`) to confirm it works, rather than waiting for the
   hourly schedule.

## How it fits together

- `scripts/sync-ics.js` fetches the ICS feed **from GitHub's runner**, not
  a browser — so the CORS block you hit testing this from the page itself
  doesn't apply here.
- It writes `events.json` (what the page actually reads) and
  `last-sync.json` (a timestamp the page displays in the header).
- `.github/workflows/sync-calendar.yml` runs that script every hour and
  commits the result **only if something changed** — a no-op hour doesn't
  create an empty commit.

## Setting categories, tentative, and notes

A bare ICS feed doesn't reliably carry Outlook's category colors, so those
are set by hand in `overrides.json`, keyed by an id built from the event's
title and start date:

```
"web-release-pd-tracking@2026-08-12": {
  "cat": "release",
  "tentative": false,
  "note": "Confirmed"
}
```

Valid `cat` values: `release`, `patch`, `mobile`, `freeze`, `note`.

Anything in the Outlook feed that has no matching entry here defaults to
`release` / not tentative, and the Action's log will list exactly which
ids are missing so you can copy them in. Run the workflow once (or check
its log) after adding new event types in Outlook to get the exact id string.

## Known limits

- `events.json` and `last-sync.json` are generated files — don't hand-edit
  them, since the next sync will overwrite whatever you typed.
- The feed reflects whatever Outlook has published, which lags real edits
  by up to a few hours (sometimes longer) — see the caveats discussed
  when this was built. For same-day changes, still tell the team directly.
- If `ICS_URL` is ever revoked or unpublished in Outlook, the Action will
  fail with a clear error in the Actions tab log rather than silently
  going stale.
