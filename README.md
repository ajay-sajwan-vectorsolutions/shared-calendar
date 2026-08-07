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

## Source of truth: ICS first, overrides.json as fallback only

For every field that has a real ICS equivalent, the feed wins — `overrides.json`
(keyed by an id built from the event's title and start date) only fills in
what the feed didn't have:

- **title / dates** — always from the feed (`SUMMARY` / `DTSTART` / `DTEND`).
  overrides.json has no way to override these.
- **tentative** — from Outlook's own `X-MICROSOFT-CDO-BUSYSTATUS` ("Show As:
  Tentative" in the Outlook UI), if the feed has it. Falls back to
  `overrides.json`'s `tentative` flag only for a feed that doesn't carry
  that field at all (e.g. a non-Outlook ICS source).
- **note** — from the feed's `DESCRIPTION`, then `LOCATION`, if either is
  non-empty. Falls back to `overrides.json`'s `note` only when the feed has
  neither.
- **cat (category/color)** — the one exception. Verified directly against
  the raw ICS text: Outlook's published feed carries **no category data at
  all**, so there's nothing to prefer over overrides.json here — `cat`
  always comes from `overrides.json`, defaulting to `"release"` when the
  event has no entry there.

```
"keystone-dev-complete@2026-08-17": {
  "cat": "devcomplete",
  "tentative": false,
  "note": "Confirmed"
}
```

Valid `cat` values: `release`, `patch`, `mobile`, `devcomplete`, `freeze`, `note`.

Anything in the Outlook feed with no matching entry in `overrides.json`
defaults to `cat:"release"`, and the Action's log lists exactly which ids
are missing so you can copy them in. Run the workflow once (or check its
log) after adding new event types in Outlook to get the exact id string.

## Rich (HTML) notes — best-effort, always with a plain-text fallback

The ICS feed itself carries no HTML — `note` above is always plain text.
Real Outlook formatting (tables, numbered lists, etc.) only exists in
Outlook's own "published calendar" web view, which loads it from an
undocumented, unsupported internal API. `scripts/sync-ics.js` calls that
same API as a **best-effort enrichment on top of** the plain-text note, never
instead of it:

- Every event always gets its plain-text `note` first, exactly as before.
- If (and only if) that event can be matched to Outlook's own copy with
  high confidence (exact title + exact start date, and exactly one such
  match — anything ambiguous is left alone), `events.json` also gets a
  `noteHtml` field with the real formatted body.
- `index.html` sanitizes `noteHtml` with an allowlist (tags: text formatting,
  links, lists, tables; attributes: `href`, `rowspan`/`colspan`, link
  `color`) immediately before display, and falls back to the plain-text
  `note` if `noteHtml` is missing, fails to sanitize, or sanitizes down to
  nothing. A note always renders something.

**This is inherently more fragile than the ICS feed** — it's a private
Microsoft API that could change or be blocked with no notice, unlike the
published ICS link. If it ever breaks, the sync doesn't fail: every event
just keeps its plain-text note, same as before this feature existed. To
turn it off deliberately (e.g. while investigating an issue) without a code
change, add a repository variable/secret `DISABLE_HTML_NOTES` set to `true`.

## Known limits

- `events.json` and `last-sync.json` are generated files — don't hand-edit
  them, since the next sync will overwrite whatever you typed.
- The feed reflects whatever Outlook has published, which lags real edits
  by up to a few hours (sometimes longer) — see the caveats discussed
  when this was built. For same-day changes, still tell the team directly.
- If `ICS_URL` is ever revoked or unpublished in Outlook, the Action will
  fail with a clear error in the Actions tab log rather than silently
  going stale.
