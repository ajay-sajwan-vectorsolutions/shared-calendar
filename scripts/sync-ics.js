#!/usr/bin/env node
/**
 * Pulls the published Outlook ICS feed and turns it into events.json —
 * the same shape the calendar page already reads.
 *
 * Runs on a GitHub Actions runner, not in a browser, so CORS doesn't apply:
 * this is a plain server-to-server fetch.
 *
 * Source-of-truth policy: the ICS feed is always read first for every field
 * that has a real ICS equivalent — title/dates (SUMMARY/DTSTART/DTEND),
 * tentative (Outlook's X-MICROSOFT-CDO-BUSYSTATUS), and note (DESCRIPTION,
 * then LOCATION). overrides.json (keyed by a stable id built from the
 * event's title + start date) is consulted only as a FALLBACK, when the
 * feed has nothing usable for that field — e.g. a non-Outlook ICS source
 * with no busystatus, or an event with no description/location.
 *
 * Category is the one exception: verified empirically (checked the raw ICS
 * text directly) that Outlook's published feed carries no category/color
 * data at all — nothing to prefer over overrides.json there, so cat comes
 * from overrides.json exclusively and defaults to "release" when absent.
 */
// node-ical builds all-day (date-only) VEVENTs using the process's own
// system timezone, not UTC — so pin it to UTC before node-ical is loaded.
// This has to happen before the require() below, or all-day dates will
// silently shift by a day on any runner not already set to UTC.
process.env.TZ = "UTC";

const fs = require("fs");
const path = require("path");
const ical = require("node-ical");

const ICS_URL = process.env.ICS_URL;
const ROOT = path.join(__dirname, "..");
const OVERRIDES_PATH = path.join(ROOT, "overrides.json");
const OUT_EVENTS = path.join(ROOT, "events.json");
const OUT_SYNC = path.join(ROOT, "last-sync.json");

// Emergency kill switch: set the repo variable/secret DISABLE_HTML_NOTES to
// "true" to stop calling Outlook's undocumented API entirely — e.g. if
// Microsoft ever changes or blocks it — without touching code. Plain-text
// notes (below) are unaffected either way.
const HTML_NOTES_ENABLED = process.env.DISABLE_HTML_NOTES !== "true";

const VALID_CATS = new Set(["release", "patch", "mobile", "devcomplete", "freeze", "note"]);

function pad(n) { return String(n).padStart(2, "0"); }

// All-day VEVENTs (DTSTART;VALUE=DATE) come back from node-ical as UTC
// midnight with no real timezone — read them with UTC getters or the date
// silently shifts back a day for anyone running this west of UTC.
function toDateStrUTC(d) {
  return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
}

// Timed VEVENTs (e.g. a 10pm ET patch window) are real instants. "Which
// calendar day" depends on which timezone you ask in — this deliberately
// asks in CALENDAR_TZ (below) rather than the runner's own timezone, so a
// 10pm ET Monday patch always lands on Monday regardless of where the
// Action happens to execute.
const CALENDAR_TZ = process.env.CALENDAR_TZ || "America/New_York";
function toDateStrLocal(d) {
  // en-CA gives YYYY-MM-DD directly, which is the whole reason to use it.
  return new Intl.DateTimeFormat("en-CA", { timeZone: CALENDAR_TZ }).format(d);
}

// node-ical gives all-day events an "end" that's exclusive (the day AFTER
// the event). Roll it back one day so multi-day bars render inclusively,
// matching what the calendar page expects.
function inclusiveEnd(ev) {
  const isAllDay = ev.datetype === "date"; // node-ical flag for date-only VEVENTs
  const toStr = isAllDay ? toDateStrUTC : toDateStrLocal;

  const start = ev.start;
  let end = ev.end || ev.start;
  if (isAllDay && end > start) {
    end = new Date(end.getTime() - 86400000);
  }
  return { start: toStr(start), end: toStr(end) };
}

function slug(title, start) {
  return (
    title.toLowerCase().trim().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "") +
    "@" + start
  );
}

// node-ical strips a leading "X-" from custom properties it doesn't parse
// natively, but otherwise passes the token through as-is — look it up
// case-insensitively so this doesn't silently break if some other calendar
// client (or a future Outlook version) capitalizes it differently.
function rawProp(item, name) {
  const target = name.toLowerCase();
  const key = Object.keys(item).find(k => k.toLowerCase() === target);
  return key ? item[key] : undefined;
}

// Outlook's own STATUS is always "CONFIRMED" regardless of Show-As — the
// real tentative/busy/free/OOF signal lives in the Microsoft extension
// property below. Returns undefined (not false!) when the feed doesn't
// have it at all, so the caller knows to fall back to overrides.json
// instead of wrongly treating "no data" as "confirmed".
function icsTentative(item) {
  const busy = rawProp(item, "MICROSOFT-CDO-BUSYSTATUS");
  return typeof busy === "string" ? busy.toUpperCase() === "TENTATIVE" : undefined;
}

// Outlook's ICS export puts a numbered-list marker ("1.", "2.", …) on its
// own line, with the item's actual text on the line after it — e.g.
// "1.\nAdmin Dashboard (Redesigned)" instead of "1. Admin Dashboard
// (Redesigned)". Verified directly against the raw feed: this is what
// Outlook sends, not a rendering bug. Rejoin bare marker lines with
// whatever follows so numbered lists in descriptions read normally.
function joinListMarkers(text) {
  return text.replace(/^[ \t]*(\d+\.)[ \t]*\r?\n[ \t]*/gm, "$1 ");
}

// Prefer the feed's own DESCRIPTION, then LOCATION, over anything hand-typed
// in overrides.json. Outlook always emits DESCRIPTION (even if just "\n"),
// so this trims whitespace-only values down to "" and treats that as "the
// feed has nothing" rather than a real (empty) note.
function icsNote(item) {
  const description = joinListMarkers(String(item.description || "").trim());
  if (description) return description;
  const location = String(item.location || "").trim();
  if (location) return location;
  return "";
}

/* ---------- optional: real HTML note bodies via Outlook's OWA API ----------
 * The ICS feed carries no HTML at all (verified directly against the raw
 * feed — no X-ALT-DESC, DESCRIPTION is plain text only). The only place the
 * real, Outlook-formatted body (tables, etc.) exists is Outlook's own
 * "published calendar" web view, which loads it from an undocumented,
 * unsupported internal JSON API (GetAnonymousCalendarSessionData ->
 * FindItem -> GetItem). Confirmed empirically (plain Node fetch, no
 * browser/cookies) that this works and returns byte-identical bodies across
 * independent sessions.
 *
 * This entire block is best-effort enrichment on top of the plain-text
 * `note` computed above, which is always present regardless. Every step
 * fails soft: a bad response, a network error, or an unmatched event just
 * means that event (or every event, if the session/search step itself
 * fails) keeps its plain-text note instead of gaining a rich one — never a
 * crash, never a missing note.
 */

// EWS wants a legacy Windows timezone name, not the IANA one CALENDAR_TZ
// uses. Only the common zones an ops person might realistically set
// CALENDAR_TZ to are listed — an unlisted zone still works, it just falls
// back to "Eastern Standard Time" for the OWA query only (today's default),
// which at worst costs a few missed matches, never a wrong note.
const IANA_TO_EWS_TZ = {
  "America/New_York": "Eastern Standard Time",
  "America/Chicago": "Central Standard Time",
  "America/Denver": "Mountain Standard Time",
  "America/Los_Angeles": "Pacific Standard Time",
  "UTC": "UTC"
};

const OWA_REQUEST_TIMEOUT_MS = 15000;
const OWA_USER_AGENT =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";

// Derives the OWA published-calendar API base from ICS_URL at runtime.
// ICS_URL is a secret (it grants read access to the whole calendar) — it
// must never be hardcoded or logged. The .ics link and the service.svc API
// link share the same {calendarId}/{publishId} path segment and differ only
// in folder name and filename, so the transform is a straight regex swap.
// Returns null (meaning "HTML notes unavailable this run") if ICS_URL
// doesn't look like the expected published-calendar shape.
function deriveOwaServiceBase(icsUrl) {
  try {
    const u = new URL(icsUrl);
    const m = u.pathname.match(/^\/owa\/calendar\/([^/]+)\/([^/]+)\/calendar\.ics$/i);
    if (!m) return null;
    u.pathname = `/owa/published/${m[1]}/${m[2]}/service.svc`;
    u.search = "";
    u.hash = "";
    return u.toString();
  } catch {
    return null;
  }
}

function owaHeaders(action, sessionId, bodyObj) {
  return {
    "content-type": "application/json; charset=utf-8",
    action,
    "x-req-source": "PublishedCalendar",
    "x-owa-actionsource": action,
    "x-owa-canary": "X-OWA-CANARY_cookie_is_null_or_empty",
    "x-owa-sessionid": sessionId,
    "x-owa-hosted-ux": "false",
    "prefer": 'exchange.behavior="IncludeThirdPartyOnlineMeetingProviders"',
    "user-agent": OWA_USER_AGENT,
    // The actual request payload travels in this header, url-encoded JSON —
    // not the POST body, which OWA leaves empty. Verified directly; this
    // isn't a documented convention.
    "x-owa-urlpostdata": encodeURIComponent(JSON.stringify(bodyObj))
  };
}

async function owaCall(base, action, n, bodyObj, sessionId) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), OWA_REQUEST_TIMEOUT_MS);
  try {
    const res = await fetch(`${base}?action=${action}&app=PublishedCalendar&n=${n}`, {
      method: "POST",
      headers: owaHeaders(action, sessionId, bodyObj),
      signal: controller.signal
    });
    if (!res.ok) throw new Error(`${action} responded HTTP ${res.status}`);
    return await res.json();
  } finally {
    clearTimeout(timer);
  }
}

function ewsHeader(tzId) {
  return {
    __type: "JsonRequestHeaders:#Exchange",
    RequestServerVersion: "V2018_01_08",
    TimeZoneContext: {
      __type: "TimeZoneContext:#Exchange",
      TimeZoneDefinition: { __type: "TimeZoneDefinitionType:#Exchange", Id: tzId }
    }
  };
}

// EWS's Paging.StartDate/EndDate want a bare "YYYY-MM-DDTHH:mm:ss.mmm" (no
// Z, no offset) — it's interpreted in whatever TimeZoneContext.Id says.
// The events passed in only ever need day-level precision here, and the
// range gets padded by a day on each side, so treating the UTC clock value
// as if it were local wall-clock time (which is all this does) is close
// enough for a search window — never used for anything more precise.
function toEwsDateTime(d) {
  const p2 = n => String(n).padStart(2, "0");
  return `${d.getUTCFullYear()}-${p2(d.getUTCMonth() + 1)}-${p2(d.getUTCDate())}` +
    `T${p2(d.getUTCHours())}:${p2(d.getUTCMinutes())}:${p2(d.getUTCSeconds())}.000`;
}

// Best-effort enrichment: returns Map<slugId, htmlString> for every event
// this could confidently match to a real Outlook item. Matching is by
// title + local start date only, requiring EXACTLY one FindItem candidate —
// zero or multiple candidates both mean "don't guess", leaving that event
// on plain text. Start dates are compared as the literal "YYYY-MM-DD" slice
// of the offset-bearing string OWA returns (e.g. "2026-08-07T21:00:00-04:00"
// -> "2026-08-07") — never round-tripped through `new Date(...).toISOString()`,
// which converts to UTC first and would silently roll an evening event onto
// the next day.
async function fetchHtmlNotes(icsUrl, events) {
  const result = new Map();
  if (!HTML_NOTES_ENABLED) {
    console.log("HTML notes disabled (DISABLE_HTML_NOTES=true) — plain text only.");
    return result;
  }
  if (events.length === 0) return result;

  const base = deriveOwaServiceBase(icsUrl);
  if (!base) {
    console.warn("ICS_URL doesn't look like an OWA published-calendar link — skipping HTML notes, plain text only.");
    return result;
  }

  const tzId = IANA_TO_EWS_TZ[CALENDAR_TZ];
  if (!tzId) {
    console.warn(`CALENDAR_TZ "${CALENDAR_TZ}" has no known EWS equivalent — HTML-note matching may miss more than usual (falls back to plain text for those).`);
  }
  const header = ewsHeader(tzId || "Eastern Standard Time");
  const sessionId = `sc-${process.pid}-${Date.now()}`;

  let folderId;
  try {
    const session = await owaCall(base, "GetAnonymousCalendarSessionData", 0,
      { __type: "GetAnonymousCalendarSessionDataJsonRequest:#Exchange", Header: header },
      sessionId);
    folderId = session?.Body?.CalendarFolder?.FolderId?.Id;
    if (!folderId) throw new Error("response had no Body.CalendarFolder.FolderId");
  } catch (err) {
    console.warn("OWA session bootstrap failed — skipping HTML notes, plain text only:", err.message);
    return result;
  }

  const starts = events.map(e => new Date(e.start + "T00:00:00Z")).sort((a, b) => a - b);
  const ends = events.map(e => new Date(e.end + "T00:00:00Z")).sort((a, b) => a - b);
  const rangeStart = new Date(starts[0].getTime() - 86400000);
  const rangeEnd = new Date(ends[ends.length - 1].getTime() + 2 * 86400000);

  let items;
  try {
    const found = await owaCall(base, "FindItem", 1, {
      __type: "FindItemJsonRequest:#Exchange",
      Header: header,
      Body: {
        __type: "FindItemRequest:#Exchange",
        ParentFolderIds: [{ __type: "FolderId:#Exchange", Id: folderId }],
        ItemShape: { __type: "ItemResponseShape:#Exchange", BaseShape: "IdOnly" },
        Traversal: "Shallow",
        Paging: {
          __type: "CalendarPageView:#Exchange",
          StartDate: toEwsDateTime(rangeStart),
          EndDate: toEwsDateTime(rangeEnd)
        }
      }
    }, sessionId);
    items = found?.Body?.ResponseMessages?.Items?.[0]?.RootFolder?.Items;
    if (!Array.isArray(items)) throw new Error("response had no Body.ResponseMessages.Items[0].RootFolder.Items");
  } catch (err) {
    console.warn("OWA FindItem failed — skipping HTML notes, plain text only:", err.message);
    return result;
  }

  const byKey = new Map();
  for (const it of items) {
    if (!it?.ItemId?.Id) continue;
    const startDate = String(it.Start || "").slice(0, 10);
    const k = `${it.Subject || ""}@@${startDate}`;
    if (!byKey.has(k)) byKey.set(k, []);
    byKey.get(k).push(it);
  }

  let matched = 0;
  for (const ev of events) {
    const candidates = byKey.get(`${ev.title}@@${ev.start}`) || [];
    if (candidates.length !== 1) continue; // none, or ambiguous — don't guess
    const itemId = candidates[0].ItemId.Id;

    try {
      const detail = await owaCall(base, "GetItem", 2, {
        __type: "GetItemJsonRequest:#Exchange",
        Header: header,
        Body: {
          __type: "GetItemRequest:#Exchange",
          ItemIds: [{ __type: "ItemId:#Exchange", Id: itemId }],
          ShapeName: "FullCalendarItem",
          ItemShape: { __type: "ItemResponseShape:#Exchange", BaseShape: "IdOnly" }
        }
      }, sessionId);
      const body = detail?.Body?.ResponseMessages?.Items?.[0]?.Items?.[0]?.Body;
      const html = body?.BodyType === "HTML" ? String(body.Value || "").trim() : "";
      if (html) {
        result.set(slug(ev.title, ev.start), html);
        matched++;
      }
    } catch (err) {
      console.warn(`GetItem failed for "${ev.title}" (${ev.start}) — that event stays plain text:`, err.message);
    }
  }

  console.log(`OWA HTML notes: matched ${matched} of ${events.length} event(s).`);
  return result;
}

function loadOverrides() {
  if (!fs.existsSync(OVERRIDES_PATH)) return {};
  try {
    return JSON.parse(fs.readFileSync(OVERRIDES_PATH, "utf8"));
  } catch (err) {
    console.error("overrides.json is not valid JSON — ignoring it this run:", err.message);
    return {};
  }
}

async function main() {
  if (!ICS_URL) {
    console.error("ICS_URL is not set (expected as a repo secret / env var). Aborting.");
    process.exit(1);
  }

  const overrides = loadOverrides();
  const data = await ical.async.fromURL(ICS_URL);

  const events = [];
  for (const item of Object.values(data)) {
    if (item.type !== "VEVENT") continue;
    if (!item.start || !item.summary) continue;

    const { start, end } = inclusiveEnd(item);
    const id = slug(item.summary, start);
    const ov = overrides[id] || {};

    // ICS first, overrides.json only fills in what the feed didn't have.
    const feedTentative = icsTentative(item);
    const tentative = typeof feedTentative === "boolean"
      ? feedTentative
      : (typeof ov.tentative === "boolean" ? ov.tentative : false);

    const feedNote = icsNote(item);
    const note = feedNote || (typeof ov.note === "string" ? ov.note : "");

    events.push({
      title: String(item.summary).trim(),
      start,
      end,
      // No ICS equivalent exists for this (verified: Outlook's feed carries
      // no category data at all) — overrides.json is the only source.
      cat: VALID_CATS.has(ov.cat) ? ov.cat : "release",
      tentative,
      note
    });
  }

  events.sort((a, b) => a.start.localeCompare(b.start) || a.title.localeCompare(b.title));

  // Best-effort: attach a real HTML body (ev.noteHtml) wherever we can
  // confidently match one. This can never fail the sync — every event
  // already has a plain-text `note` from the loop above regardless of
  // whether this succeeds, fails, or is disabled.
  try {
    const htmlNotes = await fetchHtmlNotes(ICS_URL, events);
    for (const ev of events) {
      const html = htmlNotes.get(slug(ev.title, ev.start));
      if (html) ev.noteHtml = html;
    }
  } catch (err) {
    console.warn("HTML notes enrichment failed unexpectedly — continuing with plain text only:", err.message);
  }

  // Only write + let the workflow commit if something actually changed,
  // so a no-op hourly run doesn't spam the history with empty commits.
  const prev = fs.existsSync(OUT_EVENTS) ? fs.readFileSync(OUT_EVENTS, "utf8") : "";
  const next = JSON.stringify(events, null, 2) + "\n";

  fs.writeFileSync(OUT_SYNC, JSON.stringify({
    synced_at: new Date().toISOString(),
    event_count: events.length
  }, null, 2) + "\n");

  if (prev.trim() === next.trim()) {
    console.log(`No change — ${events.length} events, feed unchanged since last sync.`);
    return;
  }

  fs.writeFileSync(OUT_EVENTS, next);
  console.log(`Wrote events.json — ${events.length} events.`);

  const unmapped = events.filter(e => !overrides[slug(e.title, e.start)]);
  if (unmapped.length) {
    console.log(
      `\n${unmapped.length} event(s) have no entry in overrides.json and defaulted to cat:"release":\n` +
      unmapped.map(e => `  - ${slug(e.title, e.start)}`).join("\n") +
      `\n\nAdd them to overrides.json to set their color / tentative flag / note.`
    );
  }
}

main().catch(err => {
  console.error("Sync failed:", err);
  process.exit(1);
});
