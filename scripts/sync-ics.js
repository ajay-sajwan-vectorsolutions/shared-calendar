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

// Prefer the feed's own DESCRIPTION, then LOCATION, over anything hand-typed
// in overrides.json. Outlook always emits DESCRIPTION (even if just "\n"),
// so this trims whitespace-only values down to "" and treats that as "the
// feed has nothing" rather than a real (empty) note.
function icsNote(item) {
  const description = String(item.description || "").trim();
  if (description) return description;
  const location = String(item.location || "").trim();
  if (location) return location;
  return "";
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
