#!/usr/bin/env node
/**
 * Pulls the published Outlook ICS feed and turns it into events.json —
 * the same shape the calendar page already reads.
 *
 * Runs on a GitHub Actions runner, not in a browser, so CORS doesn't apply:
 * this is a plain server-to-server fetch.
 *
 * Category / tentative / note are NOT reliably present in a bare ICS feed,
 * so those are supplied by hand in overrides.json, keyed by a stable id
 * built from the event's title + start date. Anything not in overrides.json
 * falls back to sensible defaults so nothing is ever dropped.
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

const VALID_CATS = new Set(["release", "patch", "mobile", "freeze", "note"]);

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

    events.push({
      title: String(item.summary).trim(),
      start,
      end,
      cat: VALID_CATS.has(ov.cat) ? ov.cat : "release",
      tentative: typeof ov.tentative === "boolean" ? ov.tentative : false,
      note: typeof ov.note === "string" ? ov.note : (item.location || "")
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
