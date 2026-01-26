/*************************************************
 * Availability menu for Google Docs
 * Dialog options:
 *   - This week 2 hour blocks        -> show full free segments that are >= 2 hours
 *   - This week full availability    -> show merged free intervals
 *   - Next week 2 hour blocks        -> same filtering for next Mon..Fri
 *   - Next week full availability    -> merged free intervals for next Mon..Fri
 *   - Optional recurring events to ignore (next two weeks)
 *
 * Window: 10:00 to 18:30 local time, Mon..Fri
 * Busy: YES, MAYBE, INVITED, OWNER. Declined is free.
 * Today clipped to now (rounded up to next 30 minutes) for “This week”.
 * Output: 12 hour numbers, no am/pm, drops :00
 **************************************************/

const TIMEZONE = Session.getScriptTimeZone() || "America/New_York";
const START_HOUR = 10;
const START_MIN  = 0;
const END_HOUR   = 18;
const END_MIN    = 30;
const FIT_BLOCK_HOURS = 2;

/* Menu */
function onOpen() {
  DocumentApp.getUi()
    .createMenu("Availability")
    .addItem("Calculate availability", "showAvailabilityDialog")
    .addToUi();
}

function showAvailabilityDialog() {
  const html = HtmlService.createHtmlOutputFromFile("availability_dialog")
    .setWidth(420)
    .setHeight(560);
  DocumentApp.getUi().showModalDialog(html, "Availability options");
}

/* Main */
function insertAvailabilityFromDialog(options) {
  if (!options || !Array.isArray(options.selections) || options.selections.length === 0) {
    throw new Error("Please select at least one availability option.");
  }
  if (!options.calendarIds || options.calendarIds.length === 0) {
    throw new Error("Please select at least one calendar.");
  }
  const ignoreRecurringIds = Array.isArray(options.ignoreRecurringIds)
    ? options.ignoreRecurringIds
    : [];
  const selections = options.selections;
  const calendars = getCalendarsFromIds_(options.calendarIds);
  const lines = [];

  for (const selection of selections) {
    const meta = selectionMeta_(selection);
    if (!meta) continue;
    lines.push(meta.label);
    const sectionLines = buildAvailabilityLines_(meta.mode, meta.rangeKey, ignoreRecurringIds, calendars);
    lines.push(...sectionLines);
    lines.push("");
  }

  if (lines.length > 0 && lines[lines.length - 1] === "") lines.pop();

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  if (!body) throw new Error("No document body found.");

  overwriteBodyWithLines_(body, lines);
  doc.saveAndClose();
}

function buildAvailabilityLines_(mode, rangeKey, ignoreRecurringIds, calendars) {
  const { startDate, endDate } = getRange_(rangeKey);
  const days = enumerateDays_(startDate, endDate).filter(isWeekday_);

  const lines = [];
  for (const day of days) {
    let dayStart = setTime_(day, START_HOUR, START_MIN);
    const dayEnd = setTime_(day, END_HOUR, END_MIN);

    // Clip today's start to current time (rounded up) for THIS_WEEK
    if (rangeKey === "THIS_WEEK" && isSameYmd_(day, todayDate_())) {
      const now = new Date();
      const nowLocal = parseLocal_(Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss"));
      const clipped = ceilToNextHalfHour_(nowLocal);
      if (clipped > dayStart) dayStart = clipped;
    }

    const free = getFreeIntervals_(dayStart, dayEnd, ignoreRecurringIds, calendars);

    const label = formatDayLabel_(day);
    let text;
    if (mode === "ALL_AVAIL") {
      text = intervalsToInlineText_(free);
    } else {
      // FIT_2H: show entire free intervals that are at least 2 hours long
      const minMs = FIT_BLOCK_HOURS * 60 * 60 * 1000;
      const longEnough = free.filter(([s, e]) => (e.getTime() - s.getTime()) >= minMs);
      text = intervalsToInlineText_(longEnough);
    }
    lines.push(`${label}: ${text || "none"}`);
  }

  return lines.length ? lines : ["No weekdays in the selected range."];
}

function selectionMeta_(selectionKey) {
  const map = {
    THIS_WEEK_2H: { mode: "FIT_2H", rangeKey: "THIS_WEEK", label: "This week — 2 hour blocks" },
    THIS_WEEK_ALL: { mode: "ALL_AVAIL", rangeKey: "THIS_WEEK", label: "This week — all availability" },
    NEXT_WEEK_2H: { mode: "FIT_2H", rangeKey: "NEXT_WEEK", label: "Next week — 2 hour blocks" },
    NEXT_WEEK_ALL: { mode: "ALL_AVAIL", rangeKey: "NEXT_WEEK", label: "Next week — all availability" },
  };
  return map[selectionKey] || null;
}

/* Safe clear + write */
function overwriteBodyWithLines_(body, lines) {
  const n = body.getNumChildren();
  for (let i = n - 1; i >= 1; i--) body.removeChild(body.getChild(i));
  let first = body.getChild(0);
  if (!first || first.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    body.insertParagraph(0, "placeholder");
    if (first) body.removeChild(first);
    first = body.getChild(0);
  } else {
    const p = first.asParagraph();
    if (!p.getText()) p.setText("placeholder");
  }

  if (lines.length === 0) {
    first.asParagraph().setText("No weekdays in the selected range.");
    return;
  }
  first.asParagraph().setText(lines[0]);
  for (let i = 1; i < lines.length; i++) body.appendParagraph(lines[i]);
}

/* Range helpers */
function getRange_(rangeKey) {
  const today = todayDate_();
  const dow = parseInt(Utilities.formatDate(today, TIMEZONE, "u"), 10); // 1 Mon..7 Sun

  if (rangeKey === "THIS_WEEK") {
    if (dow > 5) {
      return { startDate: today, endDate: new Date(today.getTime() - 86400000) };
    }
    const end = new Date(today);
    end.setDate(end.getDate() + (5 - dow));
    return { startDate: today, endDate: end };
  } else {
    const monday = new Date(today);
    monday.setDate(monday.getDate() + (8 - dow));
    const friday = new Date(monday);
    friday.setDate(friday.getDate() + 4);
    return { startDate: monday, endDate: friday };
  }
}

function enumerateDays_(startDate, endDate) {
  const out = [];
  const d = new Date(startDate);
  while (d <= endDate) { out.push(new Date(d)); d.setDate(d.getDate() + 1); }
  return out;
}
function isWeekday_(d) {
  const u = parseInt(Utilities.formatDate(d, TIMEZONE, "u"), 10);
  return u >= 1 && u <= 5;
}

/* Calendar busy/free */
function getFreeIntervals_(windowStart, windowEnd, ignoreRecurringIds, calendars) {
  if (windowStart >= windowEnd) return [];
  const calendarsToUse = calendars && calendars.length
    ? calendars
    : [CalendarApp.getDefaultCalendar()];
  const events = [];
  for (const cal of calendarsToUse) {
    events.push(...cal.getEvents(windowStart, windowEnd));
  }

  const busy = [];
  for (const ev of events) {
    if (shouldIgnoreRecurringEvent_(ev, ignoreRecurringIds)) continue;
    const status = ev.getMyStatus && ev.getMyStatus();
    const isBusy =
      status === CalendarApp.GuestStatus.YES ||
      status === CalendarApp.GuestStatus.MAYBE ||
      status === CalendarApp.GuestStatus.INVITED ||
      status === CalendarApp.GuestStatus.OWNER;

    if (!isBusy) continue;

    const s = new Date(Math.max(ev.getStartTime().getTime(), windowStart.getTime()));
    const e = new Date(Math.min(ev.getEndTime().getTime(), windowEnd.getTime()));
    if (s < e) busy.push([s, e]);
  }

  const merged = mergeIntervals_(busy);
  return subtractIntervals_(windowStart, windowEnd, merged);
}

function shouldIgnoreRecurringEvent_(event, ignoreRecurringIds) {
  if (!ignoreRecurringIds || ignoreRecurringIds.length === 0) return false;
  if (!event.isRecurringEvent || !event.isRecurringEvent()) return false;
  const series = event.getEventSeries && event.getEventSeries();
  if (!series) return false;
  return ignoreRecurringIds.indexOf(series.getId()) !== -1;
}

function getRecurringEventSeriesOptions(calendarIds) {
  const today = todayDate_();
  const endDate = new Date(today);
  endDate.setDate(endDate.getDate() + 13);
  const calendars = getCalendarsFromIds_(calendarIds);
  const events = [];
  for (const cal of calendars) {
    events.push(...cal.getEvents(today, endDate));
  }

  const seen = {};
  const options = [];
  for (const ev of events) {
    if (!ev.isRecurringEvent || !ev.isRecurringEvent()) continue;
    const series = ev.getEventSeries && ev.getEventSeries();
    if (!series) continue;
    const seriesId = series.getId();
    if (seen[seriesId]) continue;
    seen[seriesId] = true;
    const start = ev.getStartTime();
    const end = ev.getEndTime();
    const calendarName = ev.getOriginalCalendar && ev.getOriginalCalendar()
      ? ev.getOriginalCalendar().getName()
      : "Calendar";
    options.push({
      id: seriesId,
      title: ev.getTitle(),
      calendarName,
      time: `${formatTime_(start)} to ${formatTime_(end)}`,
    });
  }

  options.sort((a, b) => {
    const nameCompare = a.calendarName.localeCompare(b.calendarName);
    return nameCompare !== 0 ? nameCompare : a.title.localeCompare(b.title);
  });
  return options;
}

function getCalendarOptions() {
  const defaultCalendar = CalendarApp.getDefaultCalendar();
  const calendars = CalendarApp.getAllCalendars();
  return calendars.map((cal) => ({
    id: cal.getId(),
    name: cal.getName(),
    isPrimary: cal.getId() === defaultCalendar.getId(),
  }));
}

function getCalendarsFromIds_(calendarIds) {
  if (!calendarIds || calendarIds.length === 0) {
    return [CalendarApp.getDefaultCalendar()];
  }
  const calendars = calendarIds
    .map((id) => CalendarApp.getCalendarById(id))
    .filter(Boolean);
  return calendars.length ? calendars : [CalendarApp.getDefaultCalendar()];
}

/* Interval math */
function mergeIntervals_(intervals) {
  if (!intervals.length) return [];
  intervals.sort((a, b) => a[0] - b[0]);
  const out = [];
  let [cs, ce] = intervals[0];
  for (let i = 1; i < intervals.length; i++) {
    const [s, e] = intervals[i];
    if (s <= ce) ce = new Date(Math.max(ce.getTime(), e.getTime()));
    else { out.push([cs, ce]); [cs, ce] = [s, e]; }
  }
  out.push([cs, ce]);
  return out;
}

function subtractIntervals_(winStart, winEnd, busies) {
  const free = [];
  let cur = new Date(winStart);
  for (const [bs, be] of busies) {
    if (cur < bs) free.push([new Date(cur), new Date(bs)]);
    if (cur < be) cur = new Date(be);
  }
  if (cur < winEnd) free.push([new Date(cur), new Date(winEnd)]);
  return free;
}

/* Dates and formatting */
function todayDate_() {
  return parseLocal_(Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'00:00:00"));
}
function setTime_(dateOnly, h, m) {
  const ds = Utilities.formatDate(dateOnly, TIMEZONE, "yyyy-MM-dd");
  return parseLocal_(`${ds}T${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}:00`);
}
function parseLocal_(iso) { return new Date(iso); }
function isSameYmd_(a, b) {
  const fa = Utilities.formatDate(a, TIMEZONE, "yyyy-MM-dd");
  const fb = Utilities.formatDate(b, TIMEZONE, "yyyy-MM-dd");
  return fa === fb;
}
function ceilToNextHalfHour_(d) {
  const out = new Date(d);
  out.setSeconds(0); out.setMilliseconds(0);
  const m = out.getMinutes();
  const add = (30 - (m % 30)) % 30;
  out.setMinutes(m + add);
  return out;
}

function formatDayLabel_(d) {
  const dayName = Utilities.formatDate(d, TIMEZONE, "EEEE");
  const monthAbbrev = Utilities.formatDate(d, TIMEZONE, "MMM");
  const dayNum = Utilities.formatDate(d, TIMEZONE, "d");
  return `${dayName}, ${monthAbbrev} ${dayNum}`;
}
function intervalsToInlineText_(intervals) {
  return intervals.map(([s, e]) => `${formatTime_(s)} to ${formatTime_(e)}`).join(", ");
}
function formatTime_(dt) {
  const h24 = parseInt(Utilities.formatDate(dt, TIMEZONE, "H"), 10);
  const m   = Utilities.formatDate(dt, TIMEZONE, "mm");
  const h12 = h24 % 12 === 0 ? 12 : h24 % 12;
  return m === "00" ? `${h12}` : `${h12}:${m}`;
}
