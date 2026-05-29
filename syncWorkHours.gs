/**
 * ==========================================
 * syncWorkHours.gs
 * ==========================================
 * Google Apps Script that synchronizes calendar events tagged with #work
 * into a Google Sheets timesheet.
 *
 * Design principles:
 *   - SOLID: Each function has a single responsibility.
 *   - Defensive: Guards against null descriptions, duplicate dates, cross-midnight events.
 *   - Consistent: All column references are 1-based (Google Sheets standard).
 *   - Efficient: Batched writes to minimize Apps Script API calls.
 *
 * Data contract:
 *   - Start/End columns: text strings in "HH:mm" format.
 *   - Pause/Total columns: decimal hours (e.g. 1.5 = 1h 30m).
 *   - Context column: plain text (newline-separated if multiple blocks).
 */

// ==========================================
// CONFIGURATION
// ==========================================

const CONFIG = {
  sheetName: "Timesheet",
  calendarId: "c_535915790b75465c54a9b5a9416794d9210eb81aa97c20ca7d4c0544b12da8b3@group.calendar.google.com",
  tags: ["#work"],
  // "daysToLookBack: 7" means: today + the previous 6 days = 7 calendar days total.
  daysToLookBack: 7,
  // When true, the script clears columns C–G for dates in the window that have no matching events.
  // Set to false if those columns are sometimes edited manually.
  clearStaleDates: true,
  // ALL column numbers are 1-based (Google Sheets standard).
  columns: {
    date: 2,      // Column B
    start: 3,     // Column C
    end: 4,       // Column D
    pause: 5,     // Column E
    total: 6,     // Column F
    context: 7    // Column G
  }
};

// ==========================================
// ORCHESTRATOR
// ==========================================

function syncWorkHours() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetName);
  if (!sheet) throw new Error("Timesheet tab not found. Check CONFIG.sheetName!");

  const dateRowMap = mapSheetDatesToRows(sheet);
  const { startDate, endDate } = buildDateWindow(CONFIG.daysToLookBack);
  const groupedEvents = fetchCalendarEvents(startDate, endDate, CONFIG.tags);

  updateSheetWithEvents(sheet, dateRowMap, groupedEvents, startDate, endDate);
}

// ==========================================
// DATE WINDOW
// ==========================================

/**
 * Builds the inclusive date window for the sync.
 * Returns { startDate, endDate } where:
 *   - startDate: 00:00:00 of (today - daysToLookBack + 1)
 *   - endDate:   23:59:59 of today
 */
function buildDateWindow(daysToLookBack) {
  const now = new Date();

  const endDate = new Date(now);
  endDate.setHours(23, 59, 59, 999);

  const startDate = new Date(now);
  startDate.setDate(now.getDate() - (daysToLookBack - 1));
  startDate.setHours(0, 0, 0, 0);

  return { startDate, endDate };
}

// ==========================================
// RESPONSIBILITY 1: Read the Sheet
// ==========================================

/**
 * Creates a map of "dd.MM.yyyy" → row number.
 * Only includes cells that are actual Date objects or match the dd.MM.yyyy pattern.
 * Throws on duplicate dates to prevent silent data corruption.
 */
function mapSheetDatesToRows(sheet) {
  const data = sheet.getDataRange().getValues();
  const dateRowMap = {};
  const datePattern = /^\d{2}\.\d{2}\.\d{4}$/;

  data.forEach((row, index) => {
    const dateCell = row[CONFIG.columns.date - 1]; // Convert 1-based config to 0-based array
    if (!dateCell) return;

    let dateStr;
    if (dateCell instanceof Date) {
      dateStr = Utilities.formatDate(dateCell, Session.getScriptTimeZone(), "dd.MM.yyyy");
    } else {
      const text = dateCell.toString().trim();
      if (!datePattern.test(text)) return; // Skip headers and non-date text
      dateStr = text;
    }

    if (dateRowMap[dateStr]) {
      throw new Error(
        `Duplicate date "${dateStr}" found at rows ${dateRowMap[dateStr]} and ${index + 1}. Fix the template.`
      );
    }

    dateRowMap[dateStr] = index + 1;
  });

  return dateRowMap;
}

// ==========================================
// RESPONSIBILITY 2: Fetch and Group Calendar Events
// ==========================================

/**
 * Fetches events from the configured calendar within the given window.
 * Groups them by date string. Initializes all dates in the window so
 * the caller can detect days that should be cleared.
 *
 * Cross-midnight events are skipped with a console warning because they
 * cannot be unambiguously assigned to a single date.
 */
function fetchCalendarEvents(startDate, endDate, tags) {
  const calendar = CalendarApp.getCalendarById(CONFIG.calendarId);
  if (!calendar) throw new Error("Calendar not found. Check CONFIG.calendarId!");

  const rawEvents = calendar.getEvents(startDate, endDate);
  const eventsByDate = {};
  const tz = Session.getScriptTimeZone();

  // Pre-populate every date in the window with an empty array
  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    eventsByDate[Utilities.formatDate(d, tz, "dd.MM.yyyy")] = [];
  }

  rawEvents.forEach(event => {
    const title = event.getTitle();
    if (!matchesTag(title, tags)) return;

    const eventStart = event.getStartTime();
    const eventEnd = event.getEndTime();

    // Reject cross-midnight events
    if (!isSameDay(eventStart, eventEnd)) {
      console.warn(
        `Skipping cross-midnight event "${title}" (${eventStart.toISOString()} → ${eventEnd.toISOString()}). ` +
        `Split it into separate days in your calendar.`
      );
      return;
    }

    const dateStr = Utilities.formatDate(eventStart, tz, "dd.MM.yyyy");
    if (!eventsByDate[dateStr]) eventsByDate[dateStr] = [];
    eventsByDate[dateStr].push(event);
  });

  return eventsByDate;
}

// ==========================================
// RESPONSIBILITY 3: Process and Update Sheet
// ==========================================

/**
 * Iterates over every date in the event map and writes computed values to the sheet.
 * For dates with no events, clears the row (if CONFIG.clearStaleDates is true).
 */
function updateSheetWithEvents(sheet, dateRowMap, eventsByDate, startDate, endDate) {
  for (const dateStr in eventsByDate) {
    const targetRow = dateRowMap[dateStr];
    if (!targetRow) continue;

    const dayEvents = eventsByDate[dateStr];

    if (dayEvents.length === 0) {
      if (CONFIG.clearStaleDates) {
        clearRow(sheet, targetRow);
      }
      continue;
    }

    dayEvents.sort((a, b) => a.getStartTime() - b.getStartTime());

    const startTime = dayEvents[0].getStartTime();
    const endTime = dayEvents[dayEvents.length - 1].getEndTime();

    let totalWorkedHours = 0;
    const contexts = [];

    dayEvents.forEach(ev => {
      totalWorkedHours += (ev.getEndTime() - ev.getStartTime()) / (1000 * 60 * 60);
      const ctx = buildEventContext(ev, CONFIG.tags);
      if (ctx) contexts.push(ctx);
    });

    const pauseHours = calculatePause(dayEvents.length, startTime, endTime, totalWorkedHours);
    const tz = Session.getScriptTimeZone();
    const startStr = Utilities.formatDate(startTime, tz, "HH:mm");
    const endStr = Utilities.formatDate(endTime, tz, "HH:mm");
    const contextStr = contexts.join("\n");

    writeRow(sheet, targetRow, startStr, endStr, pauseHours, totalWorkedHours, contextStr);
  }
}

// ==========================================
// HELPER FUNCTIONS
// ==========================================

/**
 * Returns true if the event title contains any of the configured tags
 * as a distinct token (handles adjacent punctuation like "#work," or "(#work)").
 */
function matchesTag(title, tags) {
  const lowerTitle = title.toLowerCase();
  return tags.some(tag => {
    const lowerTag = tag.toLowerCase();
    const escaped = escapeRegex(lowerTag);
    // Matches the tag when preceded by start-of-string or whitespace/punctuation,
    // and followed by end-of-string, whitespace, or punctuation.
    const regex = new RegExp('(?:^|[\\s(\\[{])' + escaped + '(?=[\\s,;.!?)\\]}]|$)', 'i');
    return regex.test(lowerTitle);
  });
}

/**
 * Builds the context string from event title (minus tags) and description.
 */
function buildEventContext(event, tags) {
  let cleanTitle = event.getTitle();

  tags.forEach(tag => {
    const escaped = escapeRegex(tag);
    // Remove tag as a standalone token (with optional surrounding punctuation/whitespace)
    const regex = new RegExp('(?:^|\\s)' + escaped + '(?=[\\s,;.!?)\\]|]|$)', 'gi');
    cleanTitle = cleanTitle.replace(regex, '');
  });
  cleanTitle = cleanTitle.trim();

  const rawDesc = event.getDescription();
  const description = rawDesc ? rawDesc.trim() : "";

  if (cleanTitle && description) return cleanTitle + " - " + description;
  if (cleanTitle) return cleanTitle;
  if (description) return description;
  return "";
}

/**
 * Calculates total pause (gap) hours between work blocks.
 * Works for any number of blocks ≥ 2. Returns 0 for a single block.
 */
function calculatePause(blockCount, startTime, endTime, totalWorkedHours) {
  if (blockCount >= 2) {
    const totalSpanHours = (endTime - startTime) / (1000 * 60 * 60);
    return Math.round((totalSpanHours - totalWorkedHours) * 100) / 100; // Round to 2 decimals
  }
  return 0;
}

/**
 * Writes a single row of timesheet data using a batched setValues call.
 */
function writeRow(sheet, targetRow, startStr, endStr, pauseHours, totalHours, contextStr) {
  sheet.getRange(targetRow, CONFIG.columns.start, 1, 5)
    .setValues([[startStr, endStr, pauseHours, totalHours, contextStr]]);
}

/**
 * Clears all sync-managed columns (C–G) for a given row.
 */
function clearRow(sheet, targetRow) {
  sheet.getRange(targetRow, CONFIG.columns.start, 1, 5).clearContent();
}

/**
 * Returns true if two Date objects fall on the same calendar day (in the script timezone).
 */
function isSameDay(date1, date2) {
  const tz = Session.getScriptTimeZone();
  const d1 = Utilities.formatDate(date1, tz, "yyyyMMdd");
  const d2 = Utilities.formatDate(date2, tz, "yyyyMMdd");
  return d1 === d2;
}

/**
 * Escapes special regex characters in a string.
 */
function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
