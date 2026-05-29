# 📅 Automated Google Calendar Timesheet Sync

A modular, production-ready **Google Apps Script** that automates student or freelancer work hour tracking. It acts like a heat-seeking missile: scanning a dedicated Google Calendar for custom tags, processing shifts defensively (including automatic lunch break calculations), and injecting clean data straight into pre-built weekly blocks in a Google Sheet.

Built with strict adherence to **SOLID** software engineering principles and the **KISS** philosophy.

---

## ✨ Features

* **Visual Calendar Workflow:** No manual spreadsheet data entry. Just drag-and-drop your shifts onto your calendar, tag them with `#work`, and let the automation do the heavy lifting.
* **Smart Break Calculation:** Automatically merges up to two separate shifts per day (e.g., morning and afternoon blocks) and calculates the exact duration of your unpaid break/pause.
* **Defensive Architecture:**
  * Explicitly flags and warns you about problematic cross-midnight shifts.
  * Protects against silent data corruption by throwing loud errors if duplicate dates are accidentally introduced into your sheet template.
  * Handles floating-point arithmetic errors natively (guaranteeing clean 2-decimal rounding).
* **High Efficiency:** Utilizes batched array writes to minimize slow Google Apps Script API calls, drastically lowering execution times and cloud quota consumption.
* **Stale-Data Resolution:** Option to automatically clear sheet data if a calendar event is deleted, moved, or untagged retroactively.

---

## 🛠️ Code Architecture (SOLID & KISS)

Unlike monolithic script examples, this repository breaks down tasks into isolated, single-responsibility layers:

1. **Configuration (`CONFIG`):** Open for extension, closed for modification. Adjust columns, calendar IDs, lookback windows, and tracking tags seamlessly in one central object.
2. **Read Layer (`mapSheetDatesToRows`):** Scans the spreadsheet canvas and builds a fast text-matching memory map of row destinations.
3. **Fetch Layer (`fetchCalendarEvents`):** Handles raw interactions with the Google Calendar API and enforces strict tag matching via tokenization (`#workshop` won't trigger `#work`).
4. **Logic Layer (`calculatePause` & `buildEventContext`):** Pure functions dedicated strictly to string manipulation (titles + descriptions) and chronological date math.
5. **Write Layer (`writeRow`):** Executes high-performance batched range injections into the Google Sheets UI.

---

## 🚀 Quick Setup

### 1. Prepare your Google Sheet
Ensure your tracking tab is named `Timesheet` and has the following column layout:
* **Column B:** Date (`dd.MM.yyyy`)
* **Column C:** Start Time
* **Column D:** End Time
* **Column E:** Pause (Unpaid Break)
* **Column F:** Total Hours
* **Column G:** Context/Description

### 2. Inject the Script
1. Inside your Google Sheet, navigate to **Extensions > Apps Script**.
2. Delete any default code and paste the contents of `syncWorkHours.gs` from this repository.
3. Replace the `calendarId` string in the `CONFIG` block with your actual target Google Calendar ID.
4. Click **Save** (Floppy Disk icon).

### 3. Deploy the Weekly Trigger
1. Click the **Triggers** icon ⏰ (alarm clock) in the Apps Script sidebar.
2. Click **+ Add Trigger** in the bottom right.
3. Choose `syncWorkHours` as the function to run.
4. Set the event source to **Time-driven** -> **Week timer**.
5. Select your preferred day and time (e.g., Every Friday evening) and hit **Save**.

---

## 📋 Data Contract

* **Start/End Fields:** Outputted as plain-text strings formatted to `HH:mm`.
* **Pause/Total Fields:** Outputted as native floating-point decimal hours (e.g., `1.5` equals 1 hour and 30 minutes) making your weekly summing formulas incredibly clean.
* **Context Field:** Dynamically combines cleaned calendar titles with event descriptions, separated by newlines `\n` if multiple blocks occur on the same day.
