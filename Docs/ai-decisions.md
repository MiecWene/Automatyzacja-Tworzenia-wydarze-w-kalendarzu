# AI Decisions Log

This file contains architectural and implementation
decisions made during development with AI assistance.

Cursor should read this file before suggesting changes.

---

## Entry Template

Date: 09.03.2026

Context: Fix for incorrect days for events 

AI suggestion: 

Goal

Ensure that recurring Google Calendar events created by Automatyzacja_1.0 occur on the correct weekdays and within the intended date range as specified in the Test_Schedule_Month_Script sheet.

Current behavior analysis

Sheet integration: The script reads all rows from the active spreadsheet, specifically the Test_Schedule_Month_Script tab, via readScheduleRows and maps columns like Data startu, Data końca, Dzień #1, Godzina #1 (hh:mm), etc. into a validated row model.

Date & weekday parsing:

Data startu and Data końca are parsed by parseDateCell, which returns a Date with whatever the sheet stores (typically midnight in the spreadsheet timezone).

Dzień #1 / Dzień #2 are parsed by parseWeekdayToken, which currently only accepts English short names Mon, Tue, Wed, Thu, Fri.

Time-of-day is parsed separately by parseTimeOfDay.

Schedule model: buildGroupSchedule combines startDate, endDate, and one or two weekly slots into a schedule object attached to each group bucket.

function createSeries(calendar, bucket, emails, description, reportEntries) {
  var schedule = bucket.schedule;
  var primarySlot = schedule.slots[0];

  var start = new Date(schedule.startDate.getTime());
  start.setHours(primarySlot.time.hour, primarySlot.time.minute, 0, 0);
  var end = new Date(start.getTime() + 60 * 60 * 1000); // 1 hour

  var recurrence = CalendarApp.newRecurrence()
    .addWeeklyRule()
    .until(schedule.endDate);

  // Add all weekdays in slots
  schedule.slots.forEach(function(s) {
    recurrence.onlyOnWeekday(toAppsScriptWeekday(s.weekday));
  });

Recurring series creation:

For each group, createSeries uses schedule.startDate as-is for the first event occurrence and adds weekly recurrence rules with .onlyOnWeekday(...) for all slot weekdays.

The same assumption is duplicated in updateSeries, which rebuilds the recurrence starting from schedule.startDate.

Likely mismatch:

If Data startu in the sheet is treated by humans as a generic "start of the cycle" date (e.g. Monday of the first teaching week) while Dzień #1 is another weekday (e.g. Tuesday), the script currently uses the Monday date as the first occurrence, causing the event to appear on the wrong day compared to what the user expects.

Even when Data startu is usually the first class date, the more robust behavior is to always align the actual first event date to the first occurrence of the configured weekday on or after Data startu.

Proposed solution

Introduce weekday alignment helper:

Add a small helper alignDateToWeekday(baseDate, weekdayToken) that returns the first date on or after baseDate whose weekday matches the given Dzień #1 token (Mon, Tue, etc.).

Internally, map tokens to JS getDay() integers and advance the date by 0–6 days until they match.

Use aligned start in both series operations:

In createSeries, instead of using schedule.startDate directly, compute:

var alignedStart = alignDateToWeekday(schedule.startDate, primarySlot.weekday);

Set event start/end from alignedStart plus the configured time and duration.

In updateSeries, apply the same logic so that when the schedule changes or events are re-synced, the series is still anchored to the correct weekday.

Keep recurrence rules unchanged:

Continue to use CalendarApp.newRecurrence().addWeeklyRule().until(schedule.endDate) and .onlyOnWeekday(...) for all configured slots; the main fix is to ensure the anchor date matches the primary weekday.

Guardrails and assumptions:

If Data startu already falls on the correct Dzień #1 weekday, alignDateToWeekday will return the same date, so current behavior is preserved.

If Data startu falls on a weekend or a different weekday, the first occurrence will shift forward up to 6 days to the correct weekday, which should match how the Dzień #1 column is interpreted by users.

Implementation outline

Step 1 – Add alignment helper

Implement alignDateToWeekday(baseDate, weekdayToken) in Automatyzacja_1.0 near other helper functions (e.g. close to toAppsScriptWeekday).

Use a simple mapping Mon -> 1, Tue -> 2, ..., Fri -> 5 and standard JS Date arithmetic.

Concrete helper implementations to use:

function jsWeekdayFromToken(token) {
  // JS getDay(): Sun=0, Mon=1, ... Sat=6
  switch (token) {
    case 'Mon': return 1;
    case 'Tue': return 2;
    case 'Wed': return 3;
    case 'Thu': return 4;
    case 'Fri': return 5;
    default:
      throw new Error('Unsupported weekday token for alignment: ' + token);
  }
}

function alignDateToWeekday(baseDate, weekdayToken) {
  var targetDow = jsWeekdayFromToken(weekdayToken);
  var d = new Date(baseDate.getTime());
  var currentDow = d.getDay(); // 0–6
  var delta = (targetDow - currentDow + 7) % 7; // 0–6 days forward
  d.setDate(d.getDate() + delta);
  return d;
}

Step 2 – Update createSeries

Replace direct use of schedule.startDate with the aligned date:

Compute alignedStart using the helper and then set hours/minutes from primarySlot.time.

Leave the until(schedule.endDate) and .onlyOnWeekday(...) calls as they are.

Concrete change:

var schedule = bucket.schedule;
var primarySlot = schedule.slots[0];

var alignedStartDate = alignDateToWeekday(schedule.startDate, primarySlot.weekday);
var start = new Date(alignedStartDate.getTime());
start.setHours(primarySlot.time.hour, primarySlot.time.minute, 0, 0);
var end = new Date(start.getTime() + 60 * 60 * 1000); // 1 hour

Step 3 – Update updateSeries

Mirror the same change inside updateSeries when computing start before series.setTime(start, end) and series.setRecurrence(recurrence).

Concrete change:

var schedule = bucket.schedule;
var primarySlot = schedule.slots[0];

var alignedStartDate = alignDateToWeekday(schedule.startDate, primarySlot.weekday);
var start = new Date(alignedStartDate.getTime());
start.setHours(primarySlot.time.hour, primarySlot.time.minute, 0, 0);
var end = new Date(start.getTime() + 60 * 60 * 1000);

Step 4 – Sanity-check behavior with test rows

In the Test_Schedule_Month_Script tab, prepare 1–2 sample groups where:

Data startu is a Monday, Dzień #1 is Tuesday (or another weekday), and verify that the first event is now on the configured weekday.

Data startu already matches Dzień #1 and confirm that the behavior is unchanged.

Run runSync and inspect the created/updated series in the calendar to confirm the days match the sheet.

Step 5 – Optional logging

Optionally, for easier future debugging, log a RUN_REPORT entry (or use Logger.log) when alignDateToWeekday actually shifts the date (e.g. "Aligned start date from 2026-03-02 (Mon) to 2026-03-03 (Tue) for group XYZ").

Result

After these changes, the script will no longer rely on Data startu being exactly the same weekday as Dzień #1. Instead, it will reliably create recurring events whose first occurrence and subsequent recurrences fall on the weekdays explicitly configured in the sheet, eliminating the observed mismatch between the Test_Schedule_Month_Script tab and the calendar.



What I changed in the script
Added weekday-alignment helpers just after toAppsScriptWeekday in Automatyzacja_1.0:

Library
Lines 381-398
function jsWeekdayFromToken(token) {
  // JS getDay(): Sun=0, Mon=1, ... Sat=6
  switch (token) {
    case 'Mon': return 1;
    case 'Tue': return 2;
    case 'Wed': return 3;
    case 'Thu': return 4;
    case 'Fri': return 5;
    default:
      throw new Error('Unsupported weekday token for alignment: ' + token);
  }
}
function alignDateToWeekday(baseDate, weekdayToken) {
  var targetDow = jsWeekdayFromToken(weekdayToken);
  var d = new Date(baseDate.getTime());
  var currentDow = d.getDay(); // 0–6
  var delta = (targetDow - currentDow + 7) % 7; // 0–6 days forward
  d.setDate(d.getDate() + delta);
  return d;
}
Updated createSeries to anchor the first occurrence on the first correct weekday on/after Data startu:

Library
Lines 547-555
function createSeries(calendar, bucket, emails, description, reportEntries) {
  var schedule = bucket.schedule;
  var primarySlot = schedule.slots[0];
  var alignedStartDate = alignDateToWeekday(schedule.startDate, primarySlot.weekday);
  var start = new Date(alignedStartDate.getTime());
  start.setHours(primarySlot.time.hour, primarySlot.time.minute, 0, 0);
  var end = new Date(start.getTime() + 60 * 60 * 1000); // 1 hour
Updated updateSeries to use the same alignment when resetting an existing series:

Library
Lines 584-590
  // Update recurrence if schedule changed: for simplicity, we always reset
  var schedule = bucket.schedule;
  var primarySlot = schedule.slots[0];
  var alignedStartDate = alignDateToWeekday(schedule.startDate, primarySlot.weekday);
  var start = new Date(alignedStartDate.getTime());
  start.setHours(primarySlot.time.hour, primarySlot.time.minute, 0, 0);
  var end = new Date(start.getTime() + 60 * 60 * 1000);
How to verify in your sheet
In Test_Schedule_Month_Script, pick a test group where:
Data startu is, for example, a Monday,
Dzień #1 is Tue (or another weekday),
Data końca is a few weeks later, and at least one confirmed participant exists.
Run Schedule Sync → Run sync now (or runSync) and check the calendar:
The first event for that group should now land on the first Dzień #1 weekday on or after Data startu, and all subsequent recurrences should follow that weekday pattern within the Data startu–Data końca range.

Why Tuesday 31.03 is still missing

The script already passes endOfDay(schedule.endDate) (31.03 23:59:59.999) to .until(). Events on 31.03 are still not created. Likely causes:





UNTIL vs event start comparison

RFC 5545 and the Apps Script docs require UNTIL to have the same type as DTSTART. The event has a datetime start (e.g. 31.03 10:00). If the backend normalizes UNTIL to date-only (31.03 = 31.03 00:00:00), then 31.03 10:00 can be treated as after UNTIL and excluded.



Timezone

endOfDay() uses setHours(23, 59, 59, 999) in the script’s timezone. The calendar (and event times) may use another timezone. After conversion, “end of 31.03” in one zone can fall on 01.04 or earlier on 31.03 in another, so the last occurrence can be excluded or the boundary can be wrong.



Implementation treating UNTIL as exclusive

Some implementations treat the UNTIL instant as exclusive; then “until 31.03 23:59:59” could exclude that exact moment or the day.

Fix: use start of next day for UNTIL

Use the start of the day after Data końca as the UNTIL value (e.g. 01.04 00:00:00):





Any occurrence on 31.03 (at any time) is strictly before 01.04 00:00:00, so it is included whether the backend treats UNTIL as inclusive or exclusive.



We no longer depend on “end of day” in a specific timezone; we only add one calendar day and set time to midnight, which is robust across zones for “include the whole end date”.

Implementation in Automatyzacja_1.0





Add startOfNextDay(date)





Clone date, add one calendar day with setDate(d.getDate() + 1), then set time to 00:00:00.000 with setHours(0, 0, 0, 0). Return the new Date.  



Place it next to endOfDay (e.g. after the PARSERS / date helpers section).



Use it for recurrence UNTIL





In createSeries: replace .until(endOfDay(schedule.endDate)) with .until(startOfNextDay(schedule.endDate)).  



In updateSeries: same change, .until(startOfNextDay(schedule.endDate)).



Keep or remove endOfDay





endOfDay is no longer used for recurrence. Either remove it or leave it for possible future use; leaving it is harmless.

Result

With Data końca = 31.03, UNTIL becomes 01.04 00:00:00. All occurrences on 31.03 (including Tuesday 31.03 at 10:00) are before that instant and will be included, so events will be created on every Tuesday in March, including 31.03.

Goal

Ensure that recurring Google Calendar events created by Automatyzacja_1.0 occur on the correct weekdays and within the intended date range as specified in the Test_Schedule_Month_Script sheet.

Current behavior analysis





Sheet integration: The script reads all rows from the active spreadsheet, specifically the Test_Schedule_Month_Script tab, via readScheduleRows and maps columns like Data startu, Data końca, Dzień #1, Godzina #1 (hh:mm), etc. into a validated row model.



Date & weekday parsing:





Data startu and Data końca are parsed by parseDateCell, which returns a Date with whatever the sheet stores (typically midnight in the spreadsheet timezone).



Dzień #1 / Dzień #2 are parsed by parseWeekdayToken, which currently only accepts English short names Mon, Tue, Wed, Thu, Fri.



Time-of-day is parsed separately by parseTimeOfDay.



Schedule model: buildGroupSchedule combines startDate, endDate, and one or two weekly slots into a schedule object attached to each group bucket.

function createSeries(calendar, bucket, emails, description, reportEntries) {
  var schedule = bucket.schedule;
  var primarySlot = schedule.slots[0];

  var start = new Date(schedule.startDate.getTime());
  start.setHours(primarySlot.time.hour, primarySlot.time.minute, 0, 0);
  var end = new Date(start.getTime() + 60 * 60 * 1000); // 1 hour

  var recurrence = CalendarApp.newRecurrence()
    .addWeeklyRule()
    .until(schedule.endDate);

  // Add all weekdays in slots
  schedule.slots.forEach(function(s) {
    recurrence.onlyOnWeekday(toAppsScriptWeekday(s.weekday));
  });





Recurring series creation:





For each group, createSeries uses schedule.startDate as-is for the first event occurrence and adds weekly recurrence rules with .onlyOnWeekday(...) for all slot weekdays.



The same assumption is duplicated in updateSeries, which rebuilds the recurrence starting from schedule.startDate.



Likely mismatch:





If Data startu in the sheet is treated by humans as a generic "start of the cycle" date (e.g. Monday of the first teaching week) while Dzień #1 is another weekday (e.g. Tuesday), the script currently uses the Monday date as the first occurrence, causing the event to appear on the wrong day compared to what the user expects.



Even when Data startu is usually the first class date, the more robust behavior is to always align the actual first event date to the first occurrence of the configured weekday on or after Data startu.

Proposed solution





Introduce weekday alignment helper:





Add a small helper alignDateToWeekday(baseDate, weekdayToken) that returns the first date on or after baseDate whose weekday matches the given Dzień #1 token (Mon, Tue, etc.).



Internally, map tokens to JS getDay() integers and advance the date by 0–6 days until they match.



Use aligned start in both series operations:





In createSeries, instead of using schedule.startDate directly, compute:





var alignedStart = alignDateToWeekday(schedule.startDate, primarySlot.weekday);



Set event start/end from alignedStart plus the configured time and duration.



In updateSeries, apply the same logic so that when the schedule changes or events are re-synced, the series is still anchored to the correct weekday.



Keep recurrence rules unchanged:





Continue to use CalendarApp.newRecurrence().addWeeklyRule().until(schedule.endDate) and .onlyOnWeekday(...) for all configured slots; the main fix is to ensure the anchor date matches the primary weekday.



Guardrails and assumptions:





If Data startu already falls on the correct Dzień #1 weekday, alignDateToWeekday will return the same date, so current behavior is preserved.



If Data startu falls on a weekend or a different weekday, the first occurrence will shift forward up to 6 days to the correct weekday, which should match how the Dzień #1 column is interpreted by users.

Implementation outline





Step 1 – Add alignment helper





Implement alignDateToWeekday(baseDate, weekdayToken) in Automatyzacja_1.0 near other helper functions (e.g. close to toAppsScriptWeekday).



Use a simple mapping Mon -> 1, Tue -> 2, ..., Fri -> 5 and standard JS Date arithmetic.



Concrete helper implementations to use:

function jsWeekdayFromToken(token) {
  // JS getDay(): Sun=0, Mon=1, ... Sat=6
  switch (token) {
    case 'Mon': return 1;
    case 'Tue': return 2;
    case 'Wed': return 3;
    case 'Thu': return 4;
    case 'Fri': return 5;
    default:
      throw new Error('Unsupported weekday token for alignment: ' + token);
  }
}

function alignDateToWeekday(baseDate, weekdayToken) {
  var targetDow = jsWeekdayFromToken(weekdayToken);
  var d = new Date(baseDate.getTime());
  var currentDow = d.getDay(); // 0–6
  var delta = (targetDow - currentDow + 7) % 7; // 0–6 days forward
  d.setDate(d.getDate() + delta);
  return d;
}





Step 2 – Update createSeries





Replace direct use of schedule.startDate with the aligned date:





Compute alignedStart using the helper and then set hours/minutes from primarySlot.time.



Leave the until(schedule.endDate) and .onlyOnWeekday(...) calls as they are.



Concrete change:

var schedule = bucket.schedule;
var primarySlot = schedule.slots[0];

var alignedStartDate = alignDateToWeekday(schedule.startDate, primarySlot.weekday);
var start = new Date(alignedStartDate.getTime());
start.setHours(primarySlot.time.hour, primarySlot.time.minute, 0, 0);
var end = new Date(start.getTime() + 60 * 60 * 1000); // 1 hour





Step 3 – Update updateSeries





Mirror the same change inside updateSeries when computing start before series.setTime(start, end) and series.setRecurrence(recurrence).



Concrete change:

var schedule = bucket.schedule;
var primarySlot = schedule.slots[0];

var alignedStartDate = alignDateToWeekday(schedule.startDate, primarySlot.weekday);
var start = new Date(alignedStartDate.getTime());
start.setHours(primarySlot.time.hour, primarySlot.time.minute, 0, 0);
var end = new Date(start.getTime() + 60 * 60 * 1000);





Step 4 – Sanity-check behavior with test rows





In the Test_Schedule_Month_Script tab, prepare 1–2 sample groups where:





Data startu is a Monday, Dzień #1 is Tuesday (or another weekday), and verify that the first event is now on the configured weekday.



Data startu already matches Dzień #1 and confirm that the behavior is unchanged.



Run runSync and inspect the created/updated series in the calendar to confirm the days match the sheet.



Step 5 – Optional logging





Optionally, for easier future debugging, log a RUN_REPORT entry (or use Logger.log) when alignDateToWeekday actually shifts the date (e.g. "Aligned start date from 2026-03-02 (Mon) to 2026-03-03 (Tue) for group XYZ").

Result

After these changes, the script will no longer rely on Data startu being exactly the same weekday as Dzień #1. Instead, it will reliably create recurring events whose first occurrence and subsequent recurrences fall on the weekdays explicitly configured in the sheet, eliminating the observed mismatch between the Test_Schedule_Month_Script tab and the calendar.


Root cause





Data końca (e.g. 31.03) is read from the sheet and parsed by parseDateCell as a Date at midnight (00:00:00) on that day.



The script passes that value directly to CalendarApp.newRecurrence().until(schedule.endDate) in both createSeries (line 580) and updateSeries (line 616).



Recurrence "until" is interpreted as "include occurrences whose start is on or before this instant". So "until 31.03 00:00:00" excludes any event that starts later on 31.03 (e.g. 10:00). That is why the Tuesday 31.03 event is missing.

Fix

Use the end of the end date (same day at 23:59:59.999) when calling .until(), so the whole calendar day is included.





Add a small helper that returns end-of-day for a given date (no change to the date’s calendar day, only time set to 23:59:59.999).



In createSeries and updateSeries, pass that end-of-day value to .until() instead of schedule.endDate.

Implementation





Add helper (e.g. next to other date helpers in Automatyzacja_1.0):





endOfDay(date) – clone the date and set hours to 23, minutes to 59, seconds to 59, ms to 999; return the new Date.



createSeries (around line 577–580):





Replace .until(schedule.endDate) with .until(endOfDay(schedule.endDate)).



updateSeries (around line 613–616):





Same change: .until(endOfDay(schedule.endDate)).

No other logic needs to change (e.g. findExistingGroupSeries and trial-date checks already use the same schedule.endDate for range semantics; making the recurrence include the last day does not conflict with them).

Result

With Data końca = 31.03, recurrence will run "until 31.03 23:59:59.999", so the Tuesday 31.03 event (e.g. at 10:00) will be included and events will be created on all Tuesdays in March, including the 31st.

Files affected: Automatyzacja_1.0

Next steps: Builded the suggested config

---

## Entry – 10.03.2026

**Context**: Per-slot scheduling for Dzień #1 / Dzień #2, trial lessons attached to the correct weekly series, and a dedicated deletion tool for group events.

### Scheduling fixes (Dzień #1 / Dzień #2, inclusive end dates)

- **Per-slot series instead of one combined series**  
  - Decision: a single recurring series in Google Calendar can’t have different times on different weekdays, so each group (`Nazwa Grupy`) is modeled as up to **two separate series**:
    - One for `Dzień #1` / `Godzina #1`.
    - One for `Dzień #2` / `Godzina #2` (when present).
  - Implementation in `Automatyzacja_1.0`:
    - `buildGroupSchedule(row)` keeps `slots: [{ weekday: day1, time: time1 }, { weekday: day2, time: time2 }]` when Dzień #2 / Godzina #2 are valid.
    - `createSeries(calendar, bucket, slot, emails, description, reportEntries)` now takes a **single `slot`**, aligns the first occurrence with `alignDateToWeekday(schedule.startDate, slot.weekday)`, and creates a weekly recurrence with `.onlyOnWeekday(toAppsScriptWeekday(slot.weekday))`.
    - `updateSeries(calendar, series, bucket, slot, emails, description, reportEntries)` mirrors this logic for updates.
    - `upsertEventsForGroups` loops over `bucket.schedule.slots` and creates/updates one series per slot for the same group name.

- **End-date inclusivity**  
  - Decision: events on `Data końca` must be included for all recurrences.
  - Implementation:
    - Introduced `endOfDay(date)` early on, then later switched to `startOfNextDay(date)` and used that for `.until(...)` on the recurrence:
      - Recurrence now ends at the **start of the day after** `Data końca`, so all events on `Data końca` are included regardless of timezone/UTC normalization.

- **Robust parsing for headers, weekdays and times**  
  - Decision: tolerate Sheet quirks (non-breaking spaces in headers, Polish abbreviations, localized times).
  - Implementation:
    - `validateRow`’s header lookup now normalizes whitespace in header names (e.g. `Dzień #2` with NBSP vs `Dzień #2`).
    - `parseWeekdayToken`:
      - Accepts English short names (`Mon`, `Tue`, `Wed`, `Thu`, `Fri`, case-insensitive, with optional `.`).
      - Accepts Polish short forms (`Pn`, `Wt`, `Śr`/`Sr`, `Cz`, `Pt`) and maps them onto the same tokens.
    - `parseTimeOfDay`:
      - Accepts time cells and numeric fractions.
      - Accepts strings `H:MM`, `HH:MM`, but also `H.MM` / `H,MM` (common localized formats).
    - When choosing a canonical row for a group in `validateAndGroup`, the code now **prefers** a row that has Dzień #2 / Godzina #2 set, so `bucket.schedule.slots` includes both slots whenever any row defines them.

### Trial lessons (Data zajęć próbnych)

- **Problem**: trials were originally added as separate one-off events or to a single series per group, so:
  - They ignored Dzień #2 when computing the occurrence to attach to.
  - They sometimes failed with “No occurrence on … Instance not in recurrence window or series not found”.

- **Per-slot mapping for trials**  
  - Decision: a trial lesson should be attached **only to the series matching its weekday**, per group.
  - Implementation:
    - Introduced `jsDayToToken(jsDay)` to convert `Date.getDay()` (1–5) to `"Mon" … "Fri"`.
    - In `upsertEventsForGroups`:
      - For each `trialParticipants` row, compute `token = jsDayToToken(trialDate.getDay())`.
      - Find the matching `slot` in `bucket.schedule.slots` where `slot.weekday === token`.
        - If none: log `TRIAL_SLOT_NOT_FOUND`.
    - New helper `getRecurringEventIdForSlotByListing(calendarId, groupName, slot, startDate, endDate)`:
      - Uses `Calendar.Events.list` (Advanced Calendar API) with `singleEvents: false` and a date window (schedule range ±7 days).
      - Filters returned recurring events by:
        - `summary === groupName`
        - `recurrence` present
        - The series’ `start` day-of-week matching `slot.weekday`, and, if timed, `start` hour/minute matching `slot.time`.
      - Returns the short `id` for that slot’s series (suitable for `events.instances` / `recurringEventId`).

- **Creating trial exceptions via API**  
  - Decision: instead of patching instances directly, create **exceptions for individual occurrences** using `Events.insert` with `recurringEventId` and `originalStartTime`.
  - Implementation in `addTrialGuestToInstance(calendarId, recurringEventId, trialDate, slot, ...)`:
    - Builds `start`/`end` for the trial date using the slot’s time (using y/m/d from the date and the slot’s hour/minute in the script timezone so calendar days line up).
    - Calls `Calendar.Events.instances(calendarId, recurringEventId, { timeMin, timeMax })` to locate the occurrence for that date.
    - If an instance is found:
      - Reads or initializes `instance.attendees`, appends `{ email: trialEmail }` if not already present.
      - Constructs `exceptionEvent` with:
        - `recurringEventId` = slot’s series id,
        - `originalStartTime` from `instance.originalStartTime` (or the computed `timeMin`/timezone),
        - `start` / `end` from the instance (or `timeMin`/`timeMax`),
        - `summary` = group name, `description` from the instance,
        - `attendees` including the trial email.
      - Calls `Calendar.Events.insert(exceptionEvent, calendarId, { sendUpdates: 'none' })`.
    - Logs:
      - `TRIAL_GUEST_ADDED_TO_INSTANCE` on success.
      - `TRIAL_INSTANCE_NOT_FOUND` or `TRIAL_GUEST_ADD_FAILED` with detailed reasons when something goes wrong.

- **Email/noise control**  
  - Decision: minimize email noise from automation while accepting that guest add/remove via `CalendarApp` may still generate some notifications.
  - Implementation:
    - `createSeries` now uses:
      - `createEventSeries(title, start, end, recurrence, { description, guests: emails.join(','), sendInvites: false })`
      - No more `series.addGuest(...)` in the initial creation call.
    - API calls to create exceptions use `sendUpdates: 'none'` to suppress extra email.

### Deletion tooling (Usuwanie_wydarzen)

- **Goal**: safely delete all events for one or more groups (`Nazwa Grupy`) via a friendly UI, without touching unrelated events.

- **Separate deletion script**  
  - Implemented in `Usuwanie_wydarzen_1.0` (and integrated into the main Apps Script project):
    - Reads schedule data from `Test_Schedule_Month_Script` (using a minimal header-based parser).
    - Builds a `groupInfo` map of `groupName -> { startDate, endDate }` (widest window across rows).

- **HTML dialog and handler**  
  - `deleteEventsByGroup()`:
    - Gathers all group names from the schedule.
    - Shows an `HtmlService`-based dialog with checkboxes for each group name, plus **Select all**, **Clear**, **Delete selected**, **Cancel**.
    - On “Delete selected”, calls `google.script.run.handleDeleteGroups(selectedNames)`.
  - `handleDeleteGroups(selectedGroupNames)`:
    - For each selected group:
      - Computes `startWindow = startDate - buffer`, `endWindow = endDate + buffer`.
      - Calls `calendar.getEvents(startWindow, endWindow)`.
      - For each event with `getTitle() === groupName`:
        - If `isRecurringEvent()`:
          - Retrieves `getEventSeries()` and calls `deleteEventSeries()` once per series id.
        - Else:
          - Calls `deleteEvent()` (covers trial exceptions).
    - Aggregates counts per group: `{ seriesDeletedCount, singleDeletedCount }`.

- **Audit reporting**  
  - `writeDeleteReport(ss, results)`:
    - Writes a `Delete_Report` sheet with:
      - Timestamp, GroupName, StartDate, EndDate, SeriesDeleted, SingleDeleted, Error.
    - Gives a clear audit trail for which events were removed and when.

- **Menu integration**  
  - `onOpen()` in the main Apps Script project now creates a single menu:
    - `Schedule Sync` with:
      - `Run sync now` → `runSync()`
      - `Delete group events…` → `deleteEventsByGroup()`

### Files affected

- `Automatyzacja_1.0` – scheduling logic (per-slot series, inclusive end date), trial handling (per-slot series mapping, Calendar API exceptions, email control), menu integration.
- `Usuwanie_wydarzen_1.0` – standalone deletion tooling (HTML dialog, group-based deletion, `Delete_Report`).

---

## Entry – 14.03.2026

**Context**: The original deletion tool (`Usuwanie_wydarzen_1.0`) worked functionally but was intermittently returning a \"permission denied\" error when invoked via the HtmlService dialog. The goal was to preserve the same deletion semantics while experimenting with different UI shapes and execution contexts to identify a more reliable approach.

### 1. Script-only deletion (Usuwanie_wydarzen_1.1)

- **Goal**: Eliminate HtmlService / `google.script.run` from the path to see if the error was related to the client-side sandbox or cross-context calls, and provide a deterministic, editor-runnable deletion flow.

- **Design**  
  - New file: `Usuwanie_Wydarzen_1.1`.  
  - Introduced a versioned config and helpers (`CALENDAR_ID`, `SCHEDULE_SHEET_NAME`, `DELETE_WINDOW_BUFFER_DAYS`, etc. kept aligned with 1.0).  
  - Split concerns into:
    - `deleteEventsByGroup_1_1()` – **script-only entrypoint**:
      - Reads all groups from `Test_Schedule_Month_Script` via `buildGroupInfo_1_1`.
      - Sorts group names and, if any exist, calls `handleDeleteGroups_1_1(groupNames)` to delete events for **all** groups.
      - Logs progress and completion to `Logger` (no UI).
    - `handleDeleteGroups_1_1(selectedGroupNames)` – **core deletion**:
      - For each group:
        - Builds `startWindow` / `endWindow` as schedule range ± `DELETE_WINDOW_BUFFER_DAYS`.
        - Calls `calendar.getEvents(startWindow, endWindow)`.
        - Deletes:
          - Each recurring series once via `deleteEventSeries()` (deduplicating series by id).
          - All matching single events via `deleteEvent()` (including trial exceptions).
      - Accumulates `{ groupName, startDate, endDate, seriesDeletedCount, singleDeletedCount }` per group.
      - Writes all results via `writeDeleteReport_1_1`, reusing the `Delete_Report` layout.
  - A helper `authorizeDeleteTool_1_1()` touches Spreadsheet + Calendar once to ensure all scopes are granted before running deletion.

- **Result / lessons**  
  - This path removed HtmlService entirely; any residual \"permission denied\" would more clearly point to project-level scopes or Calendar ACLs rather than the dialog plumbing.
  - It also provided a safe, non-interactive batch delete mechanism that can be invoked directly from the Apps Script editor when needed.

### 2. Simple prompt UI (Usuwanie_wydarzen_1.1 with UI)

- **Goal**: Add a lightweight way to delete **only selected groups**, still without HtmlService, using spreadsheet-native UI.

- **Design**  
  - Function `deleteEventsByGroup_1_1_withUi()`:
    - Uses `SpreadsheetApp.getUi().prompt` instead of HtmlService:
      - Shows a short preview of available group names (first N) and asks the user to input either:
        - A comma-separated list of group names, or
        - `*` to delete all groups.
    - Parses and validates the input:
      - Splits by comma, trims, filters non-empty names.
      - Keeps only names present in `groupInfo_1_1` (ignores unknown names).
      - If the resulting list is empty, alerts the user and exits.
    - Asks for a final confirmation listing the selected group names.
    - On confirmation, calls `handleDeleteGroups_1_1(selectedGroupNames)` and summarizes the number of deleted series and single events via `ui.alert` and `Logger`.
  - This kept all deletion semantics identical to `deleteEventsByGroup_1_1` but introduced a minimal interactive filter without reintroducing HtmlService.

- **Rationale**  
  - Using `SpreadsheetApp.getUi()` keeps everything server-side inside the spreadsheet context, avoiding the HtmlService iframe and `google.script.run` layer that might trigger permission anomalies.
  - The command-line-like input model is less user-friendly than checkboxes but very robust.

### 3. Checkbox dialog v2 (Usuwanie_wydarzen_1.2)

- **Goal**: Reintroduce a **checkbox-based selection UI** (like 1.0) while keeping the battle-tested deletion core, and isolating version-specific behavior in a separate file (`Usuwanie_wydarzen_1.2`).

- **Design**  
  - New file: `Usuwanie_Wydarzen_1.2`, with versioned config (`CALENDAR_ID_1_2`, `SCHEDULE_SHEET_NAME_1_2`, etc.) and helpers, but semantically the same as 1.1.
  - Core deletion:
    - `handleDeleteGroups_1_2(selectedGroupNames)`:
      - Mirrors `handleDeleteGroups_1_1`, but uses the 1.2 config/constants.
      - Same logic for windows, deduped recurring series deletion, single-event deletion, and `Delete_Report` writing (now via `writeDeleteReport_1_2`).
  - Checkbox dialog:
    - `deleteEventsByGroup_1_2_dialog()`:
      - Reads `groupInfo_1_2` and builds a sorted list of group names.
      - If no groups, uses `SpreadsheetApp.getUi().alert` and exits.
      - Calls `SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), ...)` where `html` is generated by `buildGroupDeleteDialogHtml_1_2(groupNames)`.
    - `buildGroupDeleteDialogHtml_1_2(groupNames)`:
      - Embeds `groupNames` as JSON.
      - Renders one checkbox row per group with **Select all**, **Clear**, **Delete selected**, **Cancel** controls.
      - On “Delete selected”:
        - Gathers all checked checkbox values.
        - If none selected, shows an inline status message.
        - Otherwise calls `google.script.run.handleDeleteGroups_1_2_fromDialog(selected)` with success/failure handlers.
    - `handleDeleteGroups_1_2_fromDialog(selectedGroupNames)`:
      - Delegates to `handleDeleteGroups_1_2(selectedGroupNames || [])`.
      - Aggregates `totalSeries` and `totalSingle` across all groups.
      - Returns `{ ok: true, totalSeries, totalSingle, deleted: [...] }` to the client.
    - The client success handler displays a final status like:
      - “Done. Deleted X series and Y single events. See Delete_Report sheet for details.”

- **Rationale and differences vs 1.0**  
  - Behavior:
    - Series/single-event deletion and audit report semantics are deliberately identical to 1.0; the only structural change is that the **core deletion logic is now clearly separated** into `handleDeleteGroups_1_2`, making it easier to test and to call from multiple UIs.
  - Stability:
    - Having 1.1 (script-only, no HtmlService) and 1.1_withUi (spreadsheet UI) available provides fallbacks if HtmlService again causes \"permission denied\" in certain environments.
  - UX:
    - 1.2 restores a point-and-click checkbox UI, which is more user-friendly than comma-separated input, while still logging and reporting exactly as before.

### Deletion variants overview

- **Usuwanie_wydarzen_1.0**  
  - HtmlService checkbox dialog + `deleteEventsByGroup()` / `handleDeleteGroups()`.  
  - Original implementation; sometimes hit \"permission denied\" from the dialog context.

- **Usuwanie_wydarzen_1.1**  
  - `deleteEventsByGroup_1_1()` – script-only, deletes **all groups**; usable from the Apps Script editor.  
  - `deleteEventsByGroup_1_1_withUi()` – spreadsheet prompt UI (comma-separated list or `*`), no HtmlService.

- **Usuwanie_wydarzen_1.2**  
  - `deleteEventsByGroup_1_2_dialog()` – HtmlService checkbox dialog v2 built on top of `handleDeleteGroups_1_2`.  
  - `handleDeleteGroups_1_2()` – core deletion; shared between dialog-based and potential future non-UI entrypoints.

Across all versions, the **core invariant** is preserved: for each chosen group name, delete all recurring series and single events whose title matches the group, within that group’s schedule-based date range (with a configurable ± buffer), and write a detailed `Delete_Report` sheet for audit.

---

## Confirmed = False: do not create events when no confirmed participants (BEH-11)

**Date:** 14.03.2026

**Context:** Automatyzacja_1.0 was creating calendar events even when all participants in the Test_Schedule_Month_Script tab had the "Confirmed" column set to False. Users expect no events to be created for such groups.

**Goal:** When every participant in a group has Confirmed = False (and there are no trial participants in range), the script should skip that group and not create or update any calendar events.

**Current behavior analysis:**

- The "Confirmed" column is read in `validateRow` (HEADER_CONFIRMED) and stored as `row.confirmed`.
- `partitionParticipants` correctly puts only rows with `row.confirmed === true` into `seriesParticipants`; rows with a trial date in range go to `trialParticipants`. So who gets added as attendees already respected Confirmed.
- The decision to create a series at all is in `upsertEventsForGroups`:  
  `needSeries = seriesEmails.length > 0 || trialParticipants.length > 0 || CREATE_EVENTS_WITH_NO_CONFIRMED_ATTENDEES`
- When all rows have Confirmed = False and no trial in range: `seriesEmails` and `trialParticipants` are both empty, but `CREATE_EVENTS_WITH_NO_CONFIRMED_ATTENDEES` was set to `true` (BEH-11 default), so `needSeries` remained true and events were still created.

**Root cause:** The config flag `CREATE_EVENTS_WITH_NO_CONFIRMED_ATTENDEES = true` forces the script to create events even when there are no confirmed and no trial participants.

**Fix:** Set `CREATE_EVENTS_WITH_NO_CONFIRMED_ATTENDEES = false` so that when there are no confirmed and no trial participants, the group is skipped (GROUP_SKIPPED_NO_ATTENDEES) and no calendar events are created.

**What was changed in the script (Automatyzacja_1.0):**

- Line 17: `var CREATE_EVENTS_WITH_NO_CONFIRMED_ATTENDEES = true;` → `var CREATE_EVENTS_WITH_NO_CONFIRMED_ATTENDEES = false;` with comment updated to: "BEH-11: skip groups with no confirmed (and no trial) participants".

**Result:** Groups where all participants have Confirmed = False and no trial participants in range no longer get any calendar events; they are reported as GROUP_SKIPPED_NO_ATTENDEES. Groups with at least one confirmed or one trial participant still get events as before.

---

## Confirmed = False: do not add trial participants when Confirmed is false

**Date:** 14.03.2026

**Context:** When a row had "Data zajęć próbnych" (trial date) set but "Confirmed" set to False, that participant was still added to the calendar event as a trial guest. Users want only confirmed participants (including confirmed trial participants) to be added.

**Goal:** Trial participants (rows with Data zajęć próbnych in the schedule range) should be added to events only when Confirmed = True. Rows with a trial date but Confirmed = False must not be added to any event.

**Current behavior analysis:**

- In `partitionParticipants`, any row with a trial date in the schedule range (`hasTrial`) was always pushed to `trialParticipants`, regardless of `row.confirmed`.
- `row.confirmed` was only checked for non-trial rows (for `seriesParticipants`). So trial-eligible rows were all sent to `addTrialGuestToInstance`.

**Root cause:** The condition for adding to `trialParticipants` was only `hasTrial`; it did not require `row.confirmed`.

**Fix:** Require both a trial date in range and Confirmed = true for a row to be added to `trialParticipants`: use `if (hasTrial && row.confirmed)` instead of `if (hasTrial)`.

**What was changed in the script (Automatyzacja_1.0):**

- In `partitionParticipants` (around lines 702–706):  
  `if (hasTrial) { trials.push(row); }` → `if (hasTrial && row.confirmed) { trials.push(row); }`  
  The `else if (row.confirmed) { series.push(row); }` branch is unchanged.

**Result:** Rows with trial date set but Confirmed = false are no longer added to `trialParticipants`, so they are not added to any calendar event. Trial date + Confirmed = true still adds them as trial guests. Non-trial rows are still handled as before (only confirmed go to series).

---

## Post-creation summary prompt (EVENT_CREATED per group)

**Date:** 15.03.2026

**Context:** After running `runSync` in `Automatyzacja_1.0`, users want a quick confirmation of **which groups (`Nazwa Grupy`) actually had events created** and **how many events per group**, without digging into the `Run_Report` sheet.

**Goal:** After calendar operations complete, show a single `SpreadsheetApp.getUi().alert` that summarizes, per group name, how many events were created, or clearly states when no new events were created.

**Decisions:**

- **Source of truth:** Use existing `reportEntries` and only count entries with `action === 'EVENT_CREATED'` and a non-empty `groupName`. Do **not** count `EVENT_UPDATED`, so the summary truly reflects created events.
- **Aggregation:** Introduce `getCreatedEventsByGroup(reportEntries)` that returns a map `groupName -> count` based on those `EVENT_CREATED` entries.
- **UI behavior:** Implement `showCreatedEventsSummary(reportEntries)` that:
  - Calls `getCreatedEventsByGroup`.
  - If the map is empty, shows: `Nie utworzono żadnych nowych wydarzeń w kalendarzu.`.
  - Otherwise shows: `Utworzono wydarzenia w następujących grupach:` followed by one line per group in the form `- {groupName}: {count} wydarzenie/wydarzenia/wydarzeń` with simple Polish plural handling (1 → `wydarzenie`, 2–4 → `wydarzenia`, 5+ → `wydarzeń`).
  - Uses `SpreadsheetApp.getUi().alert('Synchronizacja kalendarza', message, ui.ButtonSet.OK)` inside a `try/catch` so that time-driven or non-UI contexts don’t throw.
- **When to show:** Call `showCreatedEventsSummary(reportEntries)` only on the successful path of `runSync`, after `writeRunReport(ss, reportEntries)`. Do not show the alert when `PHASE1_ONLY` is true or when a fatal error is caught.

**Implementation summary (Automatyzacja_1.0):**

- Added `getCreatedEventsByGroup(reportEntries)` near `makeReportEntry` / `writeRunReport`.
- Added `showCreatedEventsSummary(reportEntries)` that builds and displays the alert.
- Updated `runSync` so that, after writing the run report, it calls `showCreatedEventsSummary(reportEntries)` to surface this information to the user.

**Result:** After a normal sync, users immediately see which groups had events created and how many per group, while detailed diagnostics remain in `Run_Report`. Background/trigger executions remain safe because the UI call is guarded.

---

## Calendar ID taken from sheet instead of hardcoded (creation + deletion tools)

**Date:** 17.03.2026

**Context:** Previously, both the main automation script (`Automatyzacja_1.0`) and the deletion tool (`Usuwanie_Wydarzen/Usuwanie_wydarzen_1.0`) had a hardcoded `CALENDAR_ID` with the full calendar address in the script source. The project now keeps the effective calendar ID in the schedule sheet (`Test_Schedule_Month_Script`) as a column `Calendar ID` (placed after `Nazwa Grupy`).

**Goals:**

- Make the calendar ID configurable from the sheet, without editing script source.
- Avoid storing the real calendar address directly in code.
- Keep behavior safe: fail fast with a clear error if no calendar ID is configured anywhere.

**Decisions (Automatyzacja_1.0 – event creation):**

- Added header constant: `HEADER_CALENDAR_ID = 'Calendar ID';`.
- Added `getCalendarIdFromSheet(sheet)` that:
  - Reads `sheet.getDataRange().getValues()`.
  - Locates the `Calendar ID` column by header text (not by position).
  - Returns the first non-empty cell value in that column (trimmed), or `null` if none.
- In `runSync`:
  - After resolving `scheduleSheet`, we compute:
    - `var sheetCalendarId = getCalendarIdFromSheet(scheduleSheet);`
    - `var calendarId = sheetCalendarId || CALENDAR_ID;`
  - If `calendarId` is falsy, we throw:  
    `Calendar ID is not configured. Please fill the "Calendar ID" column in sheet "Test_Schedule_Month_Script" or set CALENDAR_ID in the script.`
  - We then call `CalendarApp.getCalendarById(calendarId)` and pass `calendarId` down into `upsertEventsForGroups`.
- Updated `upsertEventsForGroups` signature to accept `calendarId`, and replaced usages where the raw string was needed:
  - `getRecurringEventIdForSlotByListing(calendarId, ...)`
  - `addTrialGuestToInstance(calendarId, ...)`
- Adjusted config: `CALENDAR_ID` remains as a default but is now set to an empty string in code; the live ID is expected to come from the sheet.

**Decisions (Usuwanie_wydarzen_1.0 – deletion tool):**

- Config: replaced the hardcoded calendar address with an empty default:  
  `var CALENDAR_ID = '';` with a comment explaining that the main source is the sheet `Calendar ID` column.
- Introduced header + helper mirroring the main script:
  - `HEADER_CALENDAR_ID_1_1 = 'Calendar ID';`
  - `getCalendarIdFromSheet_1_1(sheet)`:
    - Same logic: find `Calendar ID` header, return first non-empty cell in that column or `null`.
- Updated `authorizeDeleteTool_1_1`:
  - Computes `sheetCalendarId` and `calendarId = sheetCalendarId || CALENDAR_ID`.
  - Throws a clear error if `calendarId` is missing, otherwise uses `CalendarApp.getCalendarById(calendarId)`.
- Updated script-only entrypoint `deleteEventsByGroup_1_1`:
  - Resolves `calendarId` from sheet/default in the same way, with the same “not configured” error message.
  - Passes `calendarId` into `handleDeleteGroups_1_1(groupNames, calendarId)`.
- Updated UI entrypoint `deleteEventsByGroup_1_1_withUi`:
  - After validating user input, resolves `calendarId` via `getCalendarIdFromSheet_1_1`.
  - If no `calendarId`, shows an alert to the user and aborts.
  - Calls `handleDeleteGroups_1_1(selectedGroupNames, calendarId)`.
- Updated `handleDeleteGroups_1_1`:
  - Signature: `function handleDeleteGroups_1_1(selectedGroupNames, calendarId)`.
  - Uses `calendarId` when calling `CalendarApp.getCalendarById(calendarId)` and in error messages, instead of the global `CALENDAR_ID`.

**Result:**

- Both the creation and deletion flows get their effective calendar ID from the sheet’s `Calendar ID` column, falling back to `CALENDAR_ID` only if explicitly configured in code (now empty by default).
- No real calendar address is stored in the repository; the sheet is the primary configuration surface.
- Errors are explicit when no calendar ID is configured, which should make misconfiguration easy to diagnose.

---

## Entry – 18.03.2026

**Context**: `Automatyzacja_1.0` calendar `description` was rendering “Classroom board” twice:
- once as plain text in `RAW_TEMPLATE` (`You will work with the miro board Classroom board`)
- and again as the hyperlink produced by `{link do Miro}` (with link text “Classroom board”).

**Fix / decision**: Update `RAW_TEMPLATE` to remove the plain-text “Classroom board” wording so `{link do Miro}` remains the single source of that label and hyperlink in the final description.

**Files affected**:
- `Automatyzacja_1.0` – `RAW_TEMPLATE` / `renderInvitationBody()` for Miro link rendering.

---

## Rerun-safe sync: managed series identity, future-only guests, obsolete future cleanup (memory)

**Date:** 01.04.2026

**Context (problem / user intent):**

- Users rerun `runSync()` after changing the sheet; they expect **existing** recurring series to be **updated** (time, recurrence, description, attendees) rather than duplicated.
- **Deletion / cleanup** must **not remove past lessons**: only cancel or remove what **would still occur in the future** when a group/slot disappears from the sheet.
- **Guest list changes** must **not** be applied in a way that strips invitees from **past** occurrences. Google Calendar often applies `addGuest` / `removeGuest` on the **whole** recurring series, which can revoke access to historical instances (including links in the event body). Confirmed attendees should be reconciled **only on future instances**, and if there are **no future instances**, guest sync should be skipped.

**Goals:**

1. Reliable matching of sheet rows to calendar series (avoid ambiguous title-only matches where possible).
2. Update series via `CalendarApp` for time/recurrence/description; sync attendees via **Calendar API** on **future instances only**.
3. After sync, **orphan** managed series (no longer present in the sheet) get **future-only** cleanup: trim `RRULE` `UNTIL` after the last past occurrence, or delete the master only when **all** instances are still in the future.
4. Keep trial flow **add-only** on single occurrences; no trial guest removals.

**Design decisions:**

- **Stable identity:** Store a deterministic private extended property on the **recurring master** (Calendar API v3), e.g. key `atwkManagedSeries_v1`, value `v1|group|weekday|hour|minute|slotIndex` (pipe in group name escaped). Lookup via `Calendar.Events.list` with `privateExtendedProperty=key=value`.
- **Legacy series:** If no property yet, fall back to existing `findExistingGroupSeriesForSlot` / `findExistingGroupSeries`, log `SERIES_MATCHED_LEGACY_FALLBACK`, then **patch** the master to set the managed property so the next run is stable. If `CalendarApp.getEventById(iCalUID)` fails, fall back to scanning `calendar.getEvents` for matching `getId() === iCalUID`.
- **Upsert flow:** `upsertSlotSeries` registers each active slot’s managed value in `desiredManagedKeys`; `runSync` calls `cleanupObsoleteManagedSeries` after upserts.
- **Guests:** `updateSeries` **does not** call `series.addGuest` / `removeGuest`. Use `syncFutureInstanceAttendees`: paginate `Calendar.Events.instances` from **now**, and for each future instance `patch` `attendees` to the desired confirmed set while preserving the organizer. Report `GUESTS_SYNC_FUTURE_ONLY_APPLIED`, `GUESTS_SYNC_SKIPPED_NO_FUTURE`, or `GUESTS_SYNC_FUTURE_ONLY_FAILED`.
- **Obsolete cleanup:** List masters with the private key (property name only, any value). For values **not** in `desiredManagedKeys`: if there is **no** future instance, do nothing; if **all** instances are future-only, `Calendar.Events.remove` the master; if there are **past** instances, `patch` recurrence with a new `RRULE` `UNTIL` derived from the **end** of the last past instance (strip prior `UNTIL`/`COUNT`, append new `UNTIL` in UTC `yyyyMMdd'T'HHmmss'Z'`). Optional `OBSOLETE_SERIES_CLEANUP_DRY_RUN` logs intent without mutating.
- **Trials:** Resolve recurring id with `getRecurringApiIdManagedOrSlot` (managed key first, then title/slot listing).

**Implementation summary (`Automatyzacja_1.0`):**

- Config: `MANAGED_SERIES_PRIVATE_KEY`, `OBSOLETE_SERIES_CLEANUP_DRY_RUN`.
- Helpers: `buildManagedSeriesPrivateValue`, `findRecurringMasterByManagedPrivate`, `ensureManagedSeriesExtendedProps`, `apiRecurringIdToEventSeries`, `findCalendarAppSeriesByIcalUid`, `upsertSlotSeries`, `syncFutureInstanceAttendees`, `buildRecurrenceTrimmedToUntil`, `trimOrDeleteObsoleteManagedSeries`, `cleanupObsoleteManagedSeries`, `getRecurringApiIdManagedOrSlot`, `slotIndexInSchedule`.
- `createSeries` takes `calendarId` + `managedValue`, tags master after create.
- `runSync`: `desiredManagedKeys` → `upsertEventsForGroups` → `cleanupObsoleteManagedSeries`.

**Report actions to look for:** `SERIES_MATCHED_BY_MANAGED_KEY`, `SERIES_MATCHED_LEGACY_FALLBACK`, `GUESTS_SYNC_*`, `OBSOLETE_SERIES_FUTURE_TRIMMED`, `OBSOLETE_SERIES_REMOVED_FUTURE_ONLY`, `OBSOLETE_SERIES_CLEANUP_FAILED`.

**Caveat (for future readers):** `series.setDescription` still updates the **series** description in Calendar; Google may propagate description changes across instances. This memory’s hard requirement was **guests on past instances**, not freezing past description text.

**Cursor / AI:** Read this entry before changing rerun behavior, guest sync, or obsolete-series cleanup in `Automatyzacja_1.0`.