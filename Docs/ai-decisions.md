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