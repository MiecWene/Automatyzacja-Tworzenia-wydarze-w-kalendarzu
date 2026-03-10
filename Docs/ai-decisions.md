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


Files affected: Automatyzacja_1.0

Next steps: Builded the suggested config