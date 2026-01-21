# Availability Calculator

This Google Apps Script project calculates availability from your primary Google Calendar and inserts results into the active Google Doc.

## How to use

1. Open the Google Doc bound to this script.
2. Use **Availability → Calculate availability** to open the dialog.
3. Select one or more availability outputs (this week and/or next week, 2-hour blocks and/or full availability).
4. (Optional) Check recurring events you want to ignore for the next two weeks.
5. Click **Insert** to overwrite the document with the requested availability blocks.

## Notes

- Availability is calculated between **10:00 and 18:30** local time, Monday through Friday.
- Declined events are treated as free time. Accepted, maybe, invited, and owner events are treated as busy.
- Recurring events are shown once per series (daily events appear as a single checkbox).
