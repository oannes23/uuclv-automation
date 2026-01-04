/**
 * UUCLV Events Automation (Combined Building Reservations + Event Feed)
 *
 * SHEET: "Form Responses 1" – canonical event data
 *
 * Column layout (1-based columns / 0-based rowValues[] index):
 *   1 /  0 : A - Approval
 *   2 /  1 : B - Approver
 *   3 /  2 : C - Timestamp
 *   4 /  3 : D - Email Address
 *   5 /  4 : E - Event Name
 *   6 /  5 : F - Event Description (please include contact info...)
 *   7 /  6 : G - Event Date
 *   8 /  7 : H - Event Start Time
 *   9 /  8 : I - Event End Time
 *  10 /  9 : J - What part(s) of the building do you want to use?
 *  11 / 10 : K - Who is the target audience for this event?
 *  12 / 11 : L - If this event is not Private, where should we advertise it?
 *  13 / 12 : M - Setup/teardown time before/after event
 *  14 / 13 : N - Key holder / AV support needed?
 *  15 / 14 : O - Do you need someone to create a graphic for this event?
 *  16 / 15 : P - If you have your own graphic already, please upload it here
 *  17 / 16 : Q - Building Calendar Event ID
 *  18 / 17 : R - Website Calendar Event ID
 *
 * SHEET: "Form Responses 2" – canonical repeating (monthly recurring) event data
 *
 * Column layout (1-based columns / 0-based rowValues[] index):
 *   1 /  0 : A - Approval
 *   2 /  1 : B - Approver
 *   3 /  2 : C - Timestamp
 *   4 /  3 : D - Email Address
 *   5 /  4 : E - Event Name
 *   6 /  5 : F - Event Description (please include contact info...)
 *   7 /  6 : G - What week of the month does your event occur on? (Every|First|Second|Third|Fourth)
 *   8 /  7 : H - What day of the week does your event occur on? (Sunday..Saturday)
 *   9 /  8 : I - Event Start Time
 *  10 /  9 : J - Event End Time
 *  11 / 10 : K - Who is the target audience for this event?
 *  12 / 11 : L - If this event is not Private, where should we advertise it?
 *  13 / 12 : M - What part(s) of the building do you want to use?
 *  14 / 13 : N - Setup/teardown time before/after event
 *  15 / 14 : O - Key holder / AV support needed?
 *  16 / 15 : P - Do you need someone to create a graphic for this event?
 *  17 / 16 : Q - If you have your own graphic already, please upload it here
 *  18 / 17 : R - Building Calendar Recurring Event ID (series ID)
 *  19 / 18 : S - Website Calendar Recurring Event ID (series ID)
 *  20 / 19 : T - Skip Months (comma-separated month numbers, e.g. "3, 7" to skip March and July)
 *
 * SHEET: "Config" – approval statuses + calendar IDs
 *
 * Expected structure:
 *   A1: "Approval Statuses"
 *   A2:A4: Pending, Approved, Rejected
 *   B1: "Member Calendar ID", B2: <member calendar ID>
 *   C1: "Public Calendar ID", C2: <public calendar ID>
 *   D1: "Building Calendar ID", D2: <building calendar ID>
 *   E1: "Recurring Year", E2: <e.g. 2026>
 */

/**
 * Helper: combine a date cell and a time cell into a single Date object.
 */
function combineDateAndTime(date, time) {
  if (!(date instanceof Date) || !(time instanceof Date)) {
    return null;
  }
  const combined = new Date(date);
  combined.setHours(time.getHours(), time.getMinutes(), time.getSeconds(), 0);
  return combined;
}

/**
 * Helper: create a new Date with the same Y/M/D as date and same HMS as time.
 */
function combineDateWithTime_(date, time) {
  if (!(date instanceof Date) || !(time instanceof Date)) {
    return null;
  }
  const d = new Date(date);
  d.setHours(time.getHours(), time.getMinutes(), time.getSeconds(), 0);
  return d;
}

/**
 * Helper: normalize strings for case-insensitive comparisons.
 */
function normalize_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).toLowerCase().trim();
}

/**
 * Helper: read the configured recurring year from Config!E2. Defaults to current year.
 */
function getRecurringYear_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    return new Date().getFullYear();
  }
  const raw = configSheet.getRange('E2').getValue();
  const y = parseInt(raw, 10);
  if (!y || isNaN(y)) {
    return new Date().getFullYear();
  }
  return y;
}

/**
 * Helper: read calendar IDs (member, public, building) from Config!B2:D2.
 */
function getCalendarIds_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    return {
      member: null,
      public: null,
      building: null,
    };
  }

  const values = configSheet.getRange('B2:D2').getValues();
  const row = values && values[0] ? values[0] : [];

  return {
    member: row[0] || null,   // B2
    public: row[1] || null,   // C2
    building: row[2] || null, // D2
  };
}

/**
 * Helper: map setup/teardown choice to minutes of padding
 * before and after the advertised event time.
 */
function getPaddingMinutes_(setupTeardownRaw) {
  const text = normalize_(setupTeardownRaw);
  if (!text || text === 'none') {
    return 0;
  }
  if (text.indexOf('30') !== -1) {
    return 30;
  }
  if (text.indexOf('1 hour') !== -1 || text === '1 hr' || text === '1hr') {
    return 60;
  }
  if (text.indexOf('2 hour') !== -1 || text === '2 hr' || text === '2hr') {
    return 120;
  }
  return 0;
}

/**
 * Helper: choose which "website" calendar to use based on target audience text.
 * Returns calendarId string or null.
 *
 * This is intentionally tolerant of small wording changes like:
 * - "Members & Friends"
 * - "Members and Friends (UUCLV)"
 * - "General public"
 */
function getWebsiteCalendarIdForTargetAudience_(targetAudienceRaw, calendarIds) {
  const t = normalize_(targetAudienceRaw);
  if (!t) return null;

  // Private events never get a website calendar event.
  if (t.indexOf('private') !== -1) return null;

  // Member calendar
  if (t.indexOf('members') !== -1) return (calendarIds && calendarIds.member) || null;

  // Public calendar
  if (t.indexOf('general public') !== -1) return (calendarIds && calendarIds.public) || null;
  if (t.indexOf('public') !== -1) return (calendarIds && calendarIds.public) || null;

  return null;
}

/**
 * Helper: append a note to a cell (doesn't overwrite existing note content).
 */
function appendNote_(range, message) {
  if (!range || typeof range.getNote !== 'function' || typeof range.setNote !== 'function') {
    return;
  }
  const existing = range.getNote() || '';
  const next = existing ? existing + '\n\n' + message : message;
  range.setNote(next);
}

/**
 * Helper: map day-of-week string (Sunday..Saturday) to RRULE BYDAY token (SU..SA).
 */
function dayOfWeekToRruleByday_(dayOfWeekRaw) {
  const text = String(dayOfWeekRaw || '').trim();
  const map = {
    Sunday: 'SU',
    Monday: 'MO',
    Tuesday: 'TU',
    Wednesday: 'WE',
    Thursday: 'TH',
    Friday: 'FR',
    Saturday: 'SA',
  };
  return map[text] || null;
}

/**
 * Helper: map day-of-week string (Sunday..Saturday) to JS weekday index (0=Sun..6=Sat).
 */
function dayOfWeekToJsIndex_(dayOfWeekRaw) {
  const text = String(dayOfWeekRaw || '').trim();
  const map = {
    Sunday: 0,
    Monday: 1,
    Tuesday: 2,
    Wednesday: 3,
    Thursday: 4,
    Friday: 5,
    Saturday: 6,
  };
  return Object.prototype.hasOwnProperty.call(map, text) ? map[text] : null;
}

/**
 * Helper: map repeat pattern string to BYSETPOS value (1-4) or null.
 */
function repeatPatternToBysetpos_(repeatPatternRaw) {
  const text = String(repeatPatternRaw || '').trim();
  const map = {
    Every: null,
    First: 1,
    Second: 2,
    Third: 3,
    Fourth: 4,
  };
  return Object.prototype.hasOwnProperty.call(map, text) ? map[text] : null;
}

/**
 * Helper: parse comma-separated month numbers into a Set of 0-indexed month values.
 * Input: "3, 7, 12" (1-indexed month numbers)
 * Output: Set {2, 6, 11} (0-indexed for JS Date months)
 */
function parseSkipMonths_(skipMonthsRaw) {
  if (!skipMonthsRaw) return new Set();
  const parts = String(skipMonthsRaw).split(',');
  const result = new Set();
  for (let i = 0; i < parts.length; i++) {
    const num = parseInt(parts[i].trim(), 10);
    if (!isNaN(num) && num >= 1 && num <= 12) {
      result.add(num - 1); // Convert to 0-indexed
    }
  }
  return result;
}

/**
 * Helper: compute all occurrence dates in a month for either:
 * - Every: all matching weekdays in that month
 * - Nth (1-4): the Nth weekday in that month (always exists for 1-4)
 *
 * Returns array of Date objects (date-only, local timezone) at midnight.
 */
function computeMonthlyOccurrenceDates_(year, month0, jsWeekdayIndex, bysetposOrNull) {
  const dates = [];
  const firstOfMonth = new Date(year, month0, 1);

  if (bysetposOrNull === null) {
    // "Every": all matching weekdays in this month.
    const d = new Date(firstOfMonth);
    const delta = (jsWeekdayIndex - d.getDay() + 7) % 7;
    d.setDate(d.getDate() + delta);
    while (d.getMonth() === month0) {
      dates.push(new Date(d.getFullYear(), d.getMonth(), d.getDate()));
      d.setDate(d.getDate() + 7);
    }
    return dates;
  }

  // Nth weekday-in-month (1-4)
  const nth = bysetposOrNull;
  const d = new Date(firstOfMonth);
  const delta = (jsWeekdayIndex - d.getDay() + 7) % 7;
  d.setDate(d.getDate() + delta + (nth - 1) * 7);
  if (d.getMonth() === month0) {
    dates.push(new Date(d.getFullYear(), d.getMonth(), d.getDate()));
  }
  return dates;
}

/**
 * Helper: compute DTSTART date (first occurrence in the configured year).
 */
function computeFirstOccurrenceDateForYear_(year, jsWeekdayIndex, bysetposOrNull) {
  for (let m = 0; m < 12; m++) {
    const dates = computeMonthlyOccurrenceDates_(year, m, jsWeekdayIndex, bysetposOrNull);
    if (dates.length) {
      return dates[0];
    }
  }
  return null;
}

/**
 * Create a recurring event series using the Advanced Calendar service (Google Calendar API).
 *
 * IMPORTANT: This requires enabling:
 * - Apps Script: Services → Advanced Google services → Calendar API
 * - GCP Project: Google Calendar API enabled
 */
function createRecurringEventSeriesAdvanced_(calendarId, title, startDateTime, endDateTime, rrule, options) {
  if (
    typeof Calendar === 'undefined' ||
    !Calendar.Events ||
    typeof Calendar.Events.insert !== 'function'
  ) {
    throw new Error(
      'Advanced Calendar service is not enabled. Enable Advanced Google Service "Calendar API" to create recurring events.'
    );
  }

  const tz = Session.getScriptTimeZone();
  const resource = {
    summary: title,
    description: (options && options.description) || '',
    location: (options && options.location) || '',
    start: {
      dateTime: startDateTime.toISOString(),
      timeZone: tz,
    },
    end: {
      dateTime: endDateTime.toISOString(),
      timeZone: tz,
    },
    recurrence: ['RRULE:' + rrule],
  };

  const created = Calendar.Events.insert(resource, calendarId);
  return created && created.id ? created.id : null;
}

/**
 * Trigger: On form submit
 *
 * Ensures the metadata columns are initialized for new submissions in:
 *  - "Form Responses 1" (one-shot events)
 *  - "Form Responses 2" (repeating events)
 *
 * For Form Responses 1:
 *  - Column A: Approval (defaults to "Pending" if blank)
 *  - Column B: Approver (cleared)
 *  - Column Q: Building Calendar Event ID (cleared)
 *  - Column R: Website Calendar Event ID (cleared)
 *
 * For Form Responses 2:
 *  - Column A: Approval (defaults to "Pending" if blank)
 *  - Column B: Approver (cleared)
 *  - Column R: Building Calendar Recurring Event ID (cleared)
 *  - Column S: Website Calendar Recurring Event ID (cleared)
 *
 * Install this as an "On form submit" trigger.
 */
function onFormSubmit(e) {
  const APPROVAL_COL = 1;   // A
  const APPROVER_COL = 2;   // B

  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  const row = e.range.getRow();
  if (row <= 1) {
    // skip header
    return;
  }

  if (sheetName === 'Form Responses 1') {
    const BUILDING_EVENT_ID_COL = 17; // Q
    const WEBSITE_EVENT_ID_COL = 18;  // R

    // Initialize Approval column to "Pending" if blank
    const approvalCell = sheet.getRange(row, APPROVAL_COL);
    if (!approvalCell.getValue()) {
      approvalCell.setValue('Pending');
    }

    // Clear Approver for a brand-new submission
    const approverCell = sheet.getRange(row, APPROVER_COL);
    if (approverCell.getValue()) {
      approverCell.setValue('');
    }

    // Clear Building and Website Event IDs so re-submissions never reuse stale IDs
    const buildingEventIdCell = sheet.getRange(row, BUILDING_EVENT_ID_COL);
    if (buildingEventIdCell.getValue()) {
      buildingEventIdCell.setValue('');
    }

    const websiteEventIdCell = sheet.getRange(row, WEBSITE_EVENT_ID_COL);
    if (websiteEventIdCell.getValue()) {
      websiteEventIdCell.setValue('');
    }
  } else if (sheetName === 'Form Responses 2') {
    const BUILDING_RECURRING_ID_COL = 18; // R
    const WEBSITE_RECURRING_ID_COL = 19;  // S

    const approvalCell = sheet.getRange(row, APPROVAL_COL);
    if (!approvalCell.getValue()) {
      approvalCell.setValue('Pending');
    }

    const approverCell = sheet.getRange(row, APPROVER_COL);
    if (approverCell.getValue()) {
      approverCell.setValue('');
    }

    const buildingSeriesIdCell = sheet.getRange(row, BUILDING_RECURRING_ID_COL);
    if (buildingSeriesIdCell.getValue()) {
      buildingSeriesIdCell.setValue('');
    }

    const websiteSeriesIdCell = sheet.getRange(row, WEBSITE_RECURRING_ID_COL);
    if (websiteSeriesIdCell.getValue()) {
      websiteSeriesIdCell.setValue('');
    }
  }
}

/**
 * Trigger: On edit
 *
 * When Approval is changed on:
 *  - 'Form Responses 1' (one-shot events): creates single calendar events and stores event IDs (Q/R)
 *  - 'Form Responses 2' (repeating events): creates RRULE-based recurring series and stores series IDs (R/S),
 *    then rebuilds the 'Recurring Instances' expansion sheet.
 *
 * When Skip Months (column T) is changed on 'Form Responses 2':
 *  - Rebuilds 'Recurring Instances' without touching calendars.
 *
 * Install this as an "On edit" trigger on the spreadsheet.
 */
function onApprovalEdit(e) {
  const FIRST_DATA_ROW = 2;
  const APPROVAL_COL = 1;   // Column A: Approval
  const APPROVER_COL = 2;   // Column B: Approver
  const SKIP_MONTHS_COL = 20; // Column T: Skip Months

  if (!e || !e.range) {
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  const row = range.getRow();
  if (row < FIRST_DATA_ROW) {
    return;
  }

  const col = range.getColumn();

  // Handle Skip Months (column T) edits on Form Responses 2:
  // Just rebuild instances, no calendar changes.
  if (sheetName === 'Form Responses 2' && col === SKIP_MONTHS_COL) {
    rebuildRecurringInstances_();
    return;
  }

  if (col !== APPROVAL_COL) {
    return;
  }

  const newValue = e.value || '';
  const oldValue = e.oldValue || '';

  const lastCol = sheet.getLastColumn();
  const rowValues = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  // Get the email of the person who edited the Approval cell
  let approverEmail = '';
  if (e.user && typeof e.user.getEmail === 'function') {
    approverEmail = e.user.getEmail();
  } else {
    approverEmail = Session.getActiveUser().getEmail();
  }

  if (approverEmail) {
    sheet.getRange(row, APPROVER_COL).setValue(approverEmail);
    rowValues[1] = approverEmail;
  }

  const calendarIds = getCalendarIds_();

  // ===========================
  // One-shot events (Form Responses 1)
  // ===========================
  if (sheetName === 'Form Responses 1') {
    if (newValue !== 'Approved' || oldValue === 'Approved') {
      return;
    }

    // Column indexes (0-based for rowValues[])
    const EMAIL_COL_INDEX = 3;             // D
    const EVENT_NAME_COL_INDEX = 4;        // E
    const DESCRIPTION_COL_INDEX = 5;       // F
    const DATE_COL_INDEX = 6;              // G
    const START_TIME_COL_INDEX = 7;        // H
    const END_TIME_COL_INDEX = 8;          // I
    const BUILDING_PARTS_COL_INDEX = 9;    // J
    const TARGET_COL_INDEX = 10;           // K
    const ADVERTISE_COL_INDEX = 11;        // L
    const SETUP_TEARDOWN_COL_INDEX = 12;   // M
    const KEYHOLDER_AV_COL_INDEX = 13;     // N
    const NEEDS_GRAPHIC_COL_INDEX = 14;    // O
    const GRAPHIC_UPLOAD_COL_INDEX = 15;   // P
    const BUILDING_EVENT_ID_COL_INDEX = 16; // Q
    const WEBSITE_EVENT_ID_COL_INDEX = 17;  // R

    const emailAddress = rowValues[EMAIL_COL_INDEX];
    const eventName = rowValues[EVENT_NAME_COL_INDEX];
    const eventDescription = rowValues[DESCRIPTION_COL_INDEX];
    const dateVal = rowValues[DATE_COL_INDEX];
    const startTimeVal = rowValues[START_TIME_COL_INDEX];
    const endTimeVal = rowValues[END_TIME_COL_INDEX];
    const buildingParts = rowValues[BUILDING_PARTS_COL_INDEX];
    const targetAudience = rowValues[TARGET_COL_INDEX];
    const advertiseWhere = rowValues[ADVERTISE_COL_INDEX];
    const setupTeardown = rowValues[SETUP_TEARDOWN_COL_INDEX];
    const keyholderAv = rowValues[KEYHOLDER_AV_COL_INDEX];
    const needsGraphic = rowValues[NEEDS_GRAPHIC_COL_INDEX];
    const graphicUpload = rowValues[GRAPHIC_UPLOAD_COL_INDEX];
    const existingBuildingEventId = rowValues[BUILDING_EVENT_ID_COL_INDEX];
    const existingWebsiteEventId = rowValues[WEBSITE_EVENT_ID_COL_INDEX];

    const baseStartDateTime = combineDateAndTime(dateVal, startTimeVal);
    const baseEndDateTime = combineDateAndTime(dateVal, endTimeVal);
    if (!eventName || !baseStartDateTime || !baseEndDateTime) {
      return;
    }

    // --- Building reservation event (if building space requested) ---
    const buildingPartsText = normalize_(buildingParts);
    if (buildingPartsText && !existingBuildingEventId && calendarIds.building) {
      const paddingMinutes = getPaddingMinutes_(setupTeardown);

      const buildingStart = new Date(baseStartDateTime);
      const buildingEnd = new Date(baseEndDateTime);
      if (paddingMinutes > 0) {
        buildingStart.setMinutes(buildingStart.getMinutes() - paddingMinutes);
        buildingEnd.setMinutes(buildingEnd.getMinutes() + paddingMinutes);
      }

      const buildingCalendar = CalendarApp.getCalendarById(calendarIds.building);
      if (buildingCalendar) {
        const buildingDescriptionParts = [];
        if (eventDescription) {
          buildingDescriptionParts.push('Event Description:\n' + eventDescription);
        }
        if (emailAddress) {
          buildingDescriptionParts.push('Contact email: ' + emailAddress);
        }
        if (targetAudience) {
          buildingDescriptionParts.push('Target audience: ' + targetAudience);
        }
        if (setupTeardown) {
          buildingDescriptionParts.push('Setup/teardown padding: ' + setupTeardown);
        }
        if (keyholderAv) {
          buildingDescriptionParts.push('Key holder / AV support: ' + keyholderAv);
        }
        if (needsGraphic) {
          buildingDescriptionParts.push('Needs graphic: ' + needsGraphic);
        }
        if (graphicUpload) {
          buildingDescriptionParts.push('Graphic upload: ' + graphicUpload);
        }
        if (advertiseWhere) {
          buildingDescriptionParts.push('Advertise where: ' + advertiseWhere);
        }
        if (approverEmail) {
          buildingDescriptionParts.push('Approved by: ' + approverEmail);
        }
        const buildingDescription = buildingDescriptionParts.join('\n\n');

        const buildingEvent = buildingCalendar.createEvent(
          eventName,
          buildingStart,
          buildingEnd,
          {
            description: buildingDescription,
            location: buildingParts || '',
          }
        );

        sheet.getRange(row, BUILDING_EVENT_ID_COL_INDEX + 1).setValue(buildingEvent.getId());
      }
    }

    // --- Website calendar event (Member or Public), based on target audience ---
    const websiteCalendarId = getWebsiteCalendarIdForTargetAudience_(targetAudience, calendarIds);

    if (websiteCalendarId && !existingWebsiteEventId) {
      const websiteCalendar = CalendarApp.getCalendarById(websiteCalendarId);
      if (websiteCalendar) {
        const websiteDescriptionParts = [];
        if (eventDescription) {
          websiteDescriptionParts.push('Description:\n' + eventDescription);
        }
        if (emailAddress) {
          websiteDescriptionParts.push('Contact email: ' + emailAddress);
        }
        if (targetAudience) {
          websiteDescriptionParts.push('Target audience: ' + targetAudience);
        }
        if (approverEmail) {
          websiteDescriptionParts.push('Approved by: ' + approverEmail);
        }
        const websiteDescription = websiteDescriptionParts.join('\n\n');

        const websiteEvent = websiteCalendar.createEvent(
          eventName,
          baseStartDateTime,
          baseEndDateTime,
          {
            description: websiteDescription,
            location: buildingParts || '',
          }
        );

        sheet.getRange(row, WEBSITE_EVENT_ID_COL_INDEX + 1).setValue(websiteEvent.getId());
      }
    }

    return;
  }

  // ===========================
  // Repeating events (Form Responses 2)
  // ===========================
  if (sheetName === 'Form Responses 2') {
    // If an approved recurring row becomes non-approved, rebuild the instance sheet.
    if (oldValue === 'Approved' && newValue !== 'Approved') {
      rebuildRecurringInstances_();
      return;
    }

    if (newValue !== 'Approved' || oldValue === 'Approved') {
      return;
    }

    // Column indexes (0-based for rowValues[])
    // NOTE: This mapping matches the current Form Responses 2 layout:
    // ... I: Start, J: End, K: Target Audience, L: Advertise Where, M: Building Spaces, ...
    const EMAIL_COL_INDEX = 3;               // D
    const EVENT_NAME_COL_INDEX = 4;          // E
    const DESCRIPTION_COL_INDEX = 5;         // F
    const REPEAT_PATTERN_COL_INDEX = 6;      // G
    const DAY_OF_WEEK_COL_INDEX = 7;         // H
    const START_TIME_COL_INDEX = 8;          // I
    const END_TIME_COL_INDEX = 9;            // J
    const TARGET_COL_INDEX = 10;             // K
    const ADVERTISE_COL_INDEX = 11;          // L
    const BUILDING_PARTS_COL_INDEX = 12;     // M
    const SETUP_TEARDOWN_COL_INDEX = 13;     // N
    const KEYHOLDER_AV_COL_INDEX = 14;       // O
    const NEEDS_GRAPHIC_COL_INDEX = 15;      // P
    const GRAPHIC_UPLOAD_COL_INDEX = 16;     // Q
    const BUILDING_SERIES_ID_COL_INDEX = 17; // R
    const WEBSITE_SERIES_ID_COL_INDEX = 18;  // S

    const emailAddress = rowValues[EMAIL_COL_INDEX];
    const eventName = rowValues[EVENT_NAME_COL_INDEX];
    const eventDescription = rowValues[DESCRIPTION_COL_INDEX];
    const repeatPatternRaw = rowValues[REPEAT_PATTERN_COL_INDEX];
    const dayOfWeekRaw = rowValues[DAY_OF_WEEK_COL_INDEX];
    const startTimeVal = rowValues[START_TIME_COL_INDEX];
    const endTimeVal = rowValues[END_TIME_COL_INDEX];
    const buildingParts = rowValues[BUILDING_PARTS_COL_INDEX];
    const targetAudience = rowValues[TARGET_COL_INDEX];
    const advertiseWhere = rowValues[ADVERTISE_COL_INDEX];
    const setupTeardown = rowValues[SETUP_TEARDOWN_COL_INDEX];
    const keyholderAv = rowValues[KEYHOLDER_AV_COL_INDEX];
    const needsGraphic = rowValues[NEEDS_GRAPHIC_COL_INDEX];
    const graphicUpload = rowValues[GRAPHIC_UPLOAD_COL_INDEX];
    const existingBuildingSeriesId = rowValues[BUILDING_SERIES_ID_COL_INDEX];
    const existingWebsiteSeriesId = rowValues[WEBSITE_SERIES_ID_COL_INDEX];

    if (!eventName || !(startTimeVal instanceof Date) || !(endTimeVal instanceof Date)) {
      return;
    }

    const year = getRecurringYear_();
    const jsWeekdayIndex = dayOfWeekToJsIndex_(dayOfWeekRaw);
    const byday = dayOfWeekToRruleByday_(dayOfWeekRaw);
    const bysetpos = repeatPatternToBysetpos_(repeatPatternRaw);
    if (jsWeekdayIndex === null || !byday) {
      return;
    }

    const firstDate = computeFirstOccurrenceDateForYear_(year, jsWeekdayIndex, bysetpos);
    if (!firstDate) {
      return;
    }

    const baseStartDateTime = combineDateWithTime_(firstDate, startTimeVal);
    const baseEndDateTime = combineDateWithTime_(firstDate, endTimeVal);
    if (!baseStartDateTime || !baseEndDateTime) {
      return;
    }

    const untilUtc = Utilities.formatDate(
      new Date(Date.UTC(year, 11, 31, 23, 59, 59)),
      'UTC',
      "yyyyMMdd'T'HHmmss'Z'"
    );
    const rruleParts = ['FREQ=MONTHLY', 'BYDAY=' + byday, 'WKST=SU', 'UNTIL=' + untilUtc];
    if (bysetpos !== null) {
      rruleParts.splice(2, 0, 'BYSETPOS=' + bysetpos);
    }
    const rrule = rruleParts.join(';');

    const buildingPartsText = normalize_(buildingParts);

    // --- Building recurring reservation series ---
    if (buildingPartsText && !existingBuildingSeriesId && calendarIds.building) {
      const paddingMinutes = getPaddingMinutes_(setupTeardown);

      const buildingStart = new Date(baseStartDateTime);
      const buildingEnd = new Date(baseEndDateTime);
      if (paddingMinutes > 0) {
        buildingStart.setMinutes(buildingStart.getMinutes() - paddingMinutes);
        buildingEnd.setMinutes(buildingEnd.getMinutes() + paddingMinutes);
      }

      const buildingDescriptionParts = [];
      if (eventDescription) {
        buildingDescriptionParts.push('Event Description:\n' + eventDescription);
      }
      if (emailAddress) {
        buildingDescriptionParts.push('Contact email: ' + emailAddress);
      }
      if (targetAudience) {
        buildingDescriptionParts.push('Target audience: ' + targetAudience);
      }
      if (repeatPatternRaw || dayOfWeekRaw) {
        buildingDescriptionParts.push('Repeats: ' + (repeatPatternRaw || '') + ' ' + (dayOfWeekRaw || ''));
      }
      if (setupTeardown) {
        buildingDescriptionParts.push('Setup/teardown padding: ' + setupTeardown);
      }
      if (keyholderAv) {
        buildingDescriptionParts.push('Key holder / AV support: ' + keyholderAv);
      }
      if (needsGraphic) {
        buildingDescriptionParts.push('Needs graphic: ' + needsGraphic);
      }
      if (graphicUpload) {
        buildingDescriptionParts.push('Graphic upload: ' + graphicUpload);
      }
      if (advertiseWhere) {
        buildingDescriptionParts.push('Advertise where: ' + advertiseWhere);
      }
      if (approverEmail) {
        buildingDescriptionParts.push('Approved by: ' + approverEmail);
      }
      const buildingDescription = buildingDescriptionParts.join('\n\n');

      try {
        const seriesId = createRecurringEventSeriesAdvanced_(
          calendarIds.building,
          eventName,
          buildingStart,
          buildingEnd,
          rrule,
          { description: buildingDescription, location: buildingParts || '' }
        );
        if (seriesId) {
          sheet.getRange(row, BUILDING_SERIES_ID_COL_INDEX + 1).setValue(seriesId);
        }
      } catch (err) {
        Logger.log(err);
        appendNote_(
          sheet.getRange(row, APPROVAL_COL),
          'Could not create Building recurring event series. Ensure Advanced Calendar service is enabled.\n\n' +
            String(err && err.message ? err.message : err)
        );
      }
    }

    // --- Website recurring series ---
    const websiteCalendarId = getWebsiteCalendarIdForTargetAudience_(targetAudience, calendarIds);

    const approvalCell = sheet.getRange(row, APPROVAL_COL);

    if (!websiteCalendarId) {
      // If it wasn't a private event but we couldn't route it, leave a breadcrumb.
      const t = normalize_(targetAudience);
      if (t && t.indexOf('private') === -1) {
        appendNote_(
          approvalCell,
          'Repeating website series not created: target audience "' +
            String(targetAudience || '') +
            '" did not map to Member/Public, or Config Member/Public Calendar ID is missing.'
        );
      }
    } else if (existingWebsiteSeriesId) {
      appendNote_(
        approvalCell,
        'Repeating website series not created because Website Calendar Recurring Event ID (S) is already set.'
      );
    }

    if (websiteCalendarId && !existingWebsiteSeriesId) {
      const websiteDescriptionParts = [];
      if (eventDescription) {
        websiteDescriptionParts.push('Description:\n' + eventDescription);
      }
      if (emailAddress) {
        websiteDescriptionParts.push('Contact email: ' + emailAddress);
      }
      if (targetAudience) {
        websiteDescriptionParts.push('Target audience: ' + targetAudience);
      }
      if (repeatPatternRaw || dayOfWeekRaw) {
        websiteDescriptionParts.push('Repeats: ' + (repeatPatternRaw || '') + ' ' + (dayOfWeekRaw || ''));
      }
      if (approverEmail) {
        websiteDescriptionParts.push('Approved by: ' + approverEmail);
      }
      const websiteDescription = websiteDescriptionParts.join('\n\n');

      try {
        const seriesId = createRecurringEventSeriesAdvanced_(
          websiteCalendarId,
          eventName,
          baseStartDateTime,
          baseEndDateTime,
          rrule,
          { description: websiteDescription, location: buildingParts || '' }
        );
        if (seriesId) {
          sheet.getRange(row, WEBSITE_SERIES_ID_COL_INDEX + 1).setValue(seriesId);
        }
      } catch (err) {
        Logger.log(err);
        appendNote_(
          approvalCell,
          'Could not create Website recurring event series. Ensure Advanced Calendar service is enabled.\n\n' +
            String(err && err.message ? err.message : err)
        );
      }
    }

    rebuildRecurringInstances_();
  }
}

/**
 * Public helper: manually rebuild the "Recurring Instances" sheet from Form Responses 2.
 */
function rebuildRecurringInstances() {
  rebuildRecurringInstances_();
}

/**
 * Script-owned generator: expands each Approved repeating event in Form Responses 2 into
 * per-instance rows for the configured year.
 */
function rebuildRecurringInstances_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Form Responses 2');
  if (!sourceSheet) {
    return;
  }

  const year = getRecurringYear_();

  let outSheet = ss.getSheetByName('Recurring Instances');
  if (!outSheet) {
    outSheet = ss.insertSheet('Recurring Instances');
  }

  const header = [
    'Approver',
    'Event Name',
    'Description',
    'Start',
    'End',
    'Target Audience',
    'Building Spaces',
    'Advertise Where',
    'Setup/Teardown',
    'Needs Graphic?',
    'Graphic',
    'Form Timestamp',
    'Building Event ID',
    'Website Event ID',
  ];

  outSheet.clearContents();
  outSheet.getRange(1, 1, 1, header.length).setValues([header]);

  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  if (lastRow < 2) {
    return;
  }

  const values = sourceSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const out = [];

  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const approval = r[0]; // A
    if (approval !== 'Approved') {
      continue;
    }

    const approver = r[1];             // B
    const timestamp = r[2];            // C
    const eventName = r[4];            // E
    const description = r[5];          // F
    const repeatPatternRaw = r[6];     // G
    const dayOfWeekRaw = r[7];         // H
    const startTimeVal = r[8];         // I
    const endTimeVal = r[9];           // J
    const targetAudience = r[10];      // K
    const advertiseWhere = r[11];      // L
    const buildingSpaces = r[12];      // M
    const setupTeardown = r[13];       // N
    const needsGraphic = r[15];        // P
    const graphic = r[16];             // Q
    const buildingSeriesId = r[17];    // R
    const websiteSeriesId = r[18];     // S
    const skipMonthsRaw = r[19];       // T - Skip Months

    const jsWeekdayIndex = dayOfWeekToJsIndex_(dayOfWeekRaw);
    const skipMonths = parseSkipMonths_(skipMonthsRaw);
    const bysetpos = repeatPatternToBysetpos_(repeatPatternRaw);
    if (jsWeekdayIndex === null) {
      continue;
    }
    if (!(startTimeVal instanceof Date) || !(endTimeVal instanceof Date)) {
      continue;
    }

    for (let m = 0; m < 12; m++) {
      if (skipMonths.has(m)) continue; // Skip cancelled months
      const dates = computeMonthlyOccurrenceDates_(year, m, jsWeekdayIndex, bysetpos);
      for (let di = 0; di < dates.length; di++) {
        const d = dates[di];
        const start = combineDateWithTime_(d, startTimeVal);
        const end = combineDateWithTime_(d, endTimeVal);
        if (!start || !end) {
          continue;
        }
        out.push([
          approver || '',
          eventName || '',
          description || '',
          start,
          end,
          targetAudience || '',
          buildingSpaces || '',
          advertiseWhere || '',
          setupTeardown || '',
          needsGraphic || '',
          graphic || '',
          timestamp || '',
          buildingSeriesId || '',
          websiteSeriesId || '',
        ]);
      }
    }
  }

  if (out.length) {
    outSheet.getRange(2, 1, out.length, header.length).setValues(out);
  }
}


