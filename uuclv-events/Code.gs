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
 * SHEET: "Config" – approval statuses + calendar IDs
 *
 * Expected structure:
 *   A1: "Approval Statuses"
 *   A2:A4: Pending, Approved, Rejected
 *   B1: "Member Calendar ID", B2: <member calendar ID>
 *   C1: "Public Calendar ID", C2: <public calendar ID>
 *   D1: "Building Calendar ID", D2: <building calendar ID>
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
 * Helper: normalize strings for case-insensitive comparisons.
 */
function normalize_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).toLowerCase().trim();
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
 * Trigger: On form submit
 *
 * Ensures the metadata columns on "Form Responses 1" are initialized:
 *  - Column A: Approval (defaults to "Pending" if blank)
 *  - Column B: Approver (cleared)
 *  - Column Q: Building Calendar Event ID (cleared)
 *  - Column R: Website Calendar Event ID (cleared)
 *
 * Install this as an "On form submit" trigger.
 */
function onFormSubmit(e) {
  const FORM_SHEET_NAME = 'Form Responses 1';
  const APPROVAL_COL = 1;   // A
  const APPROVER_COL = 2;   // B
  const BUILDING_EVENT_ID_COL = 17; // Q
  const WEBSITE_EVENT_ID_COL = 18;  // R

  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== FORM_SHEET_NAME) {
    return;
  }

  const row = e.range.getRow();
  if (row <= 1) {
    // skip header
    return;
  }

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
}

/**
 * Trigger: On edit
 *
 * When Approval is changed to 'Approved' on 'Form Responses 1':
 *  - Fills Approver with the editor's email
 *  - Creates a Building calendar event if building space is requested
 *  - Creates a Member or Public calendar event based on target audience
 *  - Stores the Calendar Event IDs to avoid duplicates
 *
 * Install this as an "On edit" trigger on the spreadsheet.
 */
function onApprovalEdit(e) {
  const SHEET_NAME = 'Form Responses 1';
  const FIRST_DATA_ROW = 2;
  const APPROVAL_COL = 1;   // Column A: Approval
  const APPROVER_COL = 2;   // Column B: Approver

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

  if (!e || !e.range) {
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== SHEET_NAME) {
    return;
  }

  const row = range.getRow();
  if (row < FIRST_DATA_ROW) {
    // Ignore header or any rows above data
    return;
  }
  if (range.getColumn() !== APPROVAL_COL) {
    // Only respond to edits in the Approval column
    return;
  }

  const newValue = e.value || '';
  const oldValue = e.oldValue || '';

  // Only act when Approval changes TO "Approved"
  if (newValue !== 'Approved') {
    return;
  }
  if (oldValue === 'Approved') {
    // Already approved previously; avoid repeated work
    return;
  }

  const lastCol = sheet.getLastColumn();
  const rowValues = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  // Grab core fields
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

  // Get the email of the person who edited the Approval cell
  let approverEmail = '';
  if (e.user && typeof e.user.getEmail === 'function') {
    approverEmail = e.user.getEmail();
  } else {
    approverEmail = Session.getActiveUser().getEmail();
  }

  // Write approver email into column B as soon as we know it
  if (approverEmail) {
    sheet.getRange(row, APPROVER_COL).setValue(approverEmail);
    rowValues[1] = approverEmail;
  }

  // Combine base start/end times for website calendars
  const baseStartDateTime = combineDateAndTime(dateVal, startTimeVal);
  const baseEndDateTime = combineDateAndTime(dateVal, endTimeVal);

  if (!eventName || !baseStartDateTime || !baseEndDateTime) {
    // Missing critical info; cannot create any events safely
    return;
  }

  const calendarIds = getCalendarIds_();

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

      // Save Building Calendar Event ID (column Q)
      sheet
        .getRange(row, BUILDING_EVENT_ID_COL_INDEX + 1)
        .setValue(buildingEvent.getId());
    }
  }

  // --- Website calendar event (Member or Public), based on target audience ---
  const targetNorm = normalize_(targetAudience);
  let websiteCalendarId = null;

  if (targetNorm === 'members and friends') {
    websiteCalendarId = calendarIds.member;
  } else if (targetNorm === 'general public') {
    websiteCalendarId = calendarIds.public;
  } else {
    // Private Event or unknown – no website calendar event
    websiteCalendarId = null;
  }

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

      // Save Website Calendar Event ID (column R)
      sheet
        .getRange(row, WEBSITE_EVENT_ID_COL_INDEX + 1)
        .setValue(websiteEvent.getId());
    }
  }
}


