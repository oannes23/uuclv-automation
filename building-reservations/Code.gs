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
 * Trigger: On form submit
 * Copies new submissions from 'Form Responses 1' into 'Approvals',
 * with Approval defaulting to 'Pending'.
 */
function onFormSubmit(e) {
  const FORM_SHEET_NAME = 'Form Responses 1';
  const APPROVALS_SHEET_NAME = 'Approvals';

  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== FORM_SHEET_NAME) return;

  const row = e.range.getRow();
  const lastCol = sheet.getLastColumn();
  const v = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  // Map columns from Form Responses 1
  const timestamp     = v[0];  // A: Timestamp
  const eventName     = v[1];  // B: Event Name
  const dateVal       = v[2];  // C: Date
  const startTimeVal  = v[3];  // D: Start Time
  const endTimeVal    = v[4];  // E: End Time
  const spaces        = v[5];  // F: Spaces
  const keyHolder     = v[6];  // G: Key holder?
  const avHelp        = v[7];  // H: AV help?
  const extraContacts = v[8];  // I: Extra contact info
  const details       = v[9];  // J: Details / Requests
  const email         = v[10]; // K: Email Address

  const startDateTime = combineDateAndTime(dateVal, startTimeVal);
  const endDateTime   = combineDateAndTime(dateVal, endTimeVal);

  // Build Contacts field (email + optional extra contact info)
  let contacts = email || '';
  if (extraContacts && String(extraContacts).trim() !== '') {
    contacts = contacts
      ? contacts + '\n' + extraContacts
      : extraContacts;
  }

  const approvalsSheet = sheet.getParent().getSheetByName(APPROVALS_SHEET_NAME);
  if (!approvalsSheet) return;

  approvalsSheet.appendRow([
    'Pending',      // A: Approval (default)
    '',             // B: Approver (filled later on approval)
    eventName,      // C
    startDateTime,  // D
    endDateTime,    // E
    spaces,         // F
    keyHolder,      // G
    avHelp,         // H
    contacts,       // I
    details,        // J
    timestamp,      // K
    ''              // L: Calendar Event ID (filled later)
  ]);
}

/**
 * Trigger: On edit
 * When Approval is changed to 'Approved' on 'Approvals' sheet:
 *  - Fills Approver with the editor's email
 *  - Creates a Google Calendar event
 *  - Stores the Calendar Event ID to avoid duplicates
 */
function onApprovalEdit(e) {
  const SHEET_NAME = 'Approvals';
  const FIRST_DATA_ROW = 2;
  const APPROVAL_COL = 1;   // Column A
  const APPROVER_COL = 2;   // Column B
  const EVENT_ID_COL = 12;  // Column L

  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const row = range.getRow();
  if (row < FIRST_DATA_ROW) return;
  if (range.getColumn() !== APPROVAL_COL) return;

  const newValue = e.value || '';
  const oldValue = e.oldValue || '';

  // Only act when Approval changes TO "Approved"
  if (newValue !== 'Approved') return;
  if (oldValue === 'Approved') return;

  const lastCol = sheet.getLastColumn();
  const rowValues = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  // Avoid creating duplicate events
  const existingEventId = rowValues[EVENT_ID_COL - 1];
  if (existingEventId) return;

  const eventName  = rowValues[2];  // C: Event Name
  const startValue = rowValues[3];  // D: Start (Date & Time)
  const endValue   = rowValues[4];  // E: End (Date & Time)
  const spaces     = rowValues[5];  // F: Spaces (used as location)
  const contacts   = rowValues[8];  // I: Contacts
  const details    = rowValues[9];  // J: Details / Requests

  if (!eventName || !startValue || !endValue) {
    return; // missing critical info
  }

  // Get the email of the person who edited the Approval cell
  let approverEmail = '';
  if (e.user && typeof e.user.getEmail === 'function') {
    approverEmail = e.user.getEmail();
  } else {
    // May fallback to script owner, depending on domain settings
    approverEmail = Session.getActiveUser().getEmail();
  }

  // Write approver email into column B
  if (approverEmail) {
    sheet.getRange(row, APPROVER_COL).setValue(approverEmail);
  }

  // Get Calendar ID from Config sheet
  const ss = sheet.getParent();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return;

  const calendarId = configSheet.getRange('D2').getValue();
  if (!calendarId) return;

  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) return;

  // Build event description
  const descriptionParts = [];
  if (approverEmail) {
    descriptionParts.push('Approved by: ' + approverEmail);
  }
  if (contacts) {
    descriptionParts.push('Contacts:\n' + contacts);
  }
  if (details) {
    descriptionParts.push('Details / Requests:\n' + details);
  }
  const description = descriptionParts.join('\n\n');

  const startDate = new Date(startValue);
  const endDate   = new Date(endValue);

  const event = calendar.createEvent(eventName, startDate, endDate, {
    description: description,
    location: spaces || ''  // ← key change: set event location from Spaces
  });

  // Save the Calendar Event ID so we don't create duplicates later
  sheet.getRange(row, EVENT_ID_COL).setValue(event.getId());
}
