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
 * Ensures the metadata columns on "Form Responses 1" are initialized:
 *  - Column A: Approval (defaults to "Pending")
 *  - Column B: Approver (cleared)
 *  - Column N: Calendar Event ID (cleared)
 */
function onFormSubmit(e) {
  const FORM_SHEET_NAME = 'Form Responses 1';
  const APPROVAL_COL = 1;   // A: Approval
  const APPROVER_COL = 2;   // B: Approver
  const EVENT_ID_COL = 14;  // N: Calendar Event ID

  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== FORM_SHEET_NAME) return;

  const row = e.range.getRow();
  if (row <= 1) return; // skip header

  // Initialize Approval column to "Pending" if blank
  const approvalCell = sheet.getRange(row, APPROVAL_COL);
  if (!approvalCell.getValue()) {
    approvalCell.setValue('Pending');
  }

  // Clear Approver (B) if this is a brand‑new submission
  const approverCell = sheet.getRange(row, APPROVER_COL);
  if (!approverCell.getValue()) {
    approverCell.setValue('');
  }

  // Clear Calendar Event ID (N) so re-submissions never reuse stale IDs
  const eventIdCell = sheet.getRange(row, EVENT_ID_COL);
  if (!eventIdCell.getValue()) {
    eventIdCell.setValue('');
  }
}

/**
 * Trigger: On edit
 * When Approval is changed to 'Approved' on 'Form Responses 1' sheet:
 *  - Fills Approver with the editor's email
 *  - Creates a Google Calendar event
 *  - Stores the Calendar Event ID to avoid duplicates
 */
function onApprovalEdit(e) {
  const SHEET_NAME = 'Form Responses 1';
  const FIRST_DATA_ROW = 2;
  const APPROVAL_COL = 1;   // Column A: Approval
  const APPROVER_COL = 2;   // Column B: Approver

  // "Form Responses 1" column layout (zero-based indexes in rowValues[]):
  //  0: A - Approval
  //  1: B - Approver
  //  2: C - Timestamp
  //  3: D - Event Name
  //  4: E - Date
  //  5: F - Start Time
  //  6: G - End Time
  //  7: H - Spaces
  //  8: I - Key holder needed?
  //  9: J - AV help needed?
  // 10: K - Extra Contacts
  // 11: L - Details / Requests
  // 12: M - Email
  // 13: N - Calendar Event ID
  const TIMESTAMP_COL_INDEX = 2;
  const EVENT_NAME_COL_INDEX = 3;
  const DATE_COL_INDEX = 4;
  const START_TIME_COL_INDEX = 5;
  const END_TIME_COL_INDEX = 6;
  const SPACES_COL_INDEX = 7;
  const EXTRA_CONTACTS_COL_INDEX = 10;
  const DETAILS_COL_INDEX = 11;
  const EMAIL_COL_INDEX = 12;
  const EVENT_ID_COL_INDEX = 13;

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
  const existingEventId = rowValues[EVENT_ID_COL_INDEX];
  if (existingEventId) return;

  const eventName     = rowValues[EVENT_NAME_COL_INDEX];      // D: Event Name
  const dateVal       = rowValues[DATE_COL_INDEX];            // E: Date
  const startTimeVal  = rowValues[START_TIME_COL_INDEX];      // F: Start Time
  const endTimeVal    = rowValues[END_TIME_COL_INDEX];        // G: End Time
  const spaces        = rowValues[SPACES_COL_INDEX];          // H: Spaces (used as location)
  const extraContacts = rowValues[EXTRA_CONTACTS_COL_INDEX];  // K: Extra Contacts
  const details       = rowValues[DETAILS_COL_INDEX];         // L: Details / Requests
  const email         = rowValues[EMAIL_COL_INDEX];           // M: Email

  const startDateTime = combineDateAndTime(dateVal, startTimeVal);
  const endDateTime   = combineDateAndTime(dateVal, endTimeVal);

  if (!eventName || !startDateTime || !endDateTime) {
    return; // missing critical info
  }

  // Build Contacts field (email + optional extra contact info)
  let contacts = email || '';
  if (extraContacts && String(extraContacts).trim() !== '') {
    contacts = contacts
      ? contacts + '\n' + extraContacts
      : extraContacts;
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
    // Also refresh the cached rowValues entry for Approver (index 1) if needed later
    rowValues[1] = approverEmail;
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

  const event = calendar.createEvent(eventName, startDateTime, endDateTime, {
    description: description,
    location: spaces || ''  // ← key change: set event location from Spaces
  });

  // Save the Calendar Event ID so we don't create duplicates later
  sheet.getRange(row, EVENT_ID_COL_INDEX + 1).setValue(event.getId());
}
