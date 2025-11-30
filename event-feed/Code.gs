/**
 * Event Feed Automation for UUCLV
 *
 * This Apps Script implements the workflow described in the Event Feed README:
 *  - Google Form → "Form Responses 1" sheet (authoritative data)
 *  - Approval workflow on column A ("Approval")
 *  - Conditional Calendar event creation when:
 *      - Approval changes to "Approved"
 *      - AND "How should we promote this event?" includes "Add to Website Calendar"
 *  - Calendar routing (Member vs Public) based on target audience + Config sheet
 *  - Logging the Calendar Event ID so events are never duplicated
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
 *  10 /  9 : J - Who is the target audience for this event?
 *  11 / 10 : K - How should we promote this event?
 *  12 / 11 : L - Do you need a graphic created for this event?
 *  13 / 12 : M - Upload graphic if you already have one
 *  14 / 13 : N - Calendar Event ID
 *
 * SHEET: "Config" – approval statuses + calendar routing
 *
 * Expected structure (starting at row 2):
 *   Col A – Approval Statuses (e.g., Pending, Approved, Rejected)
 *   Col B – Target (e.g., "Public", "Members", etc.)
 *   Col C – Member Calendar ID
 *   Col D – Public Calendar ID
 *
 * The script reads this sheet to:
 *  - Determine which calendar ID to use given a target audience string.
 *  - Support flexible mapping (multiple targets can be routed to Member or Public).
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
 * Helper: returns true if the promotion selections include
 * "Add to Website Calendar" (case-insensitive substring match).
 */
function shouldAddToWebsiteCalendar_(promotionValue) {
  const text = normalize_(promotionValue);
  if (!text) return false;
  return text.indexOf('add to website calendar') !== -1;
}

/**
 * Helper: get the Config sheet or null if not found.
 */
function getConfigSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Config');
}

/**
 * Helper: Resolve a calendar ID from the Config sheet based on the
 * target audience string in "Form Responses 1".
 *
 * Strategy:
 *  - Look through Config!B:D (starting at row 2).
 *  - For each row:
 *      B: target label
 *      C: Member Calendar ID (optional)
 *      D: Public Calendar ID (optional)
 *  - If B matches the normalized target string, prefer:
 *      - the non-empty calendar ID on that row
 *      - if both C and D are present, choose based on keywords in the target:
 *          * contains "member" → member calendar
 *          * contains "public" → public calendar
 *          * otherwise → public calendar (sensible default)
 *  - If no direct label match, fall back to:
 *      - Any first-seen Member/Public IDs in the table, keyed off keywords
 *        in the target string ("member" vs "public").
 *      - Ultimately prefer Public over Member if still ambiguous.
 */
function getCalendarIdForTarget_(targetRaw) {
  const configSheet = getConfigSheet_();
  if (!configSheet) {
    Logger.log('Config sheet not found; cannot route to calendar.');
    return null;
  }

  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('Config sheet has no data rows; cannot route to calendar.');
    return null;
  }

  const numRows = lastRow - 1;
  const data = configSheet.getRange(2, 1, numRows, 4).getValues();

  const targetNorm = normalize_(targetRaw);
  let defaultMemberId = null;
  let defaultPublicId = null;

  // First pass: find exact target label match and collect defaults.
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const targetLabel = normalize_(row[1]); // Col B – Target
    const memberId = row[2];                // Col C – Member Calendar ID
    const publicId = row[3];                // Col D – Public Calendar ID

    if (memberId && !defaultMemberId) {
      defaultMemberId = memberId;
    }
    if (publicId && !defaultPublicId) {
      defaultPublicId = publicId;
    }

    if (!targetLabel || targetLabel !== targetNorm) {
      continue;
    }

    // Direct target match; decide which calendar ID on this row to use.
    if (memberId && !publicId) {
      return memberId;
    }
    if (publicId && !memberId) {
      return publicId;
    }
    if (memberId && publicId) {
      // Both provided; decide based on keywords in the target name.
      if (targetNorm.indexOf('member') !== -1) {
        return memberId;
      }
      if (targetNorm.indexOf('public') !== -1) {
        return publicId;
      }
      // If ambiguous, default to public calendar.
      return publicId;
    }
  }

  // No direct label match; fall back using keywords in the target text.
  if (targetNorm.indexOf('member') !== -1 && defaultMemberId) {
    return defaultMemberId;
  }
  if (targetNorm.indexOf('public') !== -1 && defaultPublicId) {
    return defaultPublicId;
  }

  // Final fallback: prefer public, then member.
  return defaultPublicId || defaultMemberId || null;
}

/**
 * Trigger: On form submit
 *
 * Ensures the metadata columns on "Form Responses 1" are initialized:
 *  - Column A: Approval (defaults to "Pending" if blank)
 *  - Column B: Approver (cleared)
 *  - Column N: Calendar Event ID (cleared)
 *
 * This trigger should be installed as an "On form submit" trigger.
 */
function onFormSubmit(e) {
  const FORM_SHEET_NAME = 'Form Responses 1';
  const APPROVAL_COL = 1;   // A: Approval
  const APPROVER_COL = 2;   // B: Approver
  const EVENT_ID_COL = 14;  // N: Calendar Event ID

  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== FORM_SHEET_NAME) {
    return;
  }

  const row = e.range.getRow();
  if (row <= 1) {
    // Skip header row and any rows above.
    return;
  }

  // Initialize Approval column to "Pending" if blank.
  const approvalCell = sheet.getRange(row, APPROVAL_COL);
  if (!approvalCell.getValue()) {
    approvalCell.setValue('Pending');
  }

  // Clear Approver (B) for a brand-new submission.
  const approverCell = sheet.getRange(row, APPROVER_COL);
  if (approverCell.getValue()) {
    approverCell.setValue('');
  }

  // Clear Calendar Event ID (N) so re-submissions never reuse stale IDs.
  const eventIdCell = sheet.getRange(row, EVENT_ID_COL);
  if (eventIdCell.getValue()) {
    eventIdCell.setValue('');
  }
}

/**
 * Trigger: On edit
 *
 * When Approval is changed to 'Approved' on 'Form Responses 1':
 *  - Fills Approver with the editor's email
 *  - Checks "How should we promote this event?" for "Add to Website Calendar"
 *      * If NOT selected → no calendar event is created
 *      * If selected → creates a Google Calendar event
 *  - Routes the event to the correct calendar (Member vs Public) using Config sheet
 *  - Stores the Calendar Event ID to avoid duplicates
 *
 * This trigger should be installed as an "On edit" trigger that listens
 * on the same spreadsheet as the "Form Responses 1" sheet.
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
  //  3: D - Email Address
  //  4: E - Event Name
  //  5: F - Event Description
  //  6: G - Event Date
  //  7: H - Event Start Time
  //  8: I - Event End Time
  //  9: J - Target Audience
  // 10: K - Promotion selections
  // 11: L - Needs graphic?
  // 12: M - Upload graphic
  // 13: N - Calendar Event ID
  const EMAIL_COL_INDEX = 3;
  const EVENT_NAME_COL_INDEX = 4;
  const DESCRIPTION_COL_INDEX = 5;
  const DATE_COL_INDEX = 6;
  const START_TIME_COL_INDEX = 7;
  const END_TIME_COL_INDEX = 8;
  const TARGET_COL_INDEX = 9;
  const PROMOTION_COL_INDEX = 10;
  const EVENT_ID_COL_INDEX = 13;

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
    // Ignore header or any rows above the data range.
    return;
  }
  if (range.getColumn() !== APPROVAL_COL) {
    // Only respond to edits in the Approval column.
    return;
  }

  const newValue = e.value || '';
  const oldValue = e.oldValue || '';

  // Only act when Approval changes TO "Approved".
  if (newValue !== 'Approved') {
    return;
  }
  if (oldValue === 'Approved') {
    // Already approved previously; avoid repeated work.
    return;
  }

  const lastCol = sheet.getLastColumn();
  const rowValues = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const existingEventId = rowValues[EVENT_ID_COL_INDEX];
  const emailAddress = rowValues[EMAIL_COL_INDEX];
  const eventName = rowValues[EVENT_NAME_COL_INDEX];
  const eventDescription = rowValues[DESCRIPTION_COL_INDEX];
  const dateVal = rowValues[DATE_COL_INDEX];
  const startTimeVal = rowValues[START_TIME_COL_INDEX];
  const endTimeVal = rowValues[END_TIME_COL_INDEX];
  const targetAudience = rowValues[TARGET_COL_INDEX];
  const promotionSelections = rowValues[PROMOTION_COL_INDEX];

  // Get the email of the person who edited the Approval cell.
  let approverEmail = '';
  if (e.user && typeof e.user.getEmail === 'function') {
    approverEmail = e.user.getEmail();
  } else {
    // Fallback to active user; depending on domain settings this may be blank.
    approverEmail = Session.getActiveUser().getEmail();
  }

  // Write approver email into column B as soon as we know it.
  if (approverEmail) {
    sheet.getRange(row, APPROVER_COL).setValue(approverEmail);
    // Also refresh the cached rowValues entry for Approver (index 1) if needed.
    rowValues[1] = approverEmail;
  }

  // If we already have an Event ID, do not create another calendar event.
  // We still allow the Approver cell to update above.
  if (existingEventId) {
    return;
  }

  // Check promotion selections: only create a calendar event if the
  // user requested "Add to Website Calendar".
  if (!shouldAddToWebsiteCalendar_(promotionSelections)) {
    // Row is still fully approved and will appear in views; it simply
    // does not get a calendar event.
    return;
  }

  const startDateTime = combineDateAndTime(dateVal, startTimeVal);
  const endDateTime = combineDateAndTime(dateVal, endTimeVal);

  if (!eventName || !startDateTime || !endDateTime) {
    // Missing critical info; do not attempt to create the event.
    return;
  }

  // Resolve the appropriate calendar ID based on target audience.
  const calendarId = getCalendarIdForTarget_(targetAudience);
  if (!calendarId) {
    Logger.log('No calendar ID resolved for target "' + targetAudience + '".');
    return;
  }

  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    Logger.log('Calendar not found for ID: ' + calendarId);
    return;
  }

  // Build event description:
  //  - Event description (which should already include some contact info)
  //  - Contact email (from Email Address column)
  //  - Approval information (approver email)
  const descriptionParts = [];
  if (eventDescription) {
    descriptionParts.push('Description:\n' + eventDescription);
  }
  if (emailAddress) {
    descriptionParts.push('Contact email: ' + emailAddress);
  }
  if (approverEmail) {
    descriptionParts.push('Approved by: ' + approverEmail);
  }
  const description = descriptionParts.join('\n\n');

  // Create the calendar event.
  const event = calendar.createEvent(eventName, startDateTime, endDateTime, {
    description: description
    // Location can be added in future enhancements when a dedicated
    // location/space field is introduced to the form/sheet.
  });

  // Save the Calendar Event ID so we don't create duplicates later.
  sheet.getRange(row, EVENT_ID_COL_INDEX + 1).setValue(event.getId());
}


