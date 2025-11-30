## Event Feed System (Google Forms → Sheets → Calendar → Website)

This project implements an end‑to‑end **event promotion workflow** for UUCLV.  
It integrates **Google Forms**, **Google Sheets**, and **Google Calendar** (feeding into the website event views), and is designed to be reliable, transparent, and easy for non‑technical volunteers to manage.

This README documents the architecture, the data flow, sheet structure, formulas, and the Apps Script logic powering the automation.

---

## Overview

The system handles:

- **Event submissions** (via a Google Form)
- **Automatic normalization/storage** inside a structured Google Sheet
- **Human approval workflow**, including:
  - Approval state (`Pending`, `Approved`, `Rejected`)
  - Automatic detection of the *approver’s email*
  - Logging Calendar Event IDs
- **Automatic Google Calendar event creation** to one of **two calendars**:
  - **Public calendar**
  - **Member calendar**
- **Conditional event creation**:
  - A calendar event is created **only if** “**Add to Website Calendar**” is selected in the **How should we promote this event?** field
- **Multiple dashboard/view tabs**, including:
  - `All Upcoming`
  - `Members`
  - `Public`
  - `Social Media`
  - `Flash`
  - `Sunday`
  - `Website`

All authoritative data is stored in the **Form Responses 1** sheet, and all automations operate on that sheet so nothing is lost or duplicated.

---

## Components

### 1. Google Form — *Event Submission Form*

The form collects:

- Event Name  
- Event Description (including contact info)  
- Event Date  
- Event Start Time  
- Event End Time  
- Target audience (e.g., Public vs Members)  
- How the event should be promoted (multi‑select, including “Add to Website Calendar”)  
- Whether a graphic is needed  
- Optional uploaded graphic  
- Contact email

This form is the *single input source* for all event submissions.

---

### 2. Google Sheet — *System Backend*

The Google Sheet contains:

- `Form Responses 1` – canonical data table
- `Config` – configuration and calendar routing
- View tabs: `All Upcoming`, `Members`, `Public`, `Social Media`, `Flash`, `Sunday`, `Website`

#### `Form Responses 1` (Authoritative Data Source)

This is the canonical database. Google Forms writes to its own mapped columns, and we add two **custom columns** at the front plus one tracking column at the end:

| Column | Header                                                                                           | Purpose |
|--------|---------------------------------------------------------------------------------------------------|---------|
| **A**  | **Approval**                                                                                      | Default `Pending`; updated by approvers to `Approved` or `Rejected` |
| **B**  | **Approver**                                                                                      | Auto‑filled with the approver’s email when they approve |
| **C**  | **Timestamp**                                                                                     | Google Form submission timestamp |
| **D**  | **Email Address**                                                                                | Contact email for the submitter |
| **E**  | **Event Name**                                                                                   | Title of the event |
| **F**  | **Event Description (please include contact info for interested folks to ask questions)**        | Event details and public‑facing description |
| **G**  | **Event Date**                                                                                   | Date of the event |
| **H**  | **Event Start Time**                                                                             | Start time on the given date |
| **I**  | **Event End Time**                                                                               | End time on the given date |
| **J**  | **Who is the target audience for this event?**                                                   | Used to route to Member vs Public views/calendars |
| **K**  | **How should we promote this event?**                                                            | Multi‑select promotion options, including “Add to Website Calendar” |
| **L**  | **Do you need a graphic created for this event?**                                                | Yes/No or similar |
| **M**  | **Upload graphic if you already have one**                                                       | File upload link / identifier |
| **N**  | **Calendar Event ID**                                                                            | Stores the created event’s Calendar ID to avoid duplicates |

All automation reads/writes here. This avoids missed rows, race conditions, and brittle copy logic.

---

#### `Config` Sheet — *Approval + Calendar Routing*

The `Config` sheet controls **approval statuses** and **which calendar** to send events to based on status/target.

- **Column A – Approval Statuses**: Allowed statuses (e.g., `Pending`, `Approved`, `Rejected`). Used as **data validation** options in `Form Responses 1` → `Approval` column.
- **Column B – Target**: High‑level target segment (e.g., `Public`, `Members`) used to decide which calendar ID to apply.
- **Column C – Member Calendar ID**: Calendar ID for member‑only events.
- **Column D – Public Calendar ID**: Calendar ID for public‑facing events.

The Apps Script reads these cells to determine:

- Which statuses are valid.
- Which calendar (Member vs Public) to send a given `Approved` event to, based on the target audience and configuration.

---

### 3. View Tabs — *Website & Communications Feeds*

Each view tab is built on top of `Form Responses 1` using formulas (typically `ARRAYFORMULA`, `FILTER`, `SORT`, etc.).  
All views:

- Show only events where **Approval = Approved**
- Show only events whose **end time is in the future**
- Are sorted by upcoming start time

Individual views then apply additional filters:

- **`All Upcoming`**: All upcoming approved events (baseline dataset for the other views).
- **`Members`**: Subset of upcoming events where target audience is member‑only or internal.
- **`Public`**: Subset of upcoming events appropriate for the public calendar/website.
- **`Social Media`**: Events marked for social media promotion.
- **`Flash`**: Events marked for “flash” / short‑notice promotion.
- **`Sunday`**: Events that should be announced in Sunday‑specific channels (e.g., order of service announcements).
- **`Website`**: Events that should appear on the website (often those with “Add to Website Calendar” selected).

These views can be published or embedded as data sources for the UUCLV site or other tools.

---

## Apps Script Logic

All automation lives inside one Apps Script project with two main triggers and supporting helper functions. The code lives in **`Code.gs`**.

### Trigger 1 — `onFormSubmit`

**Event:** “On form submit”  
**Purpose:**

- Sets `Approval = Pending` for the new entry
- Ensures rows are initialized consistently
- Avoids brittle copy logic by operating directly on `Form Responses 1`

### Trigger 2 — `onApprovalEdit`

**Event:** “On edit”  
**Purpose:**

- Detects when **Approval changes to `Approved`**
- Fills the **Approver** column with the active user’s email
- Checks the **How should we promote this event?** field:
  - **If it does *not* include “Add to Website Calendar”** → **no calendar event is created**
  - **If it includes “Add to Website Calendar”** → a calendar event is created
- Determines which calendar to use:
  - Reads the target audience from `Who is the target audience for this event?`
  - Uses the `Config` sheet to map target → Member or Public calendar ID
- Creates a Google Calendar event:
  - Title = Event Name
  - Start/End datetime = combined Event Date + Start/End Time
  - Description includes:
    - Event description text
    - Contact information
    - Approval information (approver email)
  - Location (if applicable in future enhancements) can be inferred or added later
- Stores the Calendar Event ID back into `Calendar Event ID` so events are **never duplicated**

Both triggers are stable and row‑order independent because they operate on the real form data table.

---

## Data Flow (End‑to‑End)

1. **User submits the Event Form**
   - Google Forms writes a new row to `Form Responses 1`.
2. **`onFormSubmit` Trigger Fires**
   - Sets `Approval = Pending` for that row.
   - All other data remains untouched.
3. **Approver reviews events**
   - They can filter `Form Responses 1` or use one of the view tabs as a dashboard.
   - To approve, they change the Approval cell from `Pending` → `Approved`.
4. **`onApprovalEdit` Trigger Fires**
   - Detects Approval → Approved.
   - Fills Approver email.
   - Checks “How should we promote this event?” for **“Add to Website Calendar”**.
     - If selected, determines the correct calendar (Member vs Public) from the `Config` sheet and creates the event.
     - If not selected, **no calendar event** is created, but the row remains approved and available in views.
   - Writes the Calendar Event ID back to `Form Responses 1` (if an event was created).
5. **View Tabs update automatically**
   - `All Upcoming`, `Members`, `Public`, `Social Media`, `Flash`, `Sunday`, and `Website` tabs refresh via formulas.
6. **Calendar event appears on the appropriate UUCLV calendar**  
   - Website and other systems can now read from the relevant views to display upcoming events.

---

## Sheet Formulas (Conceptual)

The view tabs are each powered by formulas that:

- Pull from `Form Responses 1`
- Filter on:
  - `Approval = "Approved"`
  - Future end times (`Event Date` + `Event End Time` >= `NOW()`)
  - Target audience and/or promotion selections, depending on the tab
- Sort by start datetime

A typical pattern (for `All Upcoming`) will look similar to:

```gs
=ARRAYFORMULA({
  {"Approver","Event Name","Description","Start","End","Target",
   "Promotion","Needs Graphic?","Graphic",
   "Form Timestamp","Calendar Event ID"};
  ARRAY_CONSTRAIN(
    SORT(
      FILTER(
        {
          'Form Responses 1'!B2:B,
          'Form Responses 1'!E2:E,
          'Form Responses 1'!F2:F,
          'Form Responses 1'!G2:G + 'Form Responses 1'!H2:H,
          'Form Responses 1'!G2:G + 'Form Responses 1'!I2:I,
          'Form Responses 1'!J2:J,
          'Form Responses 1'!K2:K,
          'Form Responses 1'!L2:L,
          'Form Responses 1'!M2:M,
          'Form Responses 1'!C2:C,
          'Form Responses 1'!N2:N
        },
        'Form Responses 1'!A2:A="Approved",
        ('Form Responses 1'!G2:G + 'Form Responses 1'!I2:I) >= NOW()
      ),
      3, TRUE
    ),
    200, 11
  )
})
```

You can then adapt this pattern for more focused views.

For example, a **Flash-only** view (where the promotion options include the word “Flash”) could look like:

```gs
=ARRAYFORMULA({
  {"Approver","Event Name","Description","Start","End","Target",
   "Promotion","Needs Graphic?","Graphic",
   "Form Timestamp","Calendar Event ID"};
  ARRAY_CONSTRAIN(
    SORT(
      FILTER(
        {
          'Form Responses 1'!B2:B,
          'Form Responses 1'!E2:E,
          'Form Responses 1'!F2:F,
          'Form Responses 1'!G2:G + 'Form Responses 1'!H2:H,
          'Form Responses 1'!G2:G + 'Form Responses 1'!I2:I,
          'Form Responses 1'!J2:J,
          'Form Responses 1'!K2:K,
          'Form Responses 1'!L2:L,
          'Form Responses 1'!M2:M,
          'Form Responses 1'!C2:C,
          'Form Responses 1'!N2:N
        },
        'Form Responses 1'!A2:A="Approved",
        ('Form Responses 1'!G2:G + 'Form Responses 1'!I2:I) >= NOW(),
        IFERROR(REGEXMATCH('Form Responses 1'!K2:K,"Flash"),FALSE)
      ),
      3, TRUE
    ),
    200, 11
  )
})
```

And a **Members-only** view (where the target audience column equals `Members`) could look like:

```gs
=ARRAYFORMULA({
  {"Approver","Event Name","Description","Start","End","Target",
   "Promotion","Needs Graphic?","Graphic",
   "Form Timestamp","Calendar Event ID"};
  ARRAY_CONSTRAIN(
    SORT(
      FILTER(
        {
          'Form Responses 1'!B2:B,
          'Form Responses 1'!E2:E,
          'Form Responses 1'!F2:F,
          'Form Responses 1'!G2:G + 'Form Responses 1'!H2:H,
          'Form Responses 1'!G2:G + 'Form Responses 1'!I2:I,
          'Form Responses 1'!J2:J,
          'Form Responses 1'!K2:K,
          'Form Responses 1'!L2:L,
          'Form Responses 1'!M2:M,
          'Form Responses 1'!C2:C,
          'Form Responses 1'!N2:N
        },
        'Form Responses 1'!A2:A="Approved",
        ('Form Responses 1'!G2:G + 'Form Responses 1'!I2:I) >= NOW(),
        IFERROR(REGEXMATCH('Form Responses 1'!K2:K,"Members"),FALSE)
      ),
      3, TRUE
    ),
    200, 11
  )
})
```

Other view tabs (e.g., `Public`, `Social Media`, `Sunday`, `Website`) use the same pattern with **extra filter conditions**, such as:

- Target audience equals a specific segment (Members vs Public)
- Promotion column contains a specific keyword or option

---

## Timezone Configuration (Pacific Time)

For event times to appear correctly on the calendars and website, make sure all of the following are set to **Pacific Time (America/Los_Angeles)**:

- **Spreadsheet timezone**:  
  - In the `Form Responses 1` spreadsheet, go to **File → Settings → Locale / Time zone** and set the time zone to Pacific.
- **Apps Script project timezone**:  
  - In the Apps Script editor, go to **Project Settings → Script time zone** and set it to Pacific.
- **Target Google Calendars timezone**:  
  - In Google Calendar, open **Settings → Time zone** for the Public and Member calendars and set them to Pacific.

When all three are aligned to Pacific, the datetime calculations and created events will match the times selected on the form.

---

## Apps Script (Code.gs)

The repository should include:

- `Code.gs` – containing:
  - `combineDateAndTime(date, time)`
  - `onFormSubmit(e)`
  - `onApprovalEdit(e)`
  - Helper(s) for reading the `Config` sheet and choosing the correct calendar
  - Full documentation on expected column mapping
  - Instructions for deploying triggers

---

## Permissions Required

The Apps Script project requires:

- Spreadsheet read/write
- Calendar creation permission (for both Member and Public calendars)
- Access to the user’s email (to log Approver identity)

The first time triggers run, Google will prompt for authorization.

---

## Recommended Workflow for Volunteers

1. Open the `Form Responses 1` sheet.
2. Use a filter view: `Approval = Pending`.
3. Review details (date, time, target audience, promotion channels, description, graphics needs).
4. If approved, change `Approval` → `Approved`.
5. If “Add to Website Calendar” is selected, the event will automatically appear on the appropriate UUCLV calendar (Public or Member).
6. Check the `Calendar Event ID` column for confirmation that a calendar event was created (when applicable).

---

## Future Enhancements (Optional)

- Automatic cancellation/update of calendar events if `Approval` changes from `Approved` to `Rejected` (or details change).
- Email notifications to requesters upon approval or rejection.
- A cleaner front‑end UI for approvers (via a Google Sites widget or custom web interface).
- Additional specialized views for communications (e.g., weekly digests, print‑ready exports).
- Integration with the existing UUCLV website rendering logic for the event feed.

---

## Repository Structure

```text
/
├── README.md          # This document
└── Code.gs            # Apps Script implementation for the Event Feed workflow
```

---

## End‑to‑End Summary

The Event Feed system supports:

- Automated intake of event submissions
- Reliable state‑tracking and approvals
- Conditional calendar publishing (only when “Add to Website Calendar” is selected)
- Routing to the appropriate Public or Member calendar
- Multiple view tabs for different communications channels (All Upcoming, Members, Public, Social Media, Flash, Sunday, Website)
- A maintainable, documented workflow for volunteers and future developers


