# Building Reservation System (Google Forms → Sheets → Calendar Automation)

This project implements a complete reservation + approval workflow for building use at UUCLV.  
It integrates **Google Forms**, **Google Sheets**, and **Google Calendar** using **Apps Script**, and is designed to be reliable, transparent, and easy for non-technical volunteers to manage.

This README documents the architecture, the data flow, sheet structure, formulas, and the Apps Script logic powering the automation.

---

# Overview

The system handles:

- **Building Reservation Requests** (submitted via a Google Form)
- **Automatic normalization/storage** inside a structured Google Sheet
- **Human approval workflow**, including:
  - Approval state (`Pending`, `Approved`, `Rejected`)
  - Automatic detection of the *approver’s email*
  - Logging Calendar Event IDs
- **Automatic Google Calendar event creation** when an event is approved
- **Dashboard views**:
  - `Upcoming` – next 50 approved upcoming events (sorted by start time)

The system stores all authoritative data in **Form Responses 1**, and all automations operate on that sheet so nothing is lost or duplicated.

---

# Components

## 1. Google Form — *Building Reservation Form*

The form collects:

- Event Name  
- Date of Reservation  
- Start Time  
- End Time  
- Spaces Needed (checkbox list)  
- Whether a Key Holder is needed  
- Whether A/V support is needed  
- Contact email  
- Additional contact notes  
- Event details / special requests

This form is the *single input source* for all reservation requests.

---

## 2. Google Sheet — *System Backend*

The Google Sheet contains three tabs:

### `Form Responses 1` (Authoritative Data Source)

This is the canonical database.  
Google Forms writes to its own mapped columns, and we add three **custom columns**:

| Column | Purpose |
|--------|---------|
| **A – Approval** | Default `Pending`; updated by approvers |
| **B – Approver** | Auto-filled with the user's email when they approve |
| **C → M** | Form-owned fields: Timestamp, Event Name, Date, Start Time, End Time, Spaces, Key Holder, AV Help, Extra Contacts, Details, Email |
| **N – Calendar Event ID** | Stores the created event's ID to avoid duplicates |

All automation reads/writes here.  
This avoids missed rows, race conditions, and brittle copy logic.

---


### `Upcoming` (Next 50 Approved Future Events)

A dashboard tab that shows:

- Only events with `Approval = Approved`
- Only events whose **end time is in the future**
- Sorted by upcoming start time
- Limited to 50 rows max

Used for publishing or embedding into calendars/feeds.

---

### `Config` Sheet — *Calendar Settings*

Contains:

| Cell | Value |
|------|-------|
| `A1` | Approval Status options |
| `A2` - `A4` | Pending, Approved, and Rejected status |
| `D1` | `Calendar ID` |
| `D2` | The actual Google Calendar ID where events should be created |

The A column creates the dropdown options available on the A column of the `Form Responses 1` tab. The D column information separation makes it easy to change which calendar receives reservations.

---

# Apps Script Logic

All automation lives inside one Apps Script project with **two installable triggers** and some helper functions. You can see the code in **[`Code.gs`](Code.gs)**

### Trigger 1 — `onFormSubmit`
**Event:** “On form submit”  
**Purpose:**
- Sets `Approval = Pending` for the new entry
- Ensures rows are initialized consistently
- Avoids flaky copy triggers by operating on the source sheet directly

### Trigger 2 — `onApprovalEdit`
**Event:** “On edit”  
**Purpose:**
- Detects when Approval changes to `Approved`
- Adds the approver’s email into the `Approver` column
- Creates a Google Calendar event:
  - Title = Event Name
  - Start/End datetime = combined date + times
  - Description includes:
    - Approved by email
    - Contact info
    - Event details
  - Location = Spaces column
- Stores the Calendar Event ID so events are **never duplicated**

Both triggers are stable and row-order independent because they operate on the real form data table.

---

# Data Flow (End-to-End)

1. **User submits Building Reservation Form**
   - Google Forms writes a new row to `Form Responses 1`.

2. **onFormSubmit Trigger Fires**
   - Sets `Approval = Pending` for that row.
   - All other data remains untouched.

3. **Approver reviews events**
   - They may filter `Form Responses 1` or use the `Approvals` tab.
   - To approve, they change the Approval cell from `Pending` → `Approved`.

4. **onApprovalEdit Trigger Fires**
   - Detects Approval → Approved.
   - Fills Approver email.
   - Builds and creates a Google Calendar event.
   - Writes the Calendar Event ID back to `Form Responses 1`.

5. **Dashboard Views update automatically**
   - `Approvals` refreshes via ARRAYFORMULA.
   - `Upcoming` updates with filtered upcoming approved events.

6. **Calendar event now appears on the UUCLV calendar**  
   All relevant details (location, contacts, approver, description) appear automatically.

---

# Sheet Formulas

Upcoming Sheet Formula
```gs
=ARRAYFORMULA({
  {"Approval","Approver","Event Name","Start","End","Spaces",
   "Key holder needed?","AV help needed?","Contacts",
   "Details / Requests","Form Timestamp","Calendar Event ID"};
  ARRAY_CONSTRAIN(
    SORT(
      FILTER(
        Approvals!A2:L,
        Approvals!A2:A="Approved",
        Approvals!E2:E >= NOW()
      ),
      4, TRUE
    ),
    50, 12
  )
})
```

⸻

Apps Script (Code.gs)

The repository should include:
	•	Code.gs – containing:
	•	combineDateAndTime(date, time)
	•	onFormSubmit(e)
	•	onApprovalEdit(e)
	•	Full documentation on expected column mapping
	•	Instructions for deploying triggers

⸻

Permissions Required

The Apps Script project requires:
	•	Spreadsheet read/write
	•	Calendar creation permission
	•	Access to the user’s email (to log Approver identity)

The first time triggers run, Google will prompt for authorization.

⸻

Recommended Workflow for Volunteers
	1.	Open the Form Responses 1 sheet.
	2.	Use the filter view: Approval = Pending.
	3.	Review details (date, space, AV needs, etc.).
	4.	If approved, change Approval → Approved.
	5.	Event automatically appears on the UUCLV Calendar.
	6.	Check Calendar Event ID column for confirmation.

⸻

Future Enhancements (Optional)
	•	Automatic cancellation of calendar events if Approval changes to Rejected.
	•	Email notifications to requesters upon approval or rejection.
	•	A cleaner front-end UI for approvers (via a Google Sites widget or custom web interface).
	•	A second form (large event form) that normalizes into the same workflow.
	•	Automatic conflict detection (overlapping spaces/time windows).
	•	Integrating with the UUCLV Event Feed system.

⸻

Repository Structure
```
/
├── README.md          # This document
└── Code.gs            # Full Apps Script implementation
```
⸻

End-to-End Summary

The Building Reservation System now supports:
	•	Automated intake
	•	Reliable state-tracking
	•	Human approval workflow
	•	Automatic Google Calendar publishing
	•	Dashboard views
	•	No duplicate events
	•	No missing rows

It is stable, scalable, and maintainable — and this repo will serve as the central documentation hub for volunteers and future developers.

