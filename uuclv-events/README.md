## UUCLV Events System (Combined Building Reservations + Event Feed)

This project implements a **unified events workflow** for UUCLV, combining:

- **Building reservations**
- **Member / Public website event publishing**

It integrates **Google Forms**, **Google Sheets**, and **Google Calendar** using **Apps Script**, and is designed so that non‑technical volunteers can manage everything from a single sheet.

This README documents the architecture, sheet structure, formulas, and the Apps Script logic powering the automation.

---

## Overview

The combined system handles:

- **Event submissions** (via a single Google Form)
- **Building use requests**
- **Automatic normalization/storage** inside `Form Responses 1`
- **Human approval workflow**, including:
  - Approval state (`Pending`, `Approved`, `Rejected`)
  - Automatic detection of the *approver’s email*
  - Logging of **Building Calendar Event IDs** and **Website Calendar Event IDs**
- **Automatic Google Calendar event creation** when an event is approved:
  - **Building Reservation calendar** (only if building space was requested)
  - **Member Calendar** (for “Members and Friends”)
  - **Public Calendar** (for “General Public”)
- **Automatic setup/teardown padding** for building reservations
- A basic **`All Upcoming`** view tab for downstream use (website, communications, etc.)

All authoritative data is stored in the **`Form Responses 1`** sheet, and all automations operate on that sheet so nothing is lost or duplicated.

---

## Components

### 1. Google Form — `uuclv-events`

The unified form collects:

- Event Name  
- Event Description (including contact info)  
- Event Date  
- Event Start Time  
- Event End Time  
- **What part(s) of the building do you want to use?** (checkboxes)  
- **Who is the target audience for this event?**
  - Options should include at least:
    - `Private Event`
    - `Members and Friends`
    - `General Public`
- **If this event is not Private, where should we advertise it?**
  - Multi‑select promotion channels (newsletter, social media, etc.)
  - **Note:** The website is **not** listed here; website visibility is controlled only by the calendar events.
- **If you are requesting building space, how much setup and teardown time do you need before/after your event?**
  - Options:
    - `None`
    - `30 Minutes`
    - `1 Hour`
    - `2 Hours`
- **If you are requesting building space, do you need a key holder to open and close the building or AV support?**
- **Do you need someone to create a graphic for this event?**
- **If you have your own graphic already, please upload it here**
- Email Address (for contact + approvals)

This is the *single* input source for all building and website events.

---

### 2. Google Sheet — System Backend

The main spreadsheet contains at least:

- `Form Responses 1` – canonical data table
- `Config` – approval statuses and calendar IDs
- `All Upcoming` – combined, filtered list of approved upcoming events

#### `Form Responses 1` (Authoritative Data Source)

Columns are as follows (1‑based column letters in parentheses):

1. **Approval** (`A`)  
2. **Approver** (`B`)  
3. **Timestamp** (`C`)  
4. **Email Address** (`D`)  
5. **Event Name** (`E`)  
6. **Event Description (please include contact info for interested folks to ask questions)** (`F`)  
7. **Event Date** (`G`)  
8. **Event Start Time** (`H`)  
9. **Event End Time** (`I`)  
10. **What part(s) of the building do you want to use?** (`J`)  
11. **Who is the target audience for this event?** (`K`)  
12. **If this event is not Private, where should we advertise it?** (`L`)  
13. **If you are requesting building space, how much setup and teardown time do you need before/after your event?** (`M`)  
14. **If you are requesting building space, do you need a key holder to open and close the building or AV support?** (`N`)  
15. **Do you need someone to create a graphic for this event?** (`O`)  
16. **If you have your own graphic already, please upload it here** (`P`)  
17. **Building Calendar Event ID** (`Q`)  
18. **Website Calendar Event ID** (`R`)  

Key behaviors:

- Column **A – Approval**
  - Defaults to `Pending` on form submit (set by Apps Script).
  - Approvers change this to `Approved` or `Rejected`.
- Column **B – Approver**
  - Filled automatically with the approver’s Google Workspace email when they change `Approval` → `Approved`.
- Columns **Q – Building Calendar Event ID**, **R – Website Calendar Event ID**
  - Store the event IDs created on the Building and Member/Public calendars.
  - Prevent duplicate event creation if the row is re‑edited.

All automation reads and writes from this sheet only.

---

#### `Config` Sheet — Approval + Calendar IDs

The `Config` sheet provides:

| Column | Header                | Purpose                                           |
|--------|-----------------------|---------------------------------------------------|
| A      | `Approval Statuses`   | List of valid statuses: `Pending`, `Approved`, `Rejected` (used for data validation on `Form Responses 1!A:A`) |
| B      | `Member Calendar ID`  | Calendar ID used for **Members and Friends** events |
| C      | `Public Calendar ID`  | Calendar ID used for **General Public** events     |
| D      | `Building Calendar ID`| Calendar ID used for **building reservations**     |

Typical layout:

- `A1`: `Approval Statuses`
- `A2:A4`: `Pending`, `Approved`, `Rejected`
- `B1`: `Member Calendar ID`
- `B2`: actual Member calendar ID
- `C1`: `Public Calendar ID`
- `C2`: actual Public calendar ID
- `D1`: `Building Calendar ID`
- `D2`: actual Building Reservation calendar ID

The Apps Script reads `Config!B2:D2` to know where to create calendar events.

---

### 3. `All Upcoming` View Tab

`All Upcoming` is a read‑only view used by humans and by the website or other consumers. It:

- Shows **only events where `Approval = Approved`**
- Shows **only events whose end datetime is in the future**
- Combines the date + time into start/end datetimes
- Is sorted by upcoming start time
- Is constrained to a reasonable maximum number of rows (e.g., 200)

#### `All Upcoming` Formula

On the `All Upcoming` sheet, put the following into cell `A1`:

```gs
=ARRAYFORMULA({
  {"Approver","Event Name","Description","Start","End",
   "Target Audience","Building Spaces","Advertise Where",
   "Setup/Teardown","Needs Graphic?","Graphic",
   "Form Timestamp","Building Event ID","Website Event ID"};
  ARRAY_CONSTRAIN(
    SORT(
      FILTER(
        {
          'Form Responses 1'!B2:B,
          'Form Responses 1'!E2:E,
          'Form Responses 1'!F2:F,
          'Form Responses 1'!G2:G + 'Form Responses 1'!H2:H,
          'Form Responses 1'!G2:G + 'Form Responses 1'!I2:I,
          'Form Responses 1'!K2:K,
          'Form Responses 1'!J2:J,
          'Form Responses 1'!L2:L,
          'Form Responses 1'!M2:M,
          'Form Responses 1'!O2:O,
          'Form Responses 1'!P2:P,
          'Form Responses 1'!C2:C,
          'Form Responses 1'!Q2:Q,
          'Form Responses 1'!R2:R
        },
        'Form Responses 1'!A2:A="Approved",
        ('Form Responses 1'!G2:G + 'Form Responses 1'!I2:I) >= NOW()
      ),
      4, TRUE
    ),
    200, 14
  )
})
```

You can create additional filtered views (e.g., Members‑only, Public‑only, building‑only) by copying this pattern and adding extra `FILTER` conditions on the relevant columns.

#### Example: `Members Only Upcoming` tab

To create a tab that shows only **approved Member events** (target audience = `Members and Friends`), create a new sheet named `Members Only Upcoming` and put this into cell `A1`:

```gs
=ARRAYFORMULA({
  {"Approver","Event Name","Description","Start","End",
   "Target Audience","Building Spaces","Advertise Where",
   "Setup/Teardown","Needs Graphic?","Graphic",
   "Form Timestamp","Building Event ID","Website Event ID"};
  ARRAY_CONSTRAIN(
    SORT(
      FILTER(
        {
          'Form Responses 1'!B2:B,
          'Form Responses 1'!E2:E,
          'Form Responses 1'!F2:F,
          'Form Responses 1'!G2:G + 'Form Responses 1'!H2:H,
          'Form Responses 1'!G2:G + 'Form Responses 1'!I2:I,
          'Form Responses 1'!K2:K,
          'Form Responses 1'!J2:J,
          'Form Responses 1'!L2:L,
          'Form Responses 1'!M2:M,
          'Form Responses 1'!O2:O,
          'Form Responses 1'!P2:P,
          'Form Responses 1'!C2:C,
          'Form Responses 1'!Q2:Q,
          'Form Responses 1'!R2:R
        },
        'Form Responses 1'!A2:A="Approved",
        ('Form Responses 1'!G2:G + 'Form Responses 1'!I2:I) >= NOW(),
        'Form Responses 1'!K2:K="Members and Friends"
      ),
      4, TRUE
    ),
    200, 14
  )
})
```

This keeps the **same column order** as `All Upcoming` (so the website embed can reuse the same configuration) but adds an extra filter on the **Target Audience** column (`K`) so that only “Members and Friends” events appear.

#### Example: `Friday Flash` promotion tab

To create a tab that shows only **approved events that requested promotion in the “Friday Flash” channel** (based on **If this event is not Private, where should we advertise it?** in column `L`), create a new sheet named `Friday Flash Upcoming` and put this into cell `A1`:

```gs
=ARRAYFORMULA({
  {"Approver","Event Name","Description","Start","End",
   "Target Audience","Building Spaces","Advertise Where",
   "Setup/Teardown","Needs Graphic?","Graphic",
   "Form Timestamp","Building Event ID","Website Event ID"};
  ARRAY_CONSTRAIN(
    SORT(
      FILTER(
        {
          'Form Responses 1'!B2:B,
          'Form Responses 1'!E2:E,
          'Form Responses 1'!F2:F,
          'Form Responses 1'!G2:G + 'Form Responses 1'!H2:H,
          'Form Responses 1'!G2:G + 'Form Responses 1'!I2:I,
          'Form Responses 1'!K2:K,
          'Form Responses 1'!J2:J,
          'Form Responses 1'!L2:L,
          'Form Responses 1'!M2:M,
          'Form Responses 1'!O2:O,
          'Form Responses 1'!P2:P,
          'Form Responses 1'!C2:C,
          'Form Responses 1'!Q2:Q,
          'Form Responses 1'!R2:R
        },
        'Form Responses 1'!A2:A="Approved",
        ('Form Responses 1'!G2:G + 'Form Responses 1'!I2:I) >= NOW(),
        REGEXMATCH('Form Responses 1'!L2:L,"Friday Flash")
      ),
      4, TRUE
    ),
    200, 14
  )
})
```

This uses `REGEXMATCH` on the **Advertise Where** column (`L`) to include any row where “Friday Flash” appears (even if other channels are selected in the same cell), while keeping the same output layout as `All Upcoming`.

These view‑style tabs (especially `All Upcoming`) are also the **source of truth for the website event feeds**, which are rendered client‑side using the `events.js` snippet documented below.

---

## Apps Script Logic

All automation lives in a single Apps Script project attached to this spreadsheet, stored in **`Code.gs`** in this repo.

The project has:

- `combineDateAndTime(date, time)` – helper for combining separate date and time cells
- `normalize_(value)` – helper for case‑insensitive comparisons
- `onFormSubmit(e)` – initializes metadata columns for new submissions
- `onApprovalEdit(e)` – reacts to Approval → Approved, populates Approver, and creates calendar events

### Trigger 1 — `onFormSubmit`

**Event:** Installable trigger, “On form submit” on the `Form Responses 1` sheet.

Behavior:

- Sets `Approval` (`A`) to `Pending` for the new row if blank
- Clears `Approver` (`B`) if any stale value is present
- Clears both `Building Calendar Event ID` (`Q`) and `Website Calendar Event ID` (`R`) so re‑submissions never reuse stale IDs

### Trigger 2 — `onApprovalEdit`

**Event:** Installable trigger, “On edit” on the spreadsheet.

Behavior when `Approval` changes to `Approved` on a data row:

1. **Approver tracking**
   - Detects the editor’s Google account and writes their email to `Approver` (`B`).
2. **Base event timing**
   - Combines `Event Date` (`G`) with `Event Start Time` (`H`) and `Event End Time` (`I`) into start/end Date objects.
3. **Building reservation logic**
   - Checks **What part(s) of the building do you want to use?** (`J`):
     - If this cell is **blank**, **no building reservation** event is created.
     - If it is **non‑blank**, a Building calendar event is created (unless `Building Calendar Event ID` (`Q`) is already set).
   - Applies **setup/teardown padding** based on:
     - **If you are requesting building space, how much setup and teardown time do you need before/after your event?** (`M`)
     - Options:
       - `None` → no padding
       - `30 Minutes` → 30 minutes before and after
       - `1 Hour` → 1 hour before and after
       - `2 Hours` → 2 hours before and after
   - Example: Event 2–4 pm with `1 Hour` padding → Building reservation from **1–5 pm**.
   - Uses **What part(s) of the building…** (`J`) as the **Location** field on the Building calendar event.
4. **Member vs Public website calendar logic**
   - Reads **Who is the target audience for this event?** (`K`):
     - `Private Event` → **no website calendar event** is created.
     - `Members and Friends` → event is created on the **Member Calendar** (Config `Member Calendar ID`).
     - `General Public` → event is created on the **Public Calendar** (Config `Public Calendar ID`).
   - These events are **always** created for Members/Public when approved (no checkbox gating).
   - The **Location** field is also set from **What part(s) of the building…** (`J`) when available.
   - The created event’s ID is written to **Website Calendar Event ID** (`R`).
5. **Descriptions**
   - Building calendar event description typically includes:
     - Event description
     - Contact email
     - Target audience
     - Setup/teardown selection
     - Key holder / AV support request
     - Approved by (approver email)
   - Website calendar event description typically includes:
     - Event description
     - Contact email
     - Approved by (approver email)
6. **Idempotency**
   - If either `Building Calendar Event ID` (`Q`) or `Website Calendar Event ID` (`R`) already has a value, the script **does not** create another event for that calendar, but still updates the `Approver` field.

---

## Timezone Configuration (Pacific Time)

For event times to appear correctly on all calendars and on the website (e.g., a 6:00 pm selection shows as 6:00 pm everywhere), set all of the following to **Pacific Time (America/Los_Angeles)**:

- **Spreadsheet timezone**  
  - In the `Form Responses 1` spreadsheet: **File → Settings → Locale / Time zone**
- **Apps Script project timezone**  
  - In the Apps Script editor: **Project Settings → Script time zone**
- **Target Google Calendars timezone**  
  - In Google Calendar for:
    - Member Calendar
    - Public Calendar
    - Building Reservation Calendar

When all are aligned to Pacific time, the `combineDateAndTime` helper and the setup/teardown padding will behave as expected.

---

## Apps Script (Code.gs)

The repository includes:

- `Code.gs` – the full Apps Script implementation, containing:
  - `combineDateAndTime(date, time)`
  - `normalize_(value)`
  - Helpers for reading calendar IDs from `Config`
  - `onFormSubmit(e)`
  - `onApprovalEdit(e)`
  - Column mapping documentation inline

---

## Permissions Required

The Apps Script project will request:

- Spreadsheet read/write access
- Calendar creation access (for **all three** calendars)
- Access to the user’s email (to log Approver identity)

The first time triggers run, Google will prompt for authorization.

---

## Recommended Workflow for Volunteers

- **Step 1**: Open the `Form Responses 1` sheet.
- **Step 2**: Use a filter view: `Approval = Pending`.
- **Step 3**: Review details (date, time, spaces, target audience, setup/teardown, AV needs, graphics).
- **Step 4**: If approved, change `Approval` → `Approved`.
  - If building space was requested, a Building calendar reservation is created (with padding).
  - If target audience is `Members and Friends`, a Member calendar event is created.
  - If target audience is `General Public`, a Public calendar event is created.
  - If target audience is `Private Event`, **no** Member/Public website calendar event is created.
- **Step 5**: Check the `Building Calendar Event ID` and `Website Calendar Event ID` columns for confirmation.
- **Step 6**: Use the `All Upcoming` tab for a consolidated view of what’s coming up.

---

## Website Event Feeds (Iframe Embeds from Google Sheets)

Some parts of the UUCLV website are hosted in an environment where we can only run **client‑side JavaScript inside iframes**—we cannot upload standalone `.js` files or run server‑side code.

To support this, the project includes a **copy‑paste‑ready JavaScript snippet** (`events.js`) that:

- **Fetches a public Google Sheet tab** (typically `All Upcoming`, but any view tab with a compatible layout will work)
- **Parses rows via the Google Visualization (`gviz`) API**
- **Renders each row as a card** in a simple, readable event list
- **Allows expanding/collapsing full descriptions** by clicking the event title
- **Provides friendly error messages** when:
  - The browser is too old to run the code
  - The sheet is not publicly readable
  - The tab name or Sheet ID is wrong
  - The network request fails

All logic is completely self‑contained so it can be pasted directly into an iframe’s `<script>` tag without referencing any external files.

### How the iframe embed works

At a high level, each iframe:

1. Contains a **target container** where events will be rendered, e.g.:
   ```html
   <div id="uuclv-events-all-upcoming"></div>
   ```
2. Includes a `<script>` block that contains the contents of `events.js`, with a small **configuration section at the top** where you specify:
   - **Which sheet to read from** (`sheetId` or full `sheetUrl`)
   - **Which tab to use** (e.g., `All Upcoming`, `Members Upcoming`, `Public Upcoming`)
   - **Which DOM element to render into** (e.g., `#uuclv-events-all-upcoming`)
3. Optionally customizes some behavior (e.g., whether to inject default styles or rely on site CSS).

Because each iframe has its own isolated DOM and JavaScript environment, you can safely have **multiple iframes on the same page**, each pointing at:

- Different tabs of the **same** spreadsheet, or
- Completely different spreadsheets

…simply by copy‑pasting the snippet and adjusting the configuration per iframe.

### Requirements for the sheet/tab

For the default configuration (using `All Upcoming`):

- The tab should follow the header layout created by the `All Upcoming` formula above:
  - `Approver`, `Event Name`, `Description`, `Start`, `End`, `Target Audience`, `Building Spaces`, …
- `Start` and `End` columns should be formatted as **Date/Time** in the spreadsheet; the embed uses the **formatted text** from the sheet for display.
- The sheet (or at least the relevant tab) must be shared as:
  - **"Anyone with the link can view"**  
    (no sign‑in required, otherwise the iframe users will see an error message).

You can also create derivative tabs (e.g., “Public Only”, “Members Only”) that:

- Use the same column order as `All Upcoming`
- Add extra `FILTER` conditions on `Target Audience` or other fields

These derivative tabs will “just work” with the same embed snippet as long as the column order stays compatible.

### Using `events.js` in an iframe

1. **Make sure the sheet is public‑viewable**
   - In the Google Sheet: **Share → General access → Anyone with the link → Viewer**
2. **Create an HTML file for the iframe** (wherever your hosting system lets you edit raw HTML).
3. In that file, create a container and include the script:

   ```html
   <div id="uuclv-events-all-upcoming"></div>

   <script>
   // 1) Paste the contents of events.js here.
   // 2) Update the CONFIG values at the top of the script:
   //    - sheetId or sheetUrl
   //    - sheetTab (e.g. "All Upcoming")
   //    - targetSelector (e.g. "#uuclv-events-all-upcoming")
   </script>
   ```

4. **Optional**: place this HTML file inside an `<iframe>` on any page:

   ```html
   <iframe
     src="/path/to/uuclv-events-all-upcoming.html"
     title="UUCLV Upcoming Events"
     style="border:0;width:100%;max-width:900px;"
   ></iframe>
   ```

Each iframe can have its own copy of the snippet with different configuration, allowing multiple custom feeds on a single page.

### Browser support and error handling

The embed code is intentionally simple but assumes a **reasonably modern browser** with:

- `fetch`
- `Promise`
- Basic DOM APIs (`querySelector`, `classList`, `addEventListener`)

If these are missing, or if anything goes wrong while fetching or parsing the sheet, users will see a **clear, human‑readable message** in place of the event list, such as:

- “Your web browser is too old to display this events list.”
- “We couldn’t load the events right now. Please try again later or contact the office.”

These messages are designed for non‑technical users and do not expose raw error details.

### UX details: expanding/collapsing event descriptions

In the rendered list:

- Each event appears as a **card** showing:
  - Event title
  - “When” line (start → end)
  - Optional “Where” line (from `Building Spaces`)
- **The full description is hidden by default** to keep the list compact.
- Clicking the **event title** toggles the full description **open/closed** underneath.
  - Titles are rendered as accessible, keyboard‑focusable links with `aria-expanded` updated as they toggle.

This behavior is fully implemented inside `events.js`; no extra code is needed in the host page.

---

## Repository Structure

```text
/
├── README.md          # This document
├── Code.gs            # Apps Script implementation for the combined UUCLV events workflow
└── events.js          # Copy‑pasteable client‑side embed for website event feeds (iframe‑safe)
```


