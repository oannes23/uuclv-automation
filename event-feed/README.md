# Event Feed

This system uses a publicly viewable Google Sheet as a sort of database back end which is then read by JavaScript embdedded on a web page which is used to render various upcoming events feeds. The Google Sheet is multi-tabbed, with the Event Data tab being the entry point, each row containing columns of all relevant data needed for the Event Feed. 

A few of the tabs have some formula plumbing, and both those and the Event Data are hidden by default. The remainder of the tabs are different views of the Event Feed, each tab representing a useful perspective to look at the Events from. For example, there are both Public and Member Events tabs, and the public section of the website only pulls from the Public Events tab, but the Member section might have a page pulling from the Member Events tab.

There is an Apps Script tied to the Sheet that monitors the email inbox of the automation@uuclv.org address. If it detects any emails from Breeze's event submission form, it adds them as new rows to the Event Data tab.

## Directory Contents

**[`Code.gs`](code.gs)** - The Google App Script that is attached to the Sheet and used to automatically monitor the Breeze emails for the legacy event submission form.

**[formulas.md](formulas.md)** - Markdown file containing the important Excel formulas needed to create the Event Promotion sheet

## End to End Process






