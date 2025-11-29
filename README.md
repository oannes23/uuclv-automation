# uuclv-automation

This repo exists to hold code snippets and documentation relating to the UUCLV Google Workspace. This includes complex spreadsheet formulae, App Script snippets, and instructional markdown files for use with AI coding agents to help create, maintain, and document the former items.

Each directory will contain information related to a particular project that I am automating to make it easier to maintain.

## Projects Overview

### Event Feed

This system uses a publicly viewable Google Sheet as a sort of database back end which is then read by JavaScript embdedded on a web page which is used to render various upcoming events feeds. The Google Sheet is multi-tabbed, with the Event Data tab being the entry point, each row containing columns of all relevant data needed for the Event Feed. A few of the tabs have some formula plumbing, and both those and the Event Data are hidden by default. The remainder of the tabs are different views of the Event Feed, each tab representing a useful perspective to look at the Events from. For example, there are both Public and Member Events tabs, and the public section of the website only pulls from the Public Events tab, but the Member section might have a page pulling from the Member Events tab.



