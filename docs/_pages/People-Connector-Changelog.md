---
permalink: /People-Connector-Changelog
title: "People Connector Changelog"
excerpt: "Release notes for People Connector PowerTool."
---

[People Connector](People-Connector) Changelog

* 2021-09-30
  - Bug fix: Add to Teams Favorites, Call shortcut broken if . in folder name.
* 2021-05-05
	- Fix: copy email from Teams visit card / no selection -> take from clipboard
* 2021-04-27
	- Fix: Open LinkedIn Profile by Name from Connections/Teams Profile Name selection
* 2021-03-10
  - Fix Search by name e.g. LinkedIn from web page selection with clipboard in HTML format
* 2021-02-05
	- Fix issue from Connections visit card selection from Status update. Source Url contains profile link of the Status Update poster.-> Limit email parsing to HTML body.
* 2021-01-27
	- Added open Delve main page as default to D accelerator key, besides former open Delve Profile.
* 2021-01-26
	- Fix: Uid with domain returned OfficeId instead of Windows Id
* 2021-01-25
	- Fix: from Connections Name. Support profile links with ?key= (besides ?userid=)
* 2021-01-20
	- Fix: Open Connections profile from Teams on name selection. Convert HTML selection to text
* 2021-01-15
	- Fix: Compiled Date format in SysTray icon tooltip
	- Fix: Get Uid with Domain. Sometimes Uid was wrong/ empty.
* 2020-12-08
	- Extrat emails from Outlook Meeting Recipients
	- Increased time for mention autocompletion 1300ms
* 2020-12-02
  - Extract Domain\Uid
* 2020-11-24
  - SysTray Menu:
  	- add Tweet for support.
  	- Contact via Teams only for Conti Config.
* 2020-11-19
	- LinkedInSearch: extract name from email instead of by name
* 2020-11-18
	- Optimized one-time connection to ActiveDirectory (global variables/ AD_Init)
* 2020-11-11
	- Refactoring Outlook2Excel -> Emails2Excel. Works for any email selection (not only outlook). Based on People_Email2Name
* 2020-10-30
  - Bug fix: if selection is a single email (e.g. copy email from Teams visit card)
  - Add Connections Open Network View
* 2020-10-29
  - fixed. LinkedIn and Bing search on name selection e.g. connections. Revert to default plain text selection as input
  - fix: trim selection incl. new lines
* 2020-10-19
  - Bug fix: open from Teams visit card will catch wrong email @unq.gbl.spa(ces)
* 2020-10-07
    * If Excel application: do not getselection in Html format
* 2020-10-01
    * Remove (uid) from Connections mentions
    * First take selection in html format: Allows to run PeopleConnector from an email in Outlook body e.g. notification meeting forward.
* 2020-09-30
    * If not selection, take input from the current clipboard. Allow to run from Teams Meeting people visit card, where on only can copy the email to the clipboard (selection will mailto)
    * If no Connections used (ConnectionsRootUrl setting not set/ asked when running Connections Enhancer), Connections specific menus will not be displayed.
* 2020-09-11
    * Teams Chat will open directly in Teams App instead of going via the browser.
* 2020-09-09
    * LinkedIn Search by Name: remove numbers
* 2020-09-08
    * Create Teams Meeting
* 2020-08-05
    * MySuccess PeopleView to open orgChart
    * Remove Skype Chat from menu (Teams-Only)
* 2020-07-16
    * Bug fix: Firstname last letter truncated if no multiple lines were selected
* 2020-07-06
    * Get Name will only consider first line in case multiple lines are selected e.g. Name selection on forum entry in Connections
* 2020-07-03
    * Open Connections Profile now checks if input is an email. If not, search by name. (Menu Connections Search Profile by name removed). Works for example also from Teams on name selection.
    * Look-up from Name with Lastname, Firstname (e.g. Teams) will convert to Firstname Lastname
* 2020-07-02
    * Open LinkedIn or Bing search from email selection e.g. in Teams (convert email to name)
* 2020-05-26
    * Teams Add to Favorites - Chat: open in browser instead of app for multi-window better handling
* 2020-05-14
    * split Teams Chat in two functions: Chat will open only and Copy Link (will copy link to clipboard and ask to open)
    * Refactor accelerator keys: C will open a Teams Chat (multiple C accelerators where defined)
    * Added Bing search
* 2020-05-11
    * bug fix: extract emails with multiple dots e.g. co.uk
* 2020-02-26
    * Teams PowerShell Setting
* 2020-02-20
    * NEW: Export List of people in Outlook to Excel Table ([link](https://connectionsroot/blogs/tdalon/entry/people_connector_ol2xl))
* 2020-02-19
    * FIX copy from Connections news feed
    * add link to Release Notes / Change log
* 2020-02-13: Get userId Windows uid
* 2020-02-12: Emails to Team Members
