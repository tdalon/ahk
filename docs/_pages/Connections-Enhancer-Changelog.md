---
permalink: /connections-enhancer-changelog
title: "Connections Enhancer Changelog"
excerpt: "Release notes for HCL Connections Enhancer PowerTool."
---

[Connections Enhancer](Connections-Enhancer) Changelog

* 2021-04-22
	- Added systray menu entry to open main menu.
* 2021-04-01
	- Change Connections_CloseCodeView based on FindText instead of ImageSearch
	- Bug fix: refresh Toc for Wiki always add new Toc. Do not update (change of html macro code for wiki toc with TinyMCE editor+Connections 6.5) 
* 2021-03-12
  - [Connections 6.5 Update](https://tdalon.blogspot.com/2021/04/connections-enhancer-update-for-6p5.html)
    - Fix for Regression: Ctrl+Shift+U hotkey does not open code view if editor view is HTML Source.
    - Now support TinyMCE editor (hotkey available)
* 2021-02-08
	- Insert TOC: cancel on TOC depth input exits without inserting TOC
* 2020-12-11
	- Fix: Ctrl+S opens Editor Help (missing return)
* 2020-12-07
  - Fix: Check if Edit mode based on Mouse cursor for forums to display proper Read or Edit Menu.
  - Refactor Connections_IsWinActive/WinEdit
* 2020-11-26
  - New [SendMentions Feature](https://tdalon.blogspot.com/2020/11/connections-enhancer-send-mentions.html)
* 2020-11-24
    - Bug Fix RegEx ReConnectionsRootUrl
	- Embed ConNext icon for Conti Config
* 2020-10-30
	- Fix Connections_Email2Key. used in PeopleConnector->Open Connections Network
* 2020-10-01
  * Create Table of mentions with pictures/ Extract mentions: remove (uid) in mentions name
* 2020-09-29
    * Refactoring GetRootUrl. Add in Settings. Ask if not defined.
    * Move to public
* 2020-09-11
    * Add Create New to Main menu
* 2020-07-24
    * Clean Table: removed: td formatted as align center.
    * Clean all: will also Clean Tables
* 2020-07-22
    * Refresh Toc (Win+f5) will insert Wiki Toc on top of the page if no Toc is defined
* 2020-07-14
    * fix VS Studio Code cleaner: remove indentation from pasted code
* 2020-06-18
    * fix: format images
* 2020-05-18
    * Table of Contents
        * also insert TOC macro for wiki page
        * Refresh TOC macro wiki page will also clean-up Headings ids
        * Add Delete Toc menu
        * Insert Toc will ask for location: top,cursor,bottom
    * Switch to Source Code with cursor location kept

* 2020-03-31
    * Personalize Mentions: run as clean/ batch mode. update for new display names
    * Format Images
        * FIX: Format Image for Blog
        * Only format width if width isn't already specified
        * Ignore picture from Win10 emojis
* 2020-03-25
    * FIX: Clean Table Width to 95%
* 2020-03-06
    * FIX: Boxes with Brake icon
* 2020-03-05
    * FIX: Expand Mentions with Profile pictures: no doubling of picture if picture uses link by email (like in generated Table) instead of userid
* 2020-03-03
    * FIX: Toc for SubPages
* 2020-02-20
    * FIX: Event to Calendar: Wrong time zone
* 2020-02-19
    * Improved Event to Calendar/Meeting/Email: decode subject e.g. &->&amp;
* 2020-02-18
    * Improved Event to Calendar/Meeting: add link to event in body
    * New: Event to Email
    * Add context-sensitive help for some features
