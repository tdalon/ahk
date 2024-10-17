---
permalink: /Teamsy-Changelog
title: "Teamsy Changelog"
excerpt: "Release notes for Teamsy and Teamsy Launcher PowerTools."
---

[Teamsy](Teamsy) Changelog

See also [Teams Shortcuts Changelog](Teams-Shortcuts-Changelog) for other relevant changes.

* 2024-10-17 [Teams Favorites for Chat Group or Meeting from Message Link](https://tdalon.blogspot.com/2024/10/teams-chat-link.html) 
* 2024-03-14
  * Full Screen Shortcut update for New Teams (without F11)
* 2024-01-30
  * Teams Meeting Share: fix. Move FocusAssist to the end because of interaction with Sharing Bar hide. Add pause after maximize.
* 2023-11-24
  * New keyword "bgn" to set Background name
* 2023-11-22
  * Improved Teams_MeetingLeave: prompt for End or Leave
* 2023-11-20
  * Bug Fix Teams Meeting Share Hide Sharing Control Bar
* 2023-11-16
  * Quick Meeting Join: Dismiss Outlook Reminder. Fix for [plain meeting invitation issue](https://techcommunity.microsoft.com/t5/microsoft-teams/meeting-invitation-in-outlook-not-rendering-correctly/m-p/3268856)
* 2023-11-07
  * 'bgs' keyword for open background settings (Teams New)
* 2023-11-06
  * [Added Meeting End 'end' and 'le+' keywords](https://tdalon.blogspot.com/2023/11/teamsy-meeting-leave.html)
* 2023-11-04
  * Added mic(+-) and cam(+-) keyword 
* 2023-11-03
  * [Meeting record](https://tdalon.blogspot.com/2023/11/teams-shortcuts-record.html)
  * [Set meeting background](https://tdalon.blogspot.com/2023/11/teams-set-background.html)
* 2023-11-02
  * Teams Meeting Reactions will change Tray icon to display corresponding reaction for a few seconds
  * Teams Raise Hand supports on/off and displays shortly icon in Tray
* 2023-10-23
  * Improved [Meeting Share](https://tdalon.blogspot.com/2023/10/teams-quick-share-screen.html)
    * [Hide Sharing Control Bar](https://tdalon.blogspot.com/2023/10/teams-quick-share-screen.html). Added Keywords 'sb', 'sb+', 'sb-'
  * Change Launcher Help: takes input keyword
  * [Set status message](https://tdalon.blogspot.com/2023/10/teams-set-status-message.html)
  * [Switch tenant](https://tdalon.blogspot.com/2023/10/teams-switch-tenant.html)
* 2023-09-19
  *  [Focus Assist Integration](https://tdalon.blogspot.com/2023/09/ahk-focus-assist.html)
* 2023-06-13
  * Update/Refactoring Teams_GetMeetingWindow. Client change new split Leave button
  * Fix RaiseHand if Meeting window is minimized
* 2023-02-20
  * [Quick Join Teams Meeting](https://tdalon.blogspot.com/2023/02/teams-quick-join-meeting.html) using Outlook Calendar Integration
* 2023-01-30
  - [Outlook to Teams Group Chat](https://tdalon.blogspot.com/2023/01/outlook-to-teams-group-chat.html)
* 2023-01-26
  - Fix Get Meeting Window if AutomationId not available. Fallback FindByName "Calling controls"
  - New Keyword "e2m" Emails to Mentions
  - New keyword "cbg": Open Custom Backgrounds Folder
  - Add ["Add Members"](https://tdalon.blogspot.com/2020/03/emails-to-team.html) Menu (feature moved from PeopleConnector). New Launcher Keyword 't2xl'
* 2023-01-25
  - Bug fix meeting reactions if Reactions menus is opened
* 2023-01-23
  - New feature: [Open Chat](https://tdalon.blogspot.com/2023/01/teams-open-chat.html); Keyword 'oc','2c'
* 2023-01-20
  - Fix Command Bar: clear existing content.
  - CommandBar: @ + input
* 2023-01-11
  - New feature: [Share to Teams](https://tdalon.blogspot.com/2023/01/share-to-teams.html)
* 2023-01-09
  - Fix: Add Email to Teams Favorites. If no selection check for clipboard (Use Case: Copy email from Teams visit card)
  - [New](https://tdalon.blogspot.com/2023/01/teams-favorites.html) Launcher command "f+" Add to favorites
  - New: 'f' Launcher command to [open favorites](https://tdalon.blogspot.com/2023/01/teams-favorites.html). (No need for external application launcher to use anymore.)
  - Changed: Launcher Keyword 'f'. (Previously opened find, now favorites)
* 2022-09-06
  - Updated Shortcut for Toggle Fullscreen based on F11 hotkey. Keyword 'fs' from Launcher. New function: Teams_MeetingToggleFullscreen
* 2022-03-14
  - Added keyword 'lo','lobby' to Admit Lobby
* 2022-03-11
	- Remove [hard dependency](https://tdalon.blogspot.com/2022/03/ahk-conditional-include.html) to Connections Lib
* 2022-01-12
	- Teams Share: fix if main monitor index is not 1. restore / move meeting window to secondary monitor
	- Teams Share: update to new share design (2 Tab)
* 2021-04-12
	- Add [Mute on/ off shortcut](https://tdalon.blogspot.com/2021/04/teams-shortcuts-mute-on-off.html).
* 2021-03-31
	- Meeting Reactions now based on [FindText](https://tdalon.blogspot.com/2021/03/ahk-findtext-teams-shortcuts.html) instead of ImageSearch
	- New Meeting Actions: [FullScreen](https://tdalon.blogspot.com/2021/03/ahk-findtext-teams-shortcuts.html) (Keyword: 'fs'), TogetherMode (Keyword: 'tm')
	- New Meeting Actions: Share (Keyword: 'sh+'), Unshare (Keyword: 'sh-')
	- New Meeting Actions: Background Settings (Keyword: 'bg')
* 2021-03-22
	- Added TeamsCommandDelay Parameter: increase def value to 800ms
  - Bug fix: SmartReply from menu + clipboard sendinput sync
  - GetMeetingWindow issue when sharing: Exclude "Screen sharing toolbar" window title
  - Share buggy if multiple monitor/ move meeting window to second screen will unshare -> do not move window to second screen
	- Teams_Link2Text: handling of spaces in channel link. replaces %2520 by spaces
	- Add Channel to Favorites: Default name based on Team name and channel name
* 2021-03-12
  - [Meeting Live Reactions](https://tdalon.blogspot.com/2021/03/teams-meeting-reactions-shortcuts.html): Like, Laugh, Applause, Heart
* 2021-03-08
  - New Keyword for [add contact to favorites](https://tdalon.blogspot.com/2021/03/teams-people-favorites.html): 'p2f', 'e2f'
* 2021-03-05
  - New keywords for [favorites](https://tdalon.blogspot.com/2021/03/teams-shortcuts-favorites.html). 'of','fav': open favorites folder. '2f','2fav': add link to favorites
* 2021-03-04
  - Teamsy Launcher: added timeout and make inputbox modal/ always on top
  - New keywords 'wn' for whatsnew, 'nc' for new conversation
* 2021-03-01
  - [Teamsy Launcher](Teamsy-Launcher)
  - Added 'bg','background' keywords to open background folder
* 2021-02-24
  - fix: Restart, ClearCache, CleanRestart: replace WinActive by WinExist and Process, Exist
* 2021-02-23
  - fix Teams_GetMeetingWindow and Teams_GetMainWindow if VirtuaWin is used. (Main Window might be hidden, on another virtual desktop)
* 2021-02-22
	- add [Raise Your Hand](https://tdalon.blogspot.com/2021/02/teams-raise-hand.html)
	- add [Mute App](https://tdalon.blogspot.com/2021/02/teams-mute-app.html)
* 2021-02-15
	- fix: [Get Meeting Window](https://tdalon.blogspot.com/2020/10/get-teams-window-ahk.html#getmeetingwindow) with , in title. Improved RegEx; only if Name, Firstname () format
	- t m: activates meeting window. t n m: to create a new meeting
	- improved WinListBox: double click in list box supported
* 2021-02-08
	- Video Toggle: restore last window focus.
* 2021-01-29
	- Teamsy can be run at startup.
	- Added Icon with Help menu.
* 2021-01-27
	- Teams Mute: restore previously active window
* 2021-01-26
	- Added [Clear cache and Clean restart](https://tdalon.blogspot.com/2021/01/teams-clear-cache.html)
* 2021-01-22
	- Revert TeamsFavs: replace https to msteams to open directly in app (better multiple window handling.)
* 2021-01-21
	- Teams_GetMeetingWindow: Exclude Call in progress windows
	- Share: put back the meeting window if on the front and multiple monitors
* 2021-01-18
	- New conversation: adjusted timing for expanding the compose box
	- Added keyword for news
* 2021-01-11
	- Added keyword to open help/ documentation.
* 2021-01-08
	- Bug fix: Update SmartReply (due to change in UI Ctrl+A behavior.)
* 2021-01-05
	- Fix: Teams_GetMeetingWindow, Teams_GetMainWindow regexp escape pipe
* 2020-12-04
  - Fixed Teams_GetMainWindow: collision of previous WinId. Added check for Window name as static variable
* 2020-11-17
	- Teams_GetMainWindow: improved version using Acc (no need to minimize all windows)
* 2020-11-09
	- allow quick message via @ (previously / was prepended)
* 2020-10-28
  - Fix: Teams restart
* 2020-10-22
  - Improved new conversation: will work if content pane wasn't selected / from navigation list pane (drawback: flashing of search box if in conversation area)
* 2020-10-15
  - Mute hotkey does not require meeting window to be active. Main client window is enough.
* 2020-10-13
  - If no meeting window founds, clean exit
  - Add q:quit and r: restart
  - GetMeetingWindow with WinListBox; Excludes 1-1 chat window = containing ",". Preselect last Meeting window. Support on-hold meetings.
* 2020-10-07
  - Fix: cal keyword. (missing return)
  - Changed f->find (previously set status to free)
* 2020-10-06
  - Fix: New expanded conversation hotkey [broken](https://tdalon.blogspot.com/teamsy-new-conversation)
* 2020-09-30
    * Integrate to PowerTools Bundle. Add SysTray with link to help/ changelog. Add to Bundler.
* 2020-09-29
    * add 'cal' keyword to open calendar
