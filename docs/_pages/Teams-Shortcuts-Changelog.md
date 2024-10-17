---
permalink: /Teams-Shortcuts-Changelog
title: "Teams Shortcuts Changelog"
excerpt: "Release notes for Teams Shortcuts PowerTool."
---

[Teams Shortcuts](Teams-Shortcuts) Changelog

Some changes listed here also impact the [Teamsy Launcher](Teamsy-Launcher) features.

* 2024-10-17 [Teams Favorites for Chat Group or Meeting from Message Link](https://tdalon.blogspot.com/2024/10/teams-chat-link.html) 
* 2024-05-01
  * Fix: Global Hotkey for Meeting share can not be set
* 2024-03-12
  * Fix Teams Mute: Property Name not changed. Moved to FullDescription
* 2024-01-30
  * Teams Meeting Share: fix. Move FocusAssist to the end because of interaction with Sharing Bar hide. Add pause after maximize.
* 2023-11-22
  * Improved Teams_MeetingLeave: prompt for End or Leave
* 2023-11-03
  * [Set meeting background](https://tdalon.blogspot.com/2023/11/teams-set-background.html)
* 2023-10-27
  * Updated SysTray Icon actions
  * Teams_Mute
    * display mic Status in systray icon
    * Teams_Mute restore window
  * Teams_Video: restore window
* 2023-10-26
  * [PushToTalk by MButton Long Click on System Tray icon](https://tdalon.blogspot.com/2023/10/teams-shortcuts-ptt.html)
  * RButton on SysTray Icon activate Teams main window if not in a meeting
  * double Right click open meeting window
* 2023-10-26
  * [Improved Language Dependency Handling](https://tdalon.blogspot.com/2023/10/teams-powertools-lang.html)
  * [Opening Teams Links without Left-over browser window](https://tdalon.blogspot.com/2023/10/teams-open-link-leftover.html)
* 2023-10-24
  * Removed Teams Hotkeys for Unread and Saved (depreciated in New Client + accessible from Launcher 'u' and 's' keywords)
* 2023-10-20
  * Support for Teams New Client/ Refactoring
    * Restart, Clear Cache
    * Teams_GetMainWindow
    * Teams_GetMeetingWindow
    * Teams_GetLang
    * Tested: Teams_Mute
  * [Teams Backgrounds Import and Migrate](https://tdalon.blogspot.com/2023/10/teams-import-background.html) for Teams New Client
* 2023-10-20
  * Open Favorites .url Internet Shortcut without Browser Left-over window
* 2023-09-19
  *  [Focus Assist Integration](https://tdalon.blogspot.com/2023/09/ahk-focus-assist.html)
* 2023-08-10
  * DE Language specific implementation for mute and meeting reactions
* 2023-08-08
  * New [Teams_GetLang](https://tdalon.blogspot.com/2023/08/teams-get-client-lang.html) function
* 2023-06-19
  * fix: icon for favorites shortcuts
* 2023-06-16
  * OpenLink [ByPass SafeLink check](https://tdalon.blogspot.com/2023/01/teams-bypass-safelink.html): moved to NWS PowerTool OpenLink feature
* 2023-06-13
  * Update/Refactoring Teams_GetMeetingWindow. Client change new split Leave button
  * Fix RaiseHand if Meeting window is minimized
* 2023-04-17
  * Favorites for Chat links open in Web browser instead of app client (client does not work)
* 2023-02-20
  * [Quick Join Teams Meeting](https://tdalon.blogspot.com/2023/02/teams-quick-join-meeting.html) using Outlook Calendar Integration
* 2023-02-09
  * Update Personalize Mentions to work for Lastname, Firstname (company) mention format
* 2023-01-30
  - [Outlook to Teams Group Chat](https://tdalon.blogspot.com/2023/01/outlook-to-teams-group-chat.html)
* 2023-01-26
  - Fix Get Meeting Window if AutomationId not available. Fallback ByName "Calling controls"
  - Add ["Add Members"](https://tdalon.blogspot.com/2020/03/emails-to-team.html) Menu (feature moved from PeopleConnector). New Launcher Keyword "t2xl"
* 2023-01-25
  - Bug fix meeting reactions if Reactions menus is opened
* 2023-01-23
  - [Selection to Chat](https://tdalon.blogspot.com/2023/01/teams-open-chat.html)
* 2023-01-11
  - New feature: [Share to Teams](https://tdalon.blogspot.com/2023/01/share-to-teams.html)
  - fix: remove wrong Meeting Toggle menu entries: Launcher, Clear Cache etc.
  - added: Global hotkeys are displayed in menu label (Settings->Global Hotkeys and Meeting->Toggle)
  - New Feature: [Shift+Click on a Link Bypass SafeLinks](https://tdalon.blogspot.com/2023/01/teams-bypass-safelink.html)
  - Remove Param MeetingWinUseFindText
* 2023-01-09
  - Fix: Add Email to Teams Favorites. If no selection check for clipboard (Use Case: Copy email from Teams visit card)
  - [New](https://tdalon.blogspot.com/2023/01/teams-favorites.html) Launcher command "f+" Add to favorites
  - New: 'f' Launcher command to [open favorites](https://tdalon.blogspot.com/2023/01/teams-favorites.html). (No need for external application launcher to use anymore.)
  - Changed: Launcher Keyword 'f'. (Previously opened find, now favorites)
* 2022-10-18
  - Fix Teams Get Meeting Window (Share| Call in Progress) with meeting participants
  - Teams Conversation Reactions and Actions
* 2022-09-06
  - Updated Shortcut for Toggle Fullscreen based on F11 hotkey. Keyword 'fs' from Launcher. New function: Teams_MeetingToggleFullscreen.
* 2022-07-18
  - Conversations Reactions: Like, Heart, Laugh, Surprised, Sad, Angry
  - Conversation: Copy Link using UIA
  - Copy Link: change hotkey to Alt+C instead of Win+C because of conflict with Cortana
* 2022-07-13
  - Leave: Works also with Call is in progress Share Window. No need to activate meeting window.
  - Mute: Works also with "Call is in progress" Share Window. No need to activate meeting window.
  - Video: Works also with "Call is in progress" Share Window. No need to activate meeting window.
  - Video: Option to switch Video on or off (not only toggle)
  - Share: Works without FindText/ with UI Automation. No need to activate Meeting Window. Unshare works also with Call in progress minimized Share window.
  - Share: also reactivates Meeting window to secondary screen.
  - Share: share audio by default.
* 2022-07-08
  - Update [Teams Meeting Reactions based on UIAutomation](https://tdalon.blogspot.com/2022/07/ahk-teams-meeting-reactions-uia.html) instead of FindText
* 2022-07-01
  - Teams_Mute/PushToTalk based on UIAutomation instead of FindText
* 2022-03-14
  - Added keyword 'lo'|'lobby' to Admit Lobby
* 2022-03-11
	- Remove [hard dependency](https://tdalon.blogspot.com/2022/03/ahk-conditional-include.html) to Connections Lib
* 2022-03-08
  - Support PushToTalk Ctrl+Space global hotkey
  - Added Open Background Settings
  - Fix: Quick Share (new Background option)
* 2022-02-24
  - Added global Hotkey to Activate Meeting Window
* 2022-02-22
  - fix: if FindText does not work, restore previous meeting window
* 2022-01-12
	- Teams Share: fix if main monitor index is not 1. restore / move meeting window to secondary monitor
  - Teams Share: update to new share design (2 Tab)
* 2021-11-29
  - [Workaround](https://tdalon.blogspot.com/2021/11/teams-shortcuts-personalize-mention-fix.html) for issue with Personalize Mentions in flat standalone chats
* 2021-11-11
  - Change Hotkey Win+1->Alt+1 (Conflict with Windows native hotkey)
* 2021-11-04
  - [Fix hotkey conflict handling](https://tdalon.blogspot.com/2021/11/ahk-regex-pipe.html): escape | in regex
* 2021-10-28
  - [Fix hotkey conflict handling](https://tdalon.blogspot.com/2021/10/ahk-hotkey-conflict.html)
* 2021-10-20
  - Leave meeting hotkey changed from Ctrl+Shift+B to Ctrl+Shift+H by Microsoft
* 2021-10-18
  - [Extend Teams Shortcuts functionality to Browser (multiple tenant handling)](https://tdalon.blogspot.com/2021/10/teams-shortcuts-browser-support.html)
* 2021-09-20
  - Bug fix: Open background folder if directory does not exist
* 2021-05-12
	- [Taskbar Flashing Hotkey](https://tdalon.blogspot.com/2021/05/teams-pitfall-flashing-taskbar.html)
* 2021-04-19
	- Fix Teams Share: select desktop only if sharing (not for unshare). Extend Pause when sharing for selecting desktop. Added Param TeamsShareDelay (Default 1.5s)
	- Teams_Mute: access via main window instead of meeting window for toggle mute
* 2021-04-12
	- Add [Mute on/ off shortcut](https://tdalon.blogspot.com/2021/04/teams-shortcuts-mute-on-off.html). Improve PTT.
* 2021-04-07
	- Hotkeys Settings are saved to ini File and can be loaded from it. (Bundler: Load Config from Ini)
* 2021-04-01
	- Added TeamsClickDelay Parameter
	- Get Meeting Window: remove user prompt for window selection using FindText
* 2021-03-31
	- Meeting Reactions now based on [FindText](https://tdalon.blogspot.com/2021/03/ahk-findtext-teams-shortcuts.html) instead of ImageSearch
	- New Meeting Actions: [FullScreen](https://tdalon.blogspot.com/2021/03/ahk-findtext-teams-shortcuts.html) (Keyword: 'fs'), TogetherMode (Keyword: 'tm')
	- New Meeting Actions: Share (Keyword: 'sh+'), Unshare (Keyword: 'sh-')
	- New Meeting Actions: Background Settings (Keyword: 'bg')
* 2021-03-29
	- [Handling of Parameters in Settings](https://tdalon.github.io/ahk/Teams-Parameters)
* 2021-03-22
	- Added TeamsCommandDelay Parameter: increase def value to 800ms
  - Bug fix: SmartReply from menu + clipboard sendinput sync
  - GetMeetingWindow issue when sharing: Exclude "Screen sharing toolbar" window title
  - Share buggy if multiple monitor/ move meeting window to second screen will unshare -> do not move window to second screen
	- Teams_Link2Text: handling of spaces in channel link. replaces %2520 by spaces
	- Add Channel to Favorites: Default name based on Team name and channel name
* 2021-03-19
	- Added notification at startup
* 2021-03-12
  - [Meeting Reactions](https://tdalon.blogspot.com/2021/03/teams-meeting-reactions-shortcuts.html): Like, Laugh, Applause, Heart
* 2021-03-08
  - [Add contact to favorites](https://tdalon.blogspot.com/2021/03/teams-people-favorites.html): new menu entry
* 2021-03-05
  - Added Menu Settings-> Open Favorites Directory and Set Favorites Directory
* 2021-03-01
  - Add [Teamsy Launcher](Teamsy-Launcher)
* 2021-02-24
  - [PushToTalk Feature](https://tdalon.blogspot.com/2021/02/teams-push-to-talk.html)
* 2021-02-23
  - fix Teams_GetMeetingWindow and Teams_GetMainWindow if VirtuaWin is used. (Main Window might be hidden, on another virtual desktop)
* 2021-02-22
	- add [global hotkey](https://tdalon.github.io/ahk/teams-global-hotkeys) for [Raise Your Hand](https://tdalon.blogspot.com/2021/02/teams-raise-hand.html)
	- add [global hotkey](https://tdalon.github.io/ahk/teams-global-hotkeys) for [Mute App](https://tdalon.blogspot.com/2021/02/teams-mute-app.html)
* 2021-02-15
	- fix: [Get Meeting Window](https://tdalon.blogspot.com/2020/10/get-teams-window-ahk.html#getmeetingwindow) with , in title. Improved RegEx; only if Name, Firstname () format
* 2021-02-09
	- Bug fix. First start prompt for Connections Url
	- Remove Meeting->VLC Menus
	- Add global hotkey setting for Toggle Video
* 2021-02-08
	- [Fix](https://tdalon.blogspot.com/2021/02/ahk-tray-no-active-window.html). SysTray Actions will restore previous active window on Toggle mute and Video
* [2021-01-27](https://tdalon.blogspot.com/2021/02/teams-shortcuts-new-features-202101.html)
	- Added Setting to configure a Hotkey to toggle Mute
	- Added SysTray Mouse Click: Right Mouse Click -> Toggle video. Middle Mouse Click -> Toggle Mute
* 2021-01-26
	- Added menu in systray [Clear cache](https://tdalon.blogspot.com/2021/01/teams-clear-cache.html)
* 2021-01-18
	- New conversation: adjusted timing for expanding the compose box
* 2021-01-15
	- Extended keyword cal -> calendar
* 2021-01-14
	- Added Custom Background Setting for Library. Now it can be [set in ini file](https://tdalon.blogspot.com/2021/01/teams-custom-backgrounds.html#openlib).
* 2021-01-12
	- Change hotkey for new conversation from Win+N to Alt+N because of collision with OneNote hotkey
* 2020-12-08
	- Fixed: Personalize Mentions with Firstname with -
* 2020-12-04
  - Fixed Teams_GetMainWindow: collision of previous WinId. Added check for Window name as static variable
* 2020-11-24
  - SysTray Menu:
  	- add Tweet for support.
  	- Contact via Teams only for Conti Config.
  - [Smart Reply](https://tdalon.blogspot.com/2020/11/teams-shortcuts-smart-reply.html#getme): Add setting to store personal email/name so that Smart Reply also works without connection to Active Directory.
* 2020-11-23
	- [Smart Reply](https://tdalon.blogspot.com/2020/11/teams-shortcuts-smart-reply.html) improved version with Mention.
* 2020-11-20
	- [Send/Personalize Mentions](https://tdalon.blogspot.com/2020/11/teams-shortcuts-personalize-mentions.html): handling of case mentioned named can not be autocompleted (e.g. user not a member)
* 2020-11-17
  - Teams_GetMainWindow: improved version using Acc (no need to minimize all windows)
* 2020-11-16
	- [Personalize Mention](https://tdalon.blogspot.com/2020/11/teams-shortcuts-personalize-mentions.html) improved: Unified hotkeys for mention personalization (detect if () used by selecting/copying to clipboard)
* 2020-10-30
	- Fix new conversation with change of hotkey
	- Added: Paste Mentions (Extract Emails from Clipboard and convert Email to Name to be typed as mention in Teams
* 2020-10-28
  - Remove hotkey (Win+M; bad choice) for create a meeting. Rather launch from Teamsy.
  - Add SubMenu Meeting -> Cursor Highlighter
* 2020-09-21
    * [Smart Reply](https://tdalon.blogspot.com/2020/11/teams-shortcuts-smart-reply.html): Update to change of hotkey by Microsoft: new hotkey Alt+Shift+r instead of r for reply
* 2020-09-14
    * VLC Integration for Play
* 2020-07-27
    * Export Team List (used for NWS PowerTool IntelliPaste)
* 2020-07-24
    * Export Team Members List to Excel incl. Emails
* 2020-07-20
    * New conversation: will also focus on the subject line (Shift+Tab)
    * Changed win+r hotkey to alt+r (smart reply)
    * New meeting (alt+m)
* 2020-05-26
    * Add Win+P for /pop pop-out chat shortcut
    * Bulk add user will keep the PowerShell Window open and visible
* 2020-05-07
    * Custom Backgrounds: add open GUIDEs Backgrounds folder
* 2020-04-23
    * Add function "Open Second Instance" and "Open Web App"
* 2020-02-26
    * IntelliPaste Get Team Name by PowerShell
    * IntelliPaste: Paste Conversation: choose between Team|Channel|Message link
* 2020-02-21: add to favorites: prefill link text with Channel Name
* 2020-02-20: added [help](https://connectionsroot/blogs/tdalon/entry/teams_shortcuts_ahk) to Teams Favorites
* 2020-02-13: Improved Smart Reply
