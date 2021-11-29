---
permalink: /Teams-Shortcuts-Changelog
title: "Teams Shortcuts Changelog"
excerpt: "Release notes for Teams Shortcuts PowerTool."
---

[Teams Shortcuts](Teams-Shortcuts) Changelog

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
