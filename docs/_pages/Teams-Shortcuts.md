---
permalink: /Teams-Shortcuts
title: "Teams Shortcuts"
excerpt: "Tool empowering Microsoft Teams Windows Client Desktop application with hotkeys and improved functionality."
---

This tool extends the Microsoft Teams Windows Client with additional hotkeys and functionality.

## How to install

See separate page here [PowerTools Setup](PowerTools-Setup).

## YouTube Playlist

<div align="center"><iframe width="560" height="315" src="https://www.youtube.com/embed/videoseries?list=PLUSZfg60tAwLe8lIxZCpH38tP2jf4sv5m" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe><br><a href="https://www.youtube.com/playlist?list=PLUSZfg60tAwLe8lIxZCpH38tP2jf4sv5m">Direct Link to YouTube playlist</a></div>

## [Related Blog posts](https://tdalon.blogspot.com/search/label/teams-shortcuts)

## [Changelog](Teams-Shortcuts-Changelog)

## [News](https://twitter.com/search?q=%23TeamsShortcuts%20%23MicrosoftTeams)

## Contact

<a class="twitter-hashtag-button"
  href="https://twitter.com/intent/tweet?button_hashtag=#TeamsShortcuts&text=@tdalon #MicrosoftTeams #TeamsShortcuts"
  data-size="large">
Tweet</a>

You can also [create an issue in GitHub](https://github.com/tdalon/ahk/issues).

## How to use

You can start the Teams Shortcuts main menu by pressing the **Windows+T** Hotkey when a Teams window is active.

![TeamsShortcuts Menu](/ahk/assets/images/TeamsShortcuts_MainMenu.png)

In the menu you can also see which hotkeys directly run the features e.g. Alt+R for Smart Reply.

You can access some settings and features via the Teams Shortcuts System Tray Icon Menu.

## Feature Highlights

### Integrated Teams launcher

You can make use of the [Teamsy](Teamsy) functionality from Teams Shortcuts. See this [blog post](https://tdalon.blogspot.com/2021/03/teamsy-launcher.html).

You can either launch the launcher from the Teams Shortcuts System Tray icon menu *Launcher*.

<div style="text-align:center"><img src="/ahk/assets/images/TeamsShortcuts_Launcher.png" alt="Teams Shortcuts Launcher Menu"></div>

Or assign a [global hotkey](#global-hotkeys) to it.

### [Global hotkeys](Teams-Global-Hotkeys)

### New Conversation (Alt+N)

It will open a new conversation with the expanded compose box and move back the cursor right on the subject line - within one hotkey.

<div style="text-align:center"><img src="/ahk/assets/images/TeamsShortcuts_NewConversation.gif" alt="Teams Shortcuts New Conversation"></div>

See [main blog post](https://tdalon.blogspot.com/2020/10/teamsy-new-conversation.html)

### Smart reply (Alt+R)

See [main blog post](https://tdalon.blogspot.com/2020/11/teams-shortcuts-smart-reply.html)

<div style="text-align:center"><img src="/ahk/assets/images/TeamsShortcuts_SmartReply.gif" alt="Teams Shortcuts Smart Reply"></div>

### Send Mentions

See [main blog post](https://tdalon.blogspot.com/2020/11/teams-shortcuts-send-mentions.html)

![TeamsShortcuts Send Mentions Gif](/ahk/assets/images/TeamsShortcuts_SendMentions.gif)

### [Meeting Live Reactions](Teams-Meeting-Reactions)

### [Push To Talk](https://tdalon.blogspot.com/2021/02/teams-push-to-talk.html)

### [Mute on/ off](https://tdalon.blogspot.com/2021/04/teams-shortcuts-mute-on-off.html)

## Troubleshooting

If some features does not work properly, check-out if tuning [these parameters](Teams-Parameters) can solve the issue.

## Source code

Main code is contained in [TeamsShortcuts.ahk](https://github.com/tdalon/ahk/blob/master/TeamsShortcuts.ahk)
Main Library used is [Lib/Teams.ahk](https://github.com/tdalon/ahk/blob/master/Lib/Teams.ahk)
