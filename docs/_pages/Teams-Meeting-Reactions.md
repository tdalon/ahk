---
permalink: /Teams-Meeting-Reactions
title: "Teams Meeting Reactions Shortcuts"
excerpt: "How to setup Teams Meeting Reactions Shortcuts."
---

## Short Description

Teams Shortcuts PowerTool allows to send Teams Meeting reactions from a Launcher or Hotkey.
It is based on AutoHotkey [FindText](https://www.autohotkey.com/boards/viewtopic.php?f=6&t=17834) functionality.

## Screencast

<p style="text-align: center;"><iframe width="560" height="315" src="https://www.youtube.com/embed/sPy07IzEGu4" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe></p>


## [Related Blog post](https://tdalon.blogspot.com/2021/03/teams-meeting-reactions-shortcuts.html)

## How to use

You can run the Meeting reactions using the Launcher with the keywords as defined in [Lib/Teamsy.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Teamsy.ahk).
Current keywords as of date of documentation (might not be consistent with current implementation):

Keyword  |  Reaction
--|--
li, like   |  Like
lol, la, laugh  |  Laugh
he, heart  |  Heart, Love
ap, clap  |  Applause

This make it so much easier to LOL in a meeting (click to enlarge):

<div style="text-align:center"><a href="https://tdalon.github.io/ahk/assets/images/Teams_Lol.gif"><img src="/ahk/assets/images/Teams_Lol.gif" alt="Teams Meeting Reaction LOL"></a></div>

## How to setup/ troubleshooting

The feature relies on finding some UI element in the Teams Client using the FindText function as explained in [this post](https://tdalon.blogspot.com/2021/03/ahk-findtext-teams-shortcuts.html)

If the FindText does not work, you can overwrite the Text string to be searched with your working version in the .ini file. In the section [Teams] you will find the parameters starting with **TeamsFindText**<Action>


## Code

This feature is implemented in [/ahk/Lib/Teams.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Teams.ahk) -> Teams_MeetingReaction (function)

## Potential improvements

I could also add configurable global hotkeys for such actions but I find it much easier to run from the launcher via a natural command or keyword. (I have problems remembering hotkeys)
