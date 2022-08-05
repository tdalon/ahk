---
permalink: /Teams-Meeting-Reactions
title: "Teams Meeting Reactions Shortcuts"
excerpt: "Teams Meeting Reactions Shortcuts."
---

## Short Description

Teams Shortcuts PowerTool allows to send Teams Meeting reactions from a Launcher or Hotkey.
It is based on [AutoHotkey UI Automation Library](https://tdalon.blogspot.com/2022/07/ahk-teams-uiautomation.html).

## Screencast

<p style="text-align: center;"><iframe width="560" height="315" src="https://www.youtube.com/embed/CuhXwQamuLE" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe></p>


## Related Blog posts

* [Teams Meeting Reactions](https://tdalon.blogspot.com/2022/07/ahk-teams-meeting-reactions-uia.html)
* [Teams Audio Effects](https://tdalon.blogspot.com/2021/12/teams-audio-effects.html)

## How to use

You can run the Meeting reactions using the Launcher with the keywords as defined in [Lib/Teamsy.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Teamsy.ahk).
Current keywords as of date of documentation (might not be consistent with current implementation):

Keyword  |  Reaction
--|--
li, like   |  Like
lol, la, laugh  |  Laugh
he, heart,lo  |  Heart, Love
ap, clap  |  Applause
su  | Surprised  

This make it so much easier to LOL in a meeting (click to enlarge):

<div style="text-align:center"><a href="https://tdalon.github.io/ahk/assets/images/Teams_Lol.gif"><img src="/ahk/assets/images/Teams_Lol.gif" alt="Teams Meeting Reaction LOL"></a></div>

## Live reactions with Audio

See [Teams Audio Effects](https://tdalon.blogspot.com/2021/12/teams-audio-effects.html)

You can combine built-in live reactions with additional audio effects.

<div style="text-align:center"><a href="https://tdalon.github.io/ahk/assets/images/Teams_MeetingLiveReactions_Audio.gif"><img src="/ahk/assets/images/Teams_Lol.gif" alt="Teams Meeting Reaction LOL"></a></div>

## Code

This feature is implemented in [/ahk/Lib/Teams.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Teams.ahk) -> Teams_MeetingReaction (function)

## Potential improvements

I could also add configurable global hotkeys for such actions but I find it much easier to run from the launcher via a natural command or keyword.
