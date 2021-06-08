---
permalink: /Teamsy-Launcher
title: "Teamsy Launcher"
redirect_from:
  - /teamsy-launcher
excerpt: "Microsoft Teams Application Launcher"
---

## About

Teamsy Launcher is a standalone Launcher for Microsoft Teams.
It has the same functionality as [Teamsy](Teamsy) but does not require to use an external Application Launcher like [Launchy](http://launchy.net/) or [Executor](http://executor.dk/).

Note: the Launcher feature is also included in [Teams Shortcuts](Teams-Shortcuts). So, if you already use [Teams Shortcuts](Teams-Shortcuts), you don't need Teamsy Launcher.
But if you don't want to use the other features included in [Teams Shortcuts](Teams-Shortcuts), you can use Teamsy Launcher.


## [Main Blog Post](https://tdalon.blogspot.com/2020/07/teamsy.html)

<p style="text-align: center;"><iframe width="560" height="315" src="https://www.youtube.com/embed/zLFWKFfLHnU" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe></p>

See [all blog posts tagged with #teamsy](https://tdalon.blogspot.com/search/label/teamsy).

## Prerequisites

Teamsy Launcher is a available as small standalone .exe application. [Download link](https://github.com/tdalon/ahk/raw/master/PowerTools/TeamsyLauncher.exe)

You can also run it from its AutoHotkey source [TeamsyLauncher.ahk](https://github.com/tdalon/ahk/blob/master/TeamsyLauncher.ahk) provided you take all the dependencies/ downlad the [full repository](https://github.com/tdalon/ahk).

Contrary to [Teamsy](Teamsy) it does not require any third-party launcher application but include the Launcher feature.

## How to use

Simply run TeamsyLauncher (ahk or exe).
When it is running the first time, you shall configure a hotkey to start the launcher.
Right-Click on the System Tray Icon and go to Settings->Launcher Hotkey.
(I like to use Ctrl+Space)

Running the Hotkey will open an input dialog box where you can type the commands to be sent to Microsoft Teams Client.
You can also run the Launcher by double-clicking on the System Tray icon.

## List of supported commands/ Keywords

The list of supported commands/ keywords are the same as for [Teamsy](Teamsy).

This is implemented in the main associated library [Lib/Teamsy.ahk](https://github.com/tdalon/ahk/blob/master/Lib/Teamsy.ahk)
Here you can find the syntax for keywords and command. (I hope the code is clear enough.)
<script src="http://gist-it.appspot.com/https://github.com/tdalon/ahk/raw/master/Lib/Teamsy.ahk"></script>

## Feature Highlights

See [Teamsy](Teamsy)

## YouTube Playlist

<div align="center"><iframe width="560" height="315" src="https://www.youtube.com/embed/zLFWKFfLHnU" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe><br><a href="https://www.youtube.com/watch?v=zLFWKFfLHnU">Direct Link to YouTube video</a></div>

## Source Code

The main ahk file is [TeamsyLauncher.ahk](https://github.com/tdalon/ahk/blob/master/TeamsyLauncher.ahk)
The keywords and commands are implemented in [Lib/Teamsy.ahk](https://github.com/tdalon/ahk/blob/master/Lib/Teamsy.ahk)

The main associated library is [Lib/Teams.ahk](https://github.com/tdalon/ahk/blob/master/Lib/Teams.ahk)

## [Changelog](Teamsy-Changelog)

## [News](https://twitter.com/search?q=%23Teamsy%20%23MicrosoftTeams)
