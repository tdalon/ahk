---
permalink: /Teamsy
title: "Teamsy"
redirect_from:
  - /teamsy
excerpt: "Microsoft Teams Application Launcher Plugin and Commander"
---

## About

Teamsy is a
 * Microsoft Teams *Plugin for an Application Launcher* like [Launchy](http://launchy.net/) or [Executor](http://executor.dk/) and
 * a *Teams Commander* i.e. it can also be used from the command line to send commands to Microsoft Teams; for example, as bridge between a StreamDeck and Microsoft Teams.

You can also run Teamsy from the [Teamsy Launcher](Teamsy-Launcher) or the [Teams Shortcuts](Teams-Shortcuts) PowerTool.

Its main advantage is that it works even if your Microsoft Teams window client isn't active (contrary to built-in hotkeys) e.g. both for main window and meeting window actions.
Moreover the keyword-based syntax shall be much easier to remember that hotkeys combination.

## [Main Blog Post](https://tdalon.blogspot.com/2020/07/teamsy.html)

<p style="text-align: center;"><iframe width="560" height="315" src="https://www.youtube.com/embed/zLFWKFfLHnU" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe></p>

See [all blog posts tagged with #teamsy](https://tdalon.blogspot.com/search/label/teamsy).

## Prerequisites

Teamsy is a available as standalone .exe application. [Download link](https://github.com/tdalon/ahk/raw/main/PowerTools/Teamsy.exe)
You can also run it from its AutoHotkey source [Teamsy.ahk](https://github.com/tdalon/ahk/blob/main/Teamsy.ahk) provided you take all the dependencies/ downlad the [full repository](https://github.com/tdalon/ahk).

You can run Teamsy from the command line.
I personally use Teamsy from an Application launcher e.g. [Launchy](http://launchy.net/) or [Executor](http://executor.dk/).
I [recommend](https://tdalon.blogspot.com/2020/08/executor-my-preferred-app-launcher.html) [Executor](http://executor.dk/).

### Alternative without application launcher

If you want to run Teamsy without using an application launcher, have a look at [Teams Launcher](Teamsy-Launcher) or [Teams Shortcuts](Teams-Shortcuts)'s integrated Launcher.

## List of supported commands/ Keywords

This is implemented in the main associated library [Lib/Teamsy.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Teamsy.ahk)
Here you can find the syntax for keywords and command. (I hope the code is clear enough.)
<script src="http://gist-it.appspot.com/https://github.com/tdalon/ahk/raw/main/Lib/Teamsy.ahk"></script>

<hr>

## Feature Highlights

### Status Change

<div style="text-align:center"><img src="/ahk/assets/images/Teamsy_StatusChange.gif" alt="Teamsy Status Change"></div>

### [Meeting Live Reactions](Teams-Meeting-Reactions)

## YouTube Playlist

<div align="center"><iframe width="560" height="315" src="https://www.youtube.com/embed/zLFWKFfLHnU" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe><br><a href="https://www.youtube.com/watch?v=zLFWKFfLHnU">Direct Link to YouTube video</a></div>

## Source Code

The main ahk file is [Teamsy.ahk](https://github.com/tdalon/ahk/blob/main/Teamsy.ahk)

The main associated library is [Lib/Teamsy.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Teamsy.ahk)

The main associated library is [Lib/Teams.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Teams.ahk)

## [Changelog](Teamsy-Changelog)

## [News](https://twitter.com/search?q=%23Teamsy%20%23MicrosoftTeams)
