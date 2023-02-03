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
 * a *Teams Commander* or [*Teams Controller*](https://www.msxfaq.de/teams/client/teams_controller.htm) i.e. it can also be used from the [command line](https://tdalon.blogspot.com/2022/03/ahk-command-line.html) to send commands to Microsoft Teams; for example, as bridge between a Stream Deck and Microsoft Teams.

You can also run Teamsy from the [Teamsy Launcher](Teamsy-Launcher) or the [Teams Shortcuts](Teams-Shortcuts) PowerTool.

Its main advantage is that it works even if your Microsoft Teams window client isn't active (contrary to built-in hotkeys) e.g. both for main window and meeting window actions.
Moreover the keyword-based syntax shall be much easier to remember than hotkeys combination.

## [Main Blog Post](https://tdalon.blogspot.com/2020/07/teamsy.html)

<p style="text-align: center;"><iframe width="560" height="315" src="https://www.youtube.com/embed/zLFWKFfLHnU" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe></p>

See [all blog posts tagged with #teamsy](https://tdalon.blogspot.com/search/label/teamsy).

## Prerequisites

Teamsy is a available as standalone .exe application. [Download link](https://github.com/tdalon/ahk/raw/main/PowerTools/Teamsy.exe)

You can run Teamsy from the [command line](https://tdalon.blogspot.com/2022/03/ahk-command-line.html) or from an Application Launcher.
I [recommend](https://tdalon.blogspot.com/2020/08/executor-my-preferred-app-launcher.html) [Executor](http://executor.dk/) as application launcher.

### Setup with Executor

[Executor](http://executor.dk/)

Add a new Keyword in Executor as shown in screenshot below:

<div style="text-align:center"><img src="/ahk/assets/images/Executor_Teamsy_Setup.png" alt="Teamsy Executor Setup"></div>

As Parameter enter: "$P$" (between quotes)


### Setup with Launchy/ LaunchyQt

[Launchy](http://launchy.net/)

[LaunchyQt](https://github.com/samsonwang/LaunchyQt)

Unfortunately the Launchy Runner Plugin seems not to support passing an argument with a space in it. (It will break the argument at the first space into a second argument.)
A workaround is implemented in Teamsy.ahk so that it will aggregate the arguments together.

Thanks to this, the working configuration in Launchy looks like:

<div style="text-align:center"><img src="/ahk/assets/images/Launchy_Teamsy_Setup.png" alt="Teamsy Launchy Setup"></div>

You will need to add a custom command to the Runner plugin.
To run it in launcher, type the command keyword t in the screenshot, then TAB followed by the Teamsy command.

### Setup with a Stream Deck

Use the [compiled exe](https://tdalon.blogspot.com/2022/03/ahk-command-line.html) and pass optional arguments as command keywords.

#### [Download Teamsy.exe](https://github.com/tdalon/ahk/raw/main/PowerTools/Teamsy.exe) ###

See detailed post [here](https://tdalon.blogspot.com/2023/02/teamsy-with-deckboard.html).

### Alternative without application launcher

If you want to run Teamsy without using an application launcher, have a look at [Teams Launcher](Teamsy-Launcher) or [Teams Shortcuts](Teams-Shortcuts)'s integrated Launcher.

## List of supported commands/ Keywords

This is implemented in the main associated library [Lib/Teamsy.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Teamsy.ahk).

Here you can find the syntax for keywords and command. (I hope the code is clear enough.)
<script src="http://gist-it.appspot.com/https://github.com/tdalon/ahk/raw/main/Lib/Teamsy.ahk"></script>

### Examples

The list below might not be exhaustive. Look at the code for the full up to date implementation.

Keywords  |  Action
--|--
mu   |  Toggle mute
vi  |  Toggel video
sh  |  Toggle share
lo  |  Love (Meeting Reaction)  
li  |  Like (Meeting Reaction)    
ap  |  Applause (Meeting Reaction)
lol  |  Laugh (Meeting Reaction)
sa  |  Sad (Meeting Reaction)  
an  |  Angry (Meeting Reaction)  
rh  |  Raise your Hand (Meeting Reaction)
le  |  Leave Meeting

<hr>

## Download ##

You can download the portable standalone compiled .exe here: [Download Teamsy.exe](https://github.com/tdalon/ahk/raw/main/PowerTools/Teamsy.exe)

The source code is available here: https://github.com/tdalon/ahk

## Feature Highlights

### Status Change

<div style="text-align:center"><img src="/ahk/assets/images/Teamsy_StatusChange.gif" alt="Teamsy Status Change"></div>

### [Meeting Live Reactions](Teams-Meeting-Reactions)

### [Share To Teams](https://tdalon.blogspot.com/2023/01/share-to-teams.html)

Keyword: 's2t'

## YouTube Playlist

<div align="center"><iframe width="560" height="315" src="https://www.youtube.com/embed/zLFWKFfLHnU" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe><br><a href="https://www.youtube.com/watch?v=zLFWKFfLHnU">Direct Link to YouTube video</a></div>

## Source Code

The main ahk file is [Teamsy.ahk](https://github.com/tdalon/ahk/blob/main/Teamsy.ahk)

The main associated library is [Lib/Teamsy.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Teamsy.ahk)

The main associated library is [Lib/Teams.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Teams.ahk)

## [Changelog](Teamsy-Changelog)

## [News](https://twitter.com/search?q=%23Teamsy%20%23MicrosoftTeams)
