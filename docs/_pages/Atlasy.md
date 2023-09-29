---
permalink: /Atlasy
title: "Atlasy"
redirect_from:
  - /atlasy
excerpt: "Atlassian Launcher and Commander"
---

## About

This is currently a Work-In-Progress.
If attention is attracted and resonance is high (interest shown by personal feedback or comment), this project will be better documented and shared.

Atlasy is a tool that enables your working within the Atlassian-based tools like Jira, Confluence, R4J etc.
It adds some UX related features and also extended capability.
You can run its features from an integrated Launcher with natural keywords and commands or via a menu or hotkeys.

Short Feature highlights include:
 * bulk linking between Jira issues
 * quick search (Jira, Confluence)
 * quick navigation (e.g. R4J tree)
 * quick set an epic
 * ...


## [Main Blog Post](https://tdalon.blogspot.com/2020/07/teamsy.html)

<p style="text-align: center;"><iframe width="560" height="315" src="https://www.youtube.com/embed/zLFWKFfLHnU" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe></p>

See [all blog posts tagged with #atlasy](https://tdalon.blogspot.com/search/label/atlasy).

## Prerequisites

Atlasy is available as AutoHotkey scripts but also as standalone compiled .exe application. [Download link](https://github.com/tdalon/ahk/raw/main/PowerTools/Atlasy.exe)

## List of supported commands/ Keywords

This is implemented in the main associated library [Lib/Teamsy.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Teamsy.ahk).

In the source code you can find the full up to date syntax for keywords and command. (I hope the code is clear enough.)

### Examples

The list below might not be exhaustive. Look at the code for the full up to date implementation.

Keywords  |  Action
--|--
r   |  r4j
c  |  Confluence
j   | Jira  

<hr>

## Download ##

You can download the portable standalone compiled .exe here: [Download Atlasy.exe](https://github.com/tdalon/ahk/raw/main/PowerTools/Atlasy.exe)

The source code is available here: https://github.com/tdalon/ahk

## Feature Highlights

### Jira Integration

### R4J Integration

#### Launcher

Launcher Keyword: 'r'

##### Quick Open Project in R4J Tree

r ProjectKey
r IssueKey

r from Jira issue detailed view-> Open issue in R4J Tree

#### Hotkey to switch from detailed issue view and tree view



## YouTube Playlist

<iframe width="560" height="315" src="https://www.youtube.com/embed/videoseries?list=PLUSZfg60tAwLDw9tBZXLYH3OXlP3f4Awn" title="YouTube video player" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe>

## Source Code

The main ahk file is [Teamsy.ahk](https://github.com/tdalon/ahk/blob/main/Atlasy.ahk)

Main associated libraries are:
  * [Lib/Jira.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Jira.ahk)
  * [Lib/Confluence.ahk](https://github.com/tdalon/ahk/blob/main/Lib/Confluence.ahk)
  * [Lib/R4J.ahk](https://github.com/tdalon/ahk/blob/main/Lib/R4J.ahk)

## [Changelog](Atlasy-Changelog)

## [News](https://twitter.com/search?q=%23Atlasy)
