---
permalink: /Teams-Parameters
title: "Teams Shortcuts Parameters"
excerpt: "Parameters that can be tuned for Teams Shortcuts functionality"
---

## Short Description

Some parameters might need to be fine-tuned in order for some functionality to work properly.

## How to change Teams Shortcuts parameters

You can access Teams Shortcuts parameters from the Settings menu of the System tray Icon menu, under Settings->Parameters.

Alternatively you can also edit the parameters in the PowerTools.ini file under the [Teams] section and then load the Ini configuration from the PowerTools Bundler.

## Screencast

<p style="text-align: center;"><iframe width="560" height="315" src="https://www.youtube.com/embed/b_cBXbRZXmc" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe></p>


## Parameters Description

### TeamsCommandDelay (in ms)

Time delay in ms for which the tool will pause after entering a command in the Teams client command bar. Default is 800ms.

### TeamsMentionDelay (in ms)

Time delay in ms for which the tool will pause for a mention to be resolved. Default is 1300ms.

It is used in the features [Send Mentions](https://tdalon.github.io/ahk/Teams-Shortcuts#send-mentions) and [Smart Reply](https://tdalon.github.io/ahk/Teams-Shortcuts#smart-reply-altr).

### TeamsClickDelay (in ms)

Time delay between two clicks for the next UI element to load/ be visible. Default is 500ms.
It is used for combined click actions. Example for Meeting reactions, first you click on the reactions button, then on the specific reaction element. Or for the Meeting Actions: you click first on the 3 dots, then on the specific menu element.

### TeamsMeetingWinUseFindText (0|1)

Default is 1. See explanation in this [blog post](https://tdalon.blogspot.com/2021/04/ahk-get-teams-meeting-window.html)
