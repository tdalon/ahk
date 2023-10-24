---
title: "Microsoft Teams: Raise your Hand "
redirect_from:
  - /teams-raise-hand
excerpt: "How to raise your hand in a Microsoft Teams meeting, the powertool way, using Teamsy and Teams Shortcuts PowerTools."
categories:
  - Microsoft Teams
tags:
  - teamsy
  - teams-shortcuts
header:
  image: /assets/images/thumbnails/teams_raise_hand_og.png
  og_image: /assets/images/thumbnails/teams_raise_hand_og.png
  image_description: "Microsoft Teams: Raise your Hand with Teamsy and Teams Shortcuts PowerTools"
---

# Screencast

{% include video id="Ysytg8_lr74" provider="youtube" %}

# [Detailed Post](https://tdalon.blogspot.com/2021/02/teams-raise-hand.html)

# AutoHotkey code

```AutoHotkey
Teams_RaiseHand() {
; Toggle Raise Hand on/off
WinId := Teams_GetMeetingWindow()
If !WinId ; empty
    return
Tooltip("Teams Toggle Raise Hand...")
WinGet, curWinId, ID, A
WinActivate, ahk_id %WinId%
SendInput ^+k ; toggle video Ctl+Shift+k
WinActivate, ahk_id %curWinId%
} ; eofun
```
