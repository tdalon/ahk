---
permalink: /Muty
title: "Muty"
excerpt: "Muty is an application launcher plugin to mute your microphone or any running application."
---

Muty is an application launcher plugin to mute your microphone or any running application.

This is the brother commander utility of [Mute PowerTool](Mute-PowerTool)

Documentation is in progress.

## Requirements

It requires the SetVol utility by rlatour. See [https://rlatour.com/setvol/](https://rlatour.com/setvol/) which is freeware and portable.

For muting an application it requires [NirSoft SoundVolumeView](https://www.nirsoft.net/utils/sound_volume_view.html).

## Setup

Run from source Muty.ahk or download the standalone compiled version from PowerTools/Muty.exe

Integration in application launcher e.g. Executor:

## Usage

Keywords: ProcessName or Process.exe : mute application
Keywords on|off|0 (off)|1 (on)|2 (toggle): Mute default mic device.
?: open help

## Source code

Main function is Muty.ahk
It is mainly a command line wrapper for the Mute function in Mute.ahk
