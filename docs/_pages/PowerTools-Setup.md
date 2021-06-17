---
permalink: /PowerTools-Setup
title: "PowerTools Setup"
excerpt: "How to Setup the PowerTools"
---

## Screencast

<p style="text-align: center;"><iframe width="560" height="315" src="https://www.youtube.com/embed/cPPvPdH9IGk" frameborder="0" allow="accelerometer; autoplay; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe></p>

## Deployment and Requirements

### Standalone executables

The PowerTools are provided as compiled, standalone, fully portable executables.

You don't need to have AutoHotkey installed on your PC. And they don't require any Admin rights to run.

When running these executables, you might get some Windows warning that it is unsecure. Normally you can check off these warnings to not be displayed again, either directly from the warning message or see for example [here](https://www.windowscentral.com/how-fix-app-has-been-blocked-your-protection-windows-10).

### Run from the AutoHotkey sources

Alternatively, you could run the tools from the AutoHotkey source: simply clone the [repository]((https://github.com/tdalon/ahk/)).

But I don't ensure that the current version of the repository runs properly, because I am human and might have forgotten to push some files. <a href="https://tdalon.github.io/ahk/contact/"><i class="fa fa-address-card" aria-hidden="true"></i></a>[Contact me](Contact) if you want to run from the AHK source and stumble upon some troubles.

Most of the users I know don't bother about the AutoHotkey source. Providing the compiled version allows me _not_ to check 100% consistency between my local version and the repository and quick deployment of a version. If the requests are numerous/ the demand high, I might change my way of deploying the sources.

## Download All

<a href="https://downgit.github.io/#/home?url=https://github.com/tdalon/ahk/tree/master/PowerTools" target="_blank"><button class="btn"><i class="fa fa-download"></i> Download</button></a>

The recommended way to install the PowerTools is to download the whole [PowerTools Directory](https://github.com/tdalon/ahk/raw/master/PowerTools/) [here](https://downgit.github.io/#/home?url=https://github.com/tdalon/ahk/tree/master/PowerTools).
This includes the subdirectory with images used for ImageSearch related features like the [Teams Meeting Reactions Shortcuts](Teams-Meeting-Reactions).

## One-by-one Setup

N.B.: For features based on ImageSearch (like [Teams Shortcuts](Teams-Shortcuts) [Meeting Live reactions](Teams-Meeting-Reactions), the *img* subdirectory is also required. Therefore I recommend to always [download all](#download-all) and run what you want only.

Still you can download for each PowerTool the .exe file by going here: [repo root](https://github.com/tdalon/ahk/)->[PowerTools](https://github.com/tdalon/ahk/tree/master/PowerTools) and clicking on the .exe file you want and then click the Download button. If you do so, copy it in a separate fresh directory because some auxiliary files will be downloaded or created on the same level (e.g. PowerBundler, git_updater, ini file...).

## First run

From each PowerTool you can access functions like Open help, Add to Startup, Check for update etc. from the tool System tray icon menu.

### Set system tray icons to be displayed permanently

The first time you run a PowerTool, the tool system tray icon might be hidden. You can configure it via the Windows Taskbar Settings to be always visible. See for example instructions [here](https://www.ghacks.net/2015/03/11/manage-and-display-system-tray-icons-in-windows-10/).

### Launch on Startup

You can configure each PowerTool to be launched on your PC Startup via its Settings.

### Disable notifications at Startup

By default, a notification is displayed at start to make you aware of the system tray icon functionality.

You can disable these in the Settings. This is a setting for all PowerTools, not tool specific.

## Manage via PowerTools Bundler

If you want to manage multiple PowerTools together, the best way is to use the [PowerTools Bundler](PowerTools-Bundler).
You can download the standalone version by clicking directly [here](https://github.com/tdalon/ahk/raw/master/PowerTools/PowerToolsBundler.exe).

## Update procedure

You can update each .exe from each tool systrem tray icon menu: _Check for Update_.
And also from the [PowerTools Bundler](PowerTools-Bundler) via the _Actions-> Check for Update/ Download_.

![PowerTools Bundler Check for Update/Download](/ahk/assets/images/powertools_bundler_checkforupdate.png)

If some images (used in ImageSearch-based features) were updated, you need to update the whole directory manually by re-downloading all and overwriting your previous installation directory. Be careful not to overwrite your images if you have customized some of them.
