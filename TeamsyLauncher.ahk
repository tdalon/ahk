; See documentation https://tdalon.github.io/ahk/Teamsy-Launcher

LastCompiled = 20231026154346


#Include <Teamsy>


#SingleInstance force ; for running from editor

HKid = Launcher
RegRead, HK, HKEY_CURRENT_USER\Software\PowerTools, TeamsHotkey%HKid%
If (HK != "") {
		MenuLabel = &Run Launcher`t(%HK%)
	} Else
		MenuLabel = &Run Launcher

Menu, Tray, Add, %MenuLabel%, Teams_Launcher
SubMenuSettings := PowerTools_MenuTray()
Menu,Tray,Default,%MenuLabel%

If (HK != "") {
		Teams_HotkeyActivate(HKid,HK, False)
		MenuLabel = %HKid% Hotkey`t(%HK%)
	} Else
		MenuLabel = %HKid% Hotkey

Menu, SubMenuSettings, Add, %MenuLabel%, Teams_HotkeySet

; Tooltip
If !a_iscompiled 
    FileGetTime, LastMod , %A_ScriptFullPath%
Else 
    LastMod := LastCompiled
FormatTime LastMod, %LastMod% D1 R

sTooltip = Teamsy Launcher %LastMod%`nRight-Click on icon to access other functionalities.
Menu, Tray, Tip, %sTooltip%


; Reset Main WinId at startup because of some possible hwnd collision
PowerTools_RegWrite("TeamsMainWinId","")
return


Teams_Launcher:
Teamsy("-g")
return