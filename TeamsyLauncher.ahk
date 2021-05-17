; See documentation TODO

LastCompiled = 20210416063519


#Include <Teamsy>


#SingleInstance force ; for running from editor

Menu, Tray, Add, &Run Launcher, Teams_Launcher
SubMenuSettings := PowerTools_MenuTray()
Menu,Tray,Default,&Run Launcher


HKid = Launcher

; Hotkeys: Activate, Meeting Action Menus and Settings Menus

Menu, SubMenuSettings, Add, %HKid% Hotkey, Teams_HotkeySet
RegRead, HK, HKEY_CURRENT_USER\Software\PowerTools, Teams%HKid%Hotkey
If (HK != "") {
    Teams_HotkeyActivate(HKid,HK, False)
}
label = Teams_%HKid%Cb
If IsLabel(label)
    Menu, SubMenuMeeting, Add, Toggle %HKid%, %label% ; Requires Cb Label for not loosing active window
Else
    Menu, SubMenuMeeting, Add, Toggle %HKid%, Teams_%HKid% ; Requires Cb Label for not loosing active window


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