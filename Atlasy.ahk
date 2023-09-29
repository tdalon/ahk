; Homepage: https://tdalon.github.io/ahk/Teamsy
LastCompiled = 20230811075038
#Include <PowerTools>
#Include <Atlasy>

#SingleInstance force ; for running from editor


If (A_Args.Length() > 0) {
    ; Loop, because of Launchy Runner Plugin not handling "..." properly as one input argument but splitting arguments at space
    Loop % A_Args.Length() { 
        Arg := Arg . " " . A_Args[A_Index]
    }
    Atlasy(Trim(Arg))
    return
}

; If (A_Args.Length() = 0)  
SubMenuSettings := PowerTools_MenuTray()
Menu, SubMenuSettings, Add, Notification at Startup, MenuCb_ToggleSettingNotificationAtStartup

RegRead, SettingNotificationAtStartup, HKEY_CURRENT_USER\Software\PowerTools, NotificationAtStartup
If (SettingNotificationAtStartup = "")
	SettingNotificationAtStartup := True ; Default value
If (SettingNotificationAtStartup) {
  Menu, SubMenuSettings, Check, Notification at Startup
} Else {
  Menu, SubMenuSettings, UnCheck, Notification at Startup
}
    
; Tooltip
If !a_iscompiled 
    FileGetTime, LastMod , %A_ScriptFullPath%
Else 
    LastMod := LastCompiled
FormatTime LastMod, %LastMod% D1 R

sTooltip = %A_ScriptName% %LastMod%`nRight-Click on icon to access help/support.
Menu, Tray, Tip, %sTooltip%



HotkeyIDList = Launcher 

; Hotkeys: Activate, Meeting Action Menus and Settings Menus
Loop, Parse, HotkeyIDList, `,
{
	HKid := A_LoopField
	HKid := StrReplace(HKid," ","")
	
	RegRead, HK, HKEY_CURRENT_USER\Software\PowerTools, AtlasyHotkey%HKid%
	; Activate Hotkey
	If (HK != "") {
		Atlasy_HotkeyActivate(HKid,HK, False)
		MenuLabel = %A_LoopField% `t(%HK%)
	} Else
		MenuLabel = %A_LoopField%

	Menu, SubMenuHotkeys, Add, %MenuLabel%, Atlasy_HotkeySet
}
Menu, SubMenuSettings, Add, Hotkeys, :SubMenuHotkeys
Menu, SubMenuMeeting, Add ; Separator


Menu,Tray,NoStandard
Menu, Tray, Add, Launcher, Atlasy_Launcher
Menu, Tray, Add
Menu, Tray,Standard

return


; ######################################################################
NotifyTrayClick_208:   ; Middle click (Button up)
Atlasy_Launcher()
Return 

NotifyTrayClick_202:   ; Left click (Button up)
Menu_Show(MenuGetHandle("Tray"), False, Menu_TrayParams()*)
Return

NotifyTrayClick_205:   ; Right click (Button up)

SendInput, !{Esc} ; for call from system tray - get active window

Return 

MenuCb_ToggleSettingNotificationAtStartup:
If (SettingNotificationAtStartup := !SettingNotificationAtStartup) {
  Menu, SubMenuSettings, Check, Notification at Startup
}
Else {
  Menu, SubMenuSettings, UnCheck, Notification at Startup
}
PowerTools_RegWrite("NotificationAtStartup",SettingNotificationAtStartup)
return

; ---------------------------- FUNCTIONS ------------------------------------------ 
NotifyTrayClick(P*) {              ;  v0.41 by SKAN on D39E/D39N @ tiny.cc/notifytrayclick
    Static Msg, Fun:="NotifyTrayClick", NM:=OnMessage(0x404,Func(Fun),-1),  Chk,T:=-250,Clk:=1
      If ( (NM := Format(Fun . "_{:03X}", Msg := P[2])) && P.Count()<4 )
         Return ( T := Max(-5000, 0-(P[1] ? Abs(P[1]) : 250)) )
      Critical
      If ( ( Msg<0x201 || Msg>0x209 ) || ( IsFunc(NM) || Islabel(NM) )=0 )
         Return
      Chk := (Fun . "_" . (Msg<=0x203 ? "203" : Msg<=0x206 ? "206" : Msg<=0x209 ? "209" : ""))
      SetTimer, %NM%,  %  (Msg==0x203        || Msg==0x206        || Msg==0x209)
        ? (-1, Clk:=2) : ( Clk=2 ? ("Off", Clk:=1) : ( IsFunc(Chk) || IsLabel(Chk) ? T : -1) )
    Return True
} ; eofun