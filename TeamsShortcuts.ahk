; MS Teams Keyboard Shortcuts
; Author: Thierry Dalon
; See user documentation here: https://connectionsroot/blogs/tdalon/entry/teams_shortcuts_ahk
; Code Project Documentation is available on GitHub here: https://github.com/tdalon/ahk
; Source : https://github.com/tdalon/ahk/blob/main/TeamsShortcuts.ahk
;

LastCompiled = 20211014192529

#Include <Teams>
#Include <PowerTools>
#Include <WinClipAPI>
#Include <WinClip>

#SingleInstance force ; for running from editor
SetWorkingDir %A_ScriptDir%

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


Menu, SubMenuSettings, Add, Teams PowerShell, MenuCb_ToggleSettingTeamsPowerShell
RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If (TeamsPowerShell) 
  Menu, SubMenuSettings,Check, Teams PowerShell
Else 
  Menu, SubMenuSettings,UnCheck, Teams PowerShell

Menu, SubMenuSettings, Add, Teams Personalize Mentions, MenuCb_ToggleSettingTeamsMentionPersonalize
TeamsMentionPersonalize := PowerTools_RegRead("TeamsMentionPersonalize")

If (TeamsMentionPersonalize) 
  Menu,SubMenuSettings,Check, Teams Personalize Mentions
Else 
  Menu,SubMenuSettings,UnCheck, Teams Personalize Mentions

Menu, SubMenuSettings, Add, Update Personal Information, GetMe
Menu, SubMenuSettings, Add, Set Favorites Directory, Teams_FavsSetDir
Menu, SubMenuSettings, Add, Open Favorites Directory, Teams_FavsOpenDir


 ParamList = Mention Delay,Command Delay,Click Delay, Meeting Win Use FindText 

; Hotkeys: Activate, Meeting Action Menus and Settings Menus
Loop, Parse, ParamList, `,
	Menu, SubMenuParams, Add, %A_LoopField%, SetParam

Menu, SubMenuSettings, Add, Parameters, :SubMenuParams


HotkeyIDList = Launcher,Mute,Video,Mute App,Share,Raise Hand,Push To Talk,Clear Flash 

; Hotkeys: Activate, Meeting Action Menus and Settings Menus
Loop, Parse, HotkeyIDList, `,
{
	HKid := A_LoopField
	HKid := StrReplace(HKid," ","")
	Menu, SubMenuHotkeys, Add, %A_LoopField%, Teams_HotkeySet
	RegRead, HK, HKEY_CURRENT_USER\Software\PowerTools, TeamsHotkey%HKid%
	If (HK != "") {
		Teams_HotkeyActivate(HKid,HK, False)
	}

	label = Teams_%HKid%Cb
	If IsLabel(label)
		Menu, SubMenuMeeting, Add, Toggle %HKid%, %label% ; Requires Cb Label for not loosing active window
	Else
		Menu, SubMenuMeeting, Add, Toggle %HKid%, Teams_%HKid% ; Requires Cb Label for not loosing active window

}
Menu, SubMenuSettings, Add, Global Hotkeys, :SubMenuHotkeys
Menu, SubMenuMeeting, Add ; Separator


Menu,Tray,NoStandard
Menu, Tray, Add, Launcher, Teams_Launcher
Menu, Tray, Add, Add to Teams Favorites (Team or Channel), Link2TeamsFavs
Menu, Tray, Add, Add to Teams Favorites (Contacts), Email2TeamsFavs

Menu, SubMenuCustomBackgrounds, Add, Open Custom Backgrounds Folder, OpenCustomBackgrounds
Menu, SubMenuCustomBackgrounds, Add, Open Backgrounds Library, OpenCustomBackgroundsLibrary

Menu, Tray, Add, Custom Backgrounds, :SubMenuCustomBackgrounds
Menu, Tray, Add,Start Second Instance, Teams_OpenSecondInstance
Menu, Tray, Add,Clear Cache, Teams_ClearCache
Menu, Tray, Add,Open Web App, Teams_OpenWebApp
Menu, Tray, Add
Menu, Tray, Add,Export Team Members, Users2Excel
Menu, Tray, Add,Refresh Teams List, Teams_ExportTeams
Menu, Tray, Add

; SubMenu Meeting
Menu, SubMenuMeeting, Add, Open Teams Web Calendar, Teams_OpenWebCal
; Add Cursor Highlighter
Menu, SubMenuMeeting, Add, Cursor Highlighter, PowerTools_CursorHighlighter


; VLC Menu: not used. replaced by SplitCam
;Menu, SubMenuVLC, Add, Start VLC, VLCStart
;Menu, SubMenuVLC, Add, Set Play Mode, VLCPlayMode
;Menu, SubMenuVLC, Add, Reset, VLCReset
;Menu, SubMenuMeeting, Add, VLC, :SubMenuVLC

Menu, Tray, Add, Meeting, :SubMenuMeeting
Menu, Tray,Add
Menu, Tray,Standard

; Tooltip
If !a_iscompiled 
	FileGetTime, LastMod , %A_ScriptFullPath%
 Else 
	LastMod := LastCompiled
FormatTime LastMod, %LastMod% D1 R

sTooltip = Teams Shortcuts %LastMod%`nUse 'Win+T' to open main menu in Teams.`nClick on icon to access other functionalities.
Menu, Tray, Tip, %sTooltip%
If (SettingNotificationAtStartup)
	TrayTip Teams Shortcuts is running! , Click on icon to access Help`, Settings and functionality.


; -------------------------------------------------------------------------------------------------------------------
Menu, TeamsShortcutsMenu, add, Smart &Reply (Alt+R), SmartReply
Menu, TeamsShortcutsMenu, add, &Quote Conversation (Alt+Q), QuoteConversation
Menu, TeamsShortcutsMenu, add, &New Expanded Conversation (Alt+N), NewConversation
Menu, TeamsShortcutsMenu, add, Create E&mail with link to current conversation (Win+M), ShareByMail
Menu, TeamsShortcutsMenu, add, Send Mentions (Win+Q), SendMentions
Menu, TeamsShortcutsMenu, add, Personalize &Mention (Win+1), PersonalizeMention
Menu, TeamsShortcutsMenu, add, View &Unread (Win+U), ViewUnread
Menu, TeamsShortcutsMenu, add, View &Saved (Win+S), ViewSaved
Menu, TeamsShortcutsMenu, add, &Pop-out Chat (Win+P), Pop
Menu, TeamsShortcutsMenu, add, Add to &Favorites, Link2TeamsFavs
; -------------------------------------------------------------------------------------------------------------------

; Reset Main WinId at startup because of some possible hwnd collision
PowerTools_RegWrite("TeamsMainWinId","")

return



; ##########################   Hotkeys   ##########################################
;#IfWinActive,ahk_exe Teams.exe
#If Teams_IsWinActive()

#1:: ; <--- Personalize Mention
PersonalizeMention:
Teams_PersonalizeMention()
return

;--- Compose in Expand mode
; Alt + N
!n::  ; <--- New Expanded Conversation
NewConversation:
Teams_NewConversation()
return


; View Unread
; Win+U
#u:: ; <--- View Unread
ViewUnread:
If GetKeyState("Ctrl") {
	Run, "https://connectionsroot/blogs/tdalon/entry/teams_shortcuts_ahk" 
	return
}
Send ^e ; Select Search bar
SendInput /unread 
Sleep 500
Send {Enter}
return	

; View Saved
; Win+S
#s:: ; <--- View Saved
ViewSaved:
If GetKeyState("Ctrl") {
	Run, "https://connectionsroot/blogs/tdalon/entry/teams_shortcuts_ahk" 
	return
}
Send ^e ; Select Search bar
SendInput /saved 
Sleep 500
Send {Enter}
return

; Pop-out chat
; Win+P
#p:: ; <--- Pop-out chat
Pop:
If GetKeyState("Ctrl") {
	Run, "https://connectionsroot/blogs/tdalon/entry/teams_shortcuts_ahk" 
	return
}
Send ^e ; Select Search bar
SendInput /pop  
Sleep 500
Send {Enter}
return

; Alt+e
!e:: ; <--- Edit
Send {Enter}
Send {Tab}
Send {Enter}
Sleep 500
Send {Down 1}
Send {Enter}
return	

; Alt+.
!.:: ; <--- (...) Actions menu
Send {Enter}
Send {Tab}
Send {Enter}
return	

#t::
Menu, TeamsShortcutsMenu, Show
return

; Win+C
#c:: ; <--- Copy Link
TeamsCopyLink()
return	
; -------------------------------------------------------------------------------------------------------------------
; Alt+R - 
!r:: ; <--- Smart Reply with quotation and link to current thread
SmartReply:
Teams_SmartReply()
return

; -------------------------------------------------------------------------------------------------------------------
; Alt+Q
!q:: ; <--- Quote conversation
QuoteConversation:
Teams_SmartReply(False)
return
; -------------------------------------------------------------------------------------------------------------------

!m:: ; <--- Create eMail with link to current conversation
ShareByMail:
If GetKeyState("Ctrl") {
	;Run, "https://connectionsroot/blogs/tdalon/entry/teams_smart_reply"
	return
}
rc := TeamsCopyLink()
If (rc =false)
	return

sHTMLBody = Hello<br>Following <a href="%clipboard%">this conversation</a> in Teams:
; Create Email using ComObj
Try
	MailItem := ComObjActive("Outlook.Application").CreateItem(0)
Catch
	MailItem := ComObjCreate("Outlook.Application").CreateItem(0)
;MailItem.BodyFormat := 2 ; olFormatHTML

MailItem.Subject := linktext
MailItem.HTMLBody := sHTMLBody
;****************************** 
;~ MailItem.Attachments.Add(NewFile)
MailItem.Display ;Make email visible
;~ mailItem.Close(0) ;Creates draft version in default folder
;MailItem.Send() ;Sends the email

; Select email body
Send {Tab 3} 
Send {PgDn}
Send {Enter}
Send ^v
return

; ----------------------------------------------------------------------
#q::
SendMentions:
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/teams-shortcuts-send-mentions"
	return
}
sEmailList := Clipboard
doPerso := PowerTools_RegRead("TeamsMentionPersonalize")
Teams_Emails2Mentions(sEmailList, doPerso)

return

; ----------------------------------------------------------------------
Link2TeamsFavs:
Teams_Link2Fav()
return

; ----------------------------------------------------------------------
Email2TeamsFavs:
Teams_Emails2Favs()
return

; ----------------------------------------------------------------------
Users2Excel:
TeamLink := Clipboard
sPat = \?groupId=([^&]*)
If (RegExMatch(TeamLink,sPat,sId)) {
	sGroupId := sId1
	Teams_Users2Excel(sGroupId)
} Else
	Teams_Users2Excel()
return

; ----------------------------------------------------------------------
OpenCustomBackgrounds:
Teams_OpenBackgroundFolder()
return

OpenCustomBackgroundsLibrary:
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2021/01/teams-custom-backgrounds.html#openlib"
	return
}
sIniFile = %A_ScriptDir%\PowerTools.ini
If !FileExist(sIniFile) {
	TrayTipAutoHide("Error!","PowerTools.ini file is missing!",2000,3)
	return
}
IniRead, CustomBackgroundsLibrary, %sIniFile%, Teams, TeamsCustomBackgroundsLibrary
If (CustomBackgroundsLibrary != "ERROR")
	Run, "%CustomBackgroundsLibrary%"
Else {
    Run, notepad.exe %sIniFile%
	TrayTipAutoHide("Background Library!","Background Library location is configured in PowerTools.ini Teams->TeamsCustomBackgroundsLibrary parameter.")
}
return
; ----------------------------------------------------------------------

GetMe:
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2020/11/teams-shortcuts-smart-reply.html#getme"
	return
}
suc := People_GetMe()
If (suc) {
	TrayTipAutoHide("Personal information updated!","Email, OfficeUid, Display Name were stored to the registry.")
}
return

; ----------------------------------------------------------------------
CursorHighliter:
Run %CHFile%
return

; ######################################################################
PersonalizeMentions:
; Does not work! #TODO
Send ^a
sHtml := Clip_GetSelectionHtml()
sPat = Us)<span .* itemtype="http://schema.skype.com/Mention".*>(.*)</span>
sNewHtml := sHtml
Pos = 1 
While Pos := RegExMatch(sHtml,sPat,sFullName,Pos+StrLen(sFullName)){
    sFullName1 := RegExReplace(sFullName1," (.*)","")
	FirstName := RegExReplace(sFullName1,".*, ","")
    sNewHtml := StrReplace(sNewHtml,sFullName1,FirstName)
}
;sNewHtml := RegExReplace(sNewHtml,"s).*<html>","<html>")
WinClip.SetHTML(sNewHtml)
WinClip.Paste()

return

; ######################################################################
NotifyTrayClick_208:   ; Middle click (Button up)
Teams_MuteCb:
SendInput, !{Esc} ; for call from system tray - get active window
Teams_Mute()
Return 

NotifyTrayClick_202:   ; Left click (Button up)
Menu_Show(MenuGetHandle("Tray"), False, Menu_TrayParams()*)
Return

NotifyTrayClick_205:   ; Right click (Button up)
Teams_VideoCb:
SendInput, !{Esc} ; for call from system tray - get active window
Teams_Video()
Return 

Teams_RaiseHandCb:
SendInput, !{Esc} ; for call from system tray - get active window
Teams_RaiseHand()
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
MeetingSetup(){
; Open Teams Calendar in the Browser
Teams_OpenWebCal()
	
}


VLCStart() {
If !WinActive("ahk_exe vlc.exe") {  
    WinActivate, ahk_exe vlc.exe
    If !WinActive("ahk_exe vlc.exe") { 
        VLCExe := PowerTools_RegGet("VLCExe")
		If (VLCExe = "") 
			return
        
		Run, %VLCExe%
    }
    WinWaitActive, ahk_exe vlc.exe
}

Send ^c ; Configure Capture Device
} ; eofun
; ----------------------------------------------------------------------

VLCPlayMode(){
If !WinActive("ahk_exe vlc.exe") {  
    WinActivate, ahk_exe vlc.exe
    If !WinActive("ahk_exe vlc.exe") {
	    VLCStart()
		return
	}
}
SendInput ^h ; Minimal Interface
WinSet, AlwaysOnTop , On, ahk_exe vlc.exe
WinSet, Style, -0xC00000, ahk_exe vlc.exe ; remove title bar
} ; eofun
; ----------------------------------------------------------------------

VLCReset(){
If !WinActive("ahk_exe vlc.exe") {  
    WinActivate, ahk_exe vlc.exe
    If !WinActive("ahk_exe vlc.exe") {
	    VLCStart()
		return
	}
}
WinSet, Style, -0xC00000, ahk_exe vlc.exe ; toggle title bar
SendInput ^h ; Minimal Interface
WinSet, AlwaysOnTop , Off, ahk_exe vlc.exe
} ; eofun



; ----------------------------------------------------------------------
; https://www.autohotkey.com/boards/viewtopic.php?t=81157


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
}

; ----------------------------------------------------------------------
SetParam(MenuName){
If GetKeyState("Ctrl") {
	Run, "https://tdalon.github.io/ahk/Teams-Parameters"
	return
}
PowerTools_SetParam("Teams " . MenuName)
} ; eofun
