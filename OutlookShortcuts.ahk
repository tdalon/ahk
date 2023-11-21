; Author: Thierry Dalon
; Source: https://github.com/tdalon/ahk/blob/main/OutlookShortcuts.ahk

LastCompiled = 20230130144219

#Include <PowerTools>
#Include <Clip>

#SingleInstance force ; for running from editor
SetTitleMatchMode 2 ;allow partial match to window titles

If !a_iscompiled {
	IcoFile  := PathX(A_ScriptFullPath, "Ext:.ico").Full
	If (FileExist(IcoFile)) {
		Menu,Tray,Icon, %IcoFile%
	}
	FileGetTime, LastMod , %A_ScriptFullPath%
} Else {
	LastMod := LastCompiled
}

SubMenuSettings := PowerTools_MenuTray()
; Tooltip
If !a_iscompiled {
	FileGetTime, LastMod , %A_ScriptFullPath%
} Else {
	LastMod := LastCompiled
}
FormatTime LastMod, %LastMod% D1 R

sTooltip = Outlook Shortcuts %LastMod%`nUse 'Win+O' to open main menu in Outlook.`nClick on systray icon to access other functionalities.

Menu, Tray, Tip, %sTooltip%

; -------------------------------------------------------------------------------------------------------------------
; Add Custom Menus to MenuTray
Menu,Tray,NoStandard
Menu,Tray,Add,Open Outlook WebAccess, OWA
Menu,Tray,Add ; Separator
Menu,Tray,Standard
; -------------------------------------------------------------------------------------------------------------------

;*******************************************************************************
; Information
;*******************************************************************************
; AutoHotkey Version: 3.x
; Language: English
; Platform: XP/Vista/7
; Updated by: Toby Garcia
; Previously updated by: Ty Myrick
; Author: Lowell Heddings (How-To Geek)
; URL: http://lifehacker.com/5175724/.....gmail-keys
; Original script by: Jayp
; Original URL: http://www.ocellated.com/2009/.....t-outlook/
; Source url of this version: https://autohotkey.com/board/topic/102227-gmailkeys-for-outlook-2013/
; https://gist.github.com/mattanja/e5fffc1e72d7fa4f4aabd293af8f9053
;
; Script Function: Gmail Keys adds Gmail Shortcut Keys to Outlook
; Version 3.x updated for Outlook 2013
;
;*******************************************************************************
; Version History
;*******************************************************************************
; Version 3.1 - added delete & spam functionality and enabled move/star funcs
; Version 3.0 - updated by Toby Garcia to work with Outlook 2013
; Version 2.0 - updated by Ty Myrick to work with Outlook 2010
; Version 1.0 - updated by Lowell Heddings
; Version 0.1 - initial set of hotkeys by Jayp
;*******************************************************************************
#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.

; ----------------------------------------------------------------------

; Outlook Reminders on the top ; https://stackoverflow.com/a/35154133/2043349 
SetTitleMatchMode  2 ; windows contains
loop {
	WinWait, Reminder(s), 
	;WinSet, AlwaysOnTop, on, Reminder(s)
	;WinRestore, Reminder(s)
	WinWaitClose, Reminder(s), ,30

	WinGet WinId, ID, Reminder(s)
	SysGet, MonitorCount, MonitorCount	; or try:    SysGet, var, 80
	If (MonitorCount > 1) {
		ActiveMonIndex := Monitor_GetMonitorIndex()
		RemMonIndex := Monitor_GetMonitorIndex(WinId)
		If (ActiveMonIndex<> RemMonIndex) {
			WinActivate, ahk_id %WinId%
			SendInput #+{Right} ; Win+Shift+Right Arrow
			;Monitor_WinMove(WinId); WinShiftRight Arrow
		}
	}
} ; eo loop

; ----------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------
Menu, OutlookShortcutsMenu, add, CopyLink (Ctrl+Shift+C), CopyLink
Menu, OutlookShortcutsMenu, add, Meeting to Emails, Meeting2Emails
Menu, OutlookShortcutsMenu, add, Meeting Recipients to Excel, Meeting2Excel
Menu, OutlookShortcutsMenu, add, Item to Teams Chat, Teams_OpenChat
; -------------------------------------------------------------------------------------------------------------------

return


;************************
;Hotkeys for Outlook 2013
;************************
;As best I (Ty Myrick) can tell, the window text 'NUIDocumentWindow' is not present on 
;any other items except the main window. Also, I look for the phrase ' - Microsoft Outlook'
;in the title, which will not appear in the title (unless a user types this string into the
;subject of a message or task).

#IfWinActive, ahk_exe OUTLOOK.EXE
; Ctrl+D: Mark as completed
^d:: ; <--- Mark as Completed/ Done
If WinActive("Reminder(s) ahk_class #32770"){ ; Reminder Windows
; Keys are blocked by the UI: c,d,a,s. Alt does not work
	WinActivate
	Send +{F10} ; Shift+F10 - Open Context Menu
	Send m ; Mark as Completed
	WinSet Bottom
} else if WinActive("Inbox - ") or WinActive("Tasks - ") WinActive("To-Do List - ") {
	Send +{F10} ; Shift+F10
	Send u 
	Send m 
}	
return
; Ctrl+J
^j:: ; <---  Join Teams meeting
If WinActive("Reminder(s) ahk_class #32770"){ ; Reminder Windows
; Keys are blocked by the UI: c,d,a,s. Alt does not work
	WinActivate
	;Send +{F10} ; Shift+F10 - Open Context Menu
	SendInput {Tab}
	Send j ; Join accelerator
	TeamsExe := Teams_GetExe()
	WinWaitActive, ahk_exe %TeamsExe%,,5
	If ErrorLevel
		Return
	Sleep 1000 ; time for the window to load
	SendInput {Tab 7}{Enter}

; } else if WinActive("Inbox - ") or WinActive("Tasks - ") WinActive("To-Do List - ") {
	
}	
return
; ----------------------------------------------------------------------
!1:: ; <--- Personalize mentions
Outlook_PersonalizeMentions()
return
; ----------------------------------------------------------------------

; Ctrl+Shift+C
^+c:: ; <--- Copy Link
CopyLink:
Outlook_CopyLink()
return

#IfWinActive, - Outlook ahk_class rctrl_renwnd32
; Alt+r (Win+r is used 	by windows) 
; For calendar reminder, requires add Macro in Quick Access Toolbar (Alt+5: Set Default Reminder)
!r::  ; <--- Add Reminder and flag 
if WinActive("Inbox - ") or WinActive("Tasks - ") WinActive("To-Do List - ")
{
	Send +{F10} ; Shift+F10
	Send u 
	Send r 
} else if WinActive("Calendar - ")  {
	Send !5 
}
return

; Filter by unread
!u:: ;<--- Filter by unread
Send {Alt}
Send {H}
Send {L}
Send {u}
return

; Alt+a 
!a:: ; <--- Accept invitation
if WinActive("Calendar - ")  
{
	Send +{F10} ; Shift+F10
	Send c 
	Send s 
	Send {Enter}
}
else if WinActive("Inbox - ")
{
	Send +{F10} ; Shift+F10
	Send c 
	Send c 
	Send {Enter} 
	Send s 
	Send {Enter}
}
return

#o::
Menu, OutlookShortcutsMenu, Show
return

; Categorize All Categories -> Alt+C
; Only in Mail List view or Calendar view
!c:: ; <--- Categorize
if WinActive("Tasks - ") or WinActive("To-Do List - ") {
	Send +{F10} ; Shift+F10
	Send t 
	Send t 
	Send {Enter} 
	Send a 
} else if WinActive("Calendar - ") or WinActive(" - Outlook") {
; Window may start with the name of the current folder=> do not restrict to Inbox - 
	Send +{F10} ; Shift+F10
	Send t 
	Send a 
} 
return

; Reply to All and Delete
;a::HandleOutlookKeys("^+1", "a") ; requires Quick Step Setup with Hotkey Ctrl+Shift+1
; Reply and Delete
r:: ; <--- Reply to All and Delete
	HandleOutlookKeys("^+2", "r") ; requires Quick Step Setup with Hotkey Ctrl+Shift+2
	return
; Ctrl+Alt+R
m:: ; <--- Reply with Meeting 
	HandleOutlookKeys("^!r", "m") 
	return

; Toggle Read Ctrl+Space
^space:: ; <--- Toggle Read
	HandleOutlookKeys("!3", "^space") ; Requires add Macro in Quick Access Toolbar (Alt+3: toggle read)
	return
;  Ctrl+E
^e:: ; <--- Edit Subject
	HandleOutlookKeys("!5", "^e") ; Requires add Macro in Quick Access Toolbar (Alt+4: Edit Subject)
	return

;y::HandleOutlookKeys("^+1", "y") ;archive message using Quick Steps hotkey (ctrl+Shift+1)
;f::HandleOutlookKeys("^f", "f") ;forwards message
;r::HandleOutlookKeys("^r", "r") ;replies to message
;a::HandleOutlookKeys("^+r", "a") ;reply all
;v::HandleOutlookKeys("^+v", "v") ;move message box
;+u::HandleOutlookKeys("^u", "+u") ;marks messages as unread
;+i::HandleOutlookKeys("^q", "+i") ;marks messages as read
;j::HandleOutlookKeys("{Down}", "j") ;move down in list
;+j::HandleOutlookKeys("+{Down}", "+j") ;move down and select next item
;k::HandleOutlookKeys("{Up}", "k") ;move up
;+k::HandleOutlookKeys("+{Up}", "+k") ;move up and select next item
;o::HandleOutlookKeys("^o", "o") ;open message
s:: ; <--- Toggle Flag (star)
HandleOutlookKeys("{Insert}", "s") ;toggle flag (star)
return
; s::HandleOutlookKeys("^+g", "s") ;set follow up options (star)
;c::HandleOutlookKeys("^n", "c") ;new message
;/::HandleOutlookKeys("^e", "/") ;focus search box
.:: ; <--- Display Context Menu
HandleOutlookKeys("+{F10}", ".") ;Display context menu
return


; ------------------------------------------------------------------
Meeting2Emails:
oItem := Outlook_GetCurrentItem()
sEMailList := Outlook_Recipients2Emails(oItem)
Clip_Set(sEmailList)
TrayTipAutoHide("Meeting2Emails", "Attendees Emails were copied to the clipboard!")
return 
; ------------------------------------------------------------------
Meeting2Excel:
Outlook_Meeting2Excel()
return 


; ######################################################################
NotifyTrayClick_202:   ; Left click (Button up)
Menu_Show(MenuGetHandle("Tray"), False, Menu_TrayParams()*)
Return

NotifyTrayClick_205:   ; Right click (Button up)
Return 

;+1::HandleOutlookKeys("!4", "+1") ;Mark message as spam using Block Sender hotkey in Quick Access Toolbar (Alt+4)

;Passes Outlook a special key combination for custom keystrokes or normal key value, depending on context
HandleOutlookKeys( specialKey, normalKey )
{
	;Find out which control in Outlook has focus
	ControlGetFocus currentCtrl, A
	; MsgBox %currentCtrl%
	; MsgBox, Control with focus = %currentCtrl%
	;Set list of controls that should respond to specialKey. Controls are the list of emails and the main
	;(and minor) controls of the reading pane, including controls when viewing certain attachments.
	;Currently I handle archiving when viewing attachments of Word, Excel, Powerpoint, Text, jpgs, pdfs
	;The control 'RichEdit20WPT1'  'RichEdit20WPT5'(email subject line) is used extensively for inline editing. Thus it 
	;had to be removed. If an email's subject has focus, it won't archive...
	ctrlList = Acrobat Preview Window1,AfxWndW5,AfxWndW6,EXCEL71,MsoCommandBar1,OlkPicturePreviewer1,paneClassDC1,OutlookGrid1,OutlookGrid2,RichEdit20WPT2,RichEdit20WPT4,RICHEDIT50W1,SUPERGRID2,SUPERGRID1,
	;_WwG1 =Reading Pane
	; MsgBox % currentCtrl
	if currentCtrl in %ctrlList%
		Send %specialKey%		

	;Allow typing normalKey somewhere else in the main Outlook window. (Like the search field or the folder pane.)
	else
		Send %normalKey%
}


OWA(){
	Run, https://outlook.office.com/
}

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
