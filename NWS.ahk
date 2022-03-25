; Author: Thierry Dalon
; Documentation: https://tdalon.github.io/ahk/NWS-PowerTool
; Code Project Documentation is available on GitHub here: https://github.com/tdalon/ahk
; Source: https://github.com/tdalon/ahk/blob/main/NWS.ahk

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SetWorkingDir %A_ScriptDir%

#Include <Clip>
#Include <IntelliPaste>
#Include <Connections>
#Include <Jira>
#Include <Confluence>
#Include <Login>
#Include <PowerTools>
#Include <Browser>
#Include <Teams>
#Include <SharePoint>
#Include <Explorer>
#Include <People>

LastCompiled = 20220325091524

; AutoExecute Section must be on the top of the script
#Warn All, OutputDebug

GroupAdd, Explorer, ahk_class CabinetWClass         
GroupAdd, Explorer, ahk_class ExploreWClass         
GroupAdd, Explorer, ahk_exe FreeCommander.exe
GroupAdd, Explorer, ahk_exe TOTALCMD.EXE

GroupAdd, OpenLinks, ahk_exe outlook.exe
GroupAdd, OpenLinks, ahk_exe powerpoint.exe
GroupAdd, OpenLinks, ahk_exe onenote.exe
GroupAdd, OpenLinks, ahk_exe word.exe
GroupAdd, OpenLinks, ahk_exe winword.exe
GroupAdd, OpenLinks, ahk_exe Teams.exe
GroupAdd, OpenLinks, ahk_exe lync.exe ; Skype

GroupAdd, MSOffice, ahk_exe outlook.exe
GroupAdd, MSOffice, ahk_exe powerpoint.exe
GroupAdd, MSOffice, ahk_exe POWERPNT.exe
GroupAdd, MSOffice, ahk_exe onenote.exe
GroupAdd, MSOffice, ahk_exe word.exe
GroupAdd, MSOffice, ahk_exe WINWORD.exe
GroupAdd, MSOffice, ahk_exe excel.exe
GroupAdd, MSOffice, ahk_exe teams.exe
GroupAdd, MSOffice, ahk_exe lync.exe

GroupAdd, NoIntelliPasteIns, ahk_exe XMind.exe
GroupAdd, NoIntelliPasteIns, ahk_exe freemind.exe

#SingleInstance force ; for running from editor - avoid warning another instance is running
SetTitleMatchMode, 2 ; partial match

Config := PowerTools_GetConfig() ; check also if defined
PowerTools_ConnectionsRootUrl := PowerTools_RegRead("ConnectionsRootUrl")

SubMenuSettings := PowerTools_MenuTray()
Menu,Tray,Insert,Settings,PowerTools Bundler, PowerTools_RunBundler

; -------------------------------------------------------------------------------------------------------------------
; SETTINGS
Menu, SubMenuSettings, Add, Notification at Startup, MenuCb_ToggleSettingNotificationAtStartup

RegRead, SettingNotificationAtStartup, HKEY_CURRENT_USER\Software\PowerTools, NotificationAtStartup
If (SettingNotificationAtStartup = "")
	SettingNotificationAtStartup := True ; Default value
If (SettingNotificationAtStartup) {
  Menu, SubMenuSettings, Check, Notification at Startup
} Else {
  Menu, SubMenuSettings, UnCheck, Notification at Startup
}

; IntelliPaste Hotkey setting
Menu, SubMenuSettingsIntelliPaste, Add, &Hotkey, IntelliPaste_HotkeySet
Menu, SubMenuSettingsIntelliPaste, Add, &Refresh Teams List and SPSync.ini, IntelliPaste_Refresh
Menu, SubMenuSettingsIntelliPaste, Add, Help, IntelliPaste_Help
Menu, SubMenuSettings, Add, IntelliPaste, :SubMenuSettingsIntelliPaste

Menu, SubMenuSettings, Add, Set JiraUserName, SetJiraUserName
Menu, SubMenuSettings, Add, Set Phone Number, SetPhoneNumber
Menu, SubMenuSettings, Add, Teams PowerShell, MenuCb_ToggleSettingTeamsPowerShell
RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If (TeamsPowerShell) 
	Menu,SubMenuSettings,Check, Teams PowerShell
Else 
  	Menu,SubMenuSettings,UnCheck, Teams PowerShell

; -------------------------------------------------------------------------------------------------------------------
; Setting - IntelliPasteHotkey
RegRead, IntelliPasteHotkey, HKEY_CURRENT_USER\Software\PowerTools, IntelliPasteHotkey
If ErrorLevel { ; regkey not set-> take default
	IntelliPasteHotkey = Insert
	PowerTools_RegWrite("IntelliPasteHotkey",IntelliPasteHotkey)
}

If (IntelliPasteHotkey == "Insert") {
	Hotkey, IfWinNotActive, ahk_group NoIntelliPasteIns
	Hotkey, %IntelliPasteHotkey%, IntelliPaste, On 
	Hotkey, IfWinNotActive,
} Else
	Hotkey, %IntelliPasteHotkey%, IntelliPaste, On 
         

; -------------------------------------------------------------------------------------------------------------------
; Tooltip
If !a_iscompiled 
	FileGetTime, LastMod , %A_ScriptFullPath%
 Else 
	LastMod := LastCompiled
FormatTime LastMod, %LastMod% D1 R
sTooltip = NWS PowerTool %LastMod%.`nClick on icon to access Help and Settings.
Menu, Tray, Tip, %sTooltip%

If (SettingNotificationAtStartup)
	TrayTip NWS PowerTool is running! , Click on icon to access Help and Settings.

; -------------------------------------------------------------------------------------------------------------------
; Add Custom Menus to MenuTray
Menu,Tray,NoStandard
If FileExist("Lib/Conti.ahk") & (Config = "Conti") {
	Menu,SubMenuFavs,Add, Open NWS Search, Conti_NWSSearch
	Menu,SubMenuFavs,Add, Create Ticket (ESS), SysTrayCreateTicket
	Menu,SubMenuFavs,Add, KSSE, Conti_KSSE
}
Menu, SubMenuFavs,Add, Cursor Highlighter, PowerTools_CursorHighlighter
Menu, Tray, Add, Tools, :SubMenuFavs

Menu, Tray,Add, Toggle AlwaysOnTop (Ctrl+Shift+Space), SysTrayToggleAlwaysOnTop
Menu, Tray,Add, Toggle Title Bar, SysTrayToggleTitleBar
Menu, SubMenuODB, Add, Open Permissions Settings,ODBOpenPermissions
Menu, SubMenuODB, Add, Open Document Library in Classic View,ODBOpenDocLibClassic
Menu, Tray, Add, OneDrive, :SubMenuODB

Menu, SubMenuODM, Add, ODM Set Path, ODMSetPath
Menu, SubMenuODM, Add, ODM Edit, ODMEdit
Menu, SubMenuODM, Add, ODM Run, ODMRun

Menu, SubMenuODM, Add, ODM AutoStart, MenuCb_ToggleODMAutoStart
RegRead, ODMAutoStart, HKEY_CURRENT_USER\Software\PowerTools, ODMAutoStart
If (ODMAutoStart) 
  	Menu,SubMenuODM,Check, ODM AutoStart
Else 
	Menu,SubMenuODM,UnCheck, ODM AutoStart

Menu, Tray, Add, OneDrive Mapper, :SubMenuODM

If (ODMAutoStart) {
	RegRead, ODMPath, HKEY_CURRENT_USER\Software\PowerTools, ODMPath
    If ODMPath {
		RunWait, PowerShell.exe -ExecutionPolicy Bypass -Command %ODMPath% ,, Hide
	} Else {
		TrayTipAutoHide("ODM Wrong Setup","ODM AutoStart is set but .ps1 file is not set.")
	}
}

Menu, Tray,Add ; Separator
Menu, Tray,Standard
; -------------------------------------------------------------------------------------------------------------------
; NWS Menu (Shown with Win+F1 hotkey in Browser)
Menu, NWSMenu, add, (Browser) Intelli &Copy current Url (Ctrl+Shift+C), IntelliCopyActiveUrl
Menu, NWSMenu, add, (Browser) Share by E&mail current Url (Ctrl+Shift+M), EmailShareActiveUrl
Menu, NWSMenu, add, (Browser) Share Url to Teams, TeamsShareActiveUrl
Menu, NWSMenu, add, (Browser) Quick &Search (Win+F), QuickSearch

If FileExist("Lib/Conti.ahk") & (Config = "Conti")
	Menu, NWSMenu, add, (Browser) Create IT &Ticket, Conti_CreateTicket
; -------------------------------------------------------------------------------------------------------------------

; EDIT : SCRIPT PARAMETERS
DefExplorerExe := "explorer.exe" ;*[NWS]
IfNotExist, %DefExplorerExe%
	DefExplorerExe := "explorer.exe"
		
; Start VPN (only for Conti config)
If FileExist("Lib/Conti.ahk") & (Config = "Conti") {
	If ! (Login_IsNet("Conti"))
		Login_VPNConnect()
}

return

; ####################################################################
; Hotkeys

; -------------------------------------------------------------------------------------------------------------------
^+Space:: WinSet, AlwaysOnTop, Toggle, A
return
;================================================================================================
;  CapsLock processing.  Must double tap CapsLock to toggle CapsLock mode on or off.
; https://www.howtogeek.com/446418/how-to-use-caps-lock-as-a-modifier-key-on-windows/
;================================================================================================
; Must double tap CapsLock to toggle CapsLock mode on or off.
CapsLock:: ; <--- Must double tap CapsLock to toggle CapsLock mode on or off.
    KeyWait, CapsLock                                                   ; Wait forever until Capslock is released.
    KeyWait, CapsLock, D T0.2                                           ; ErrorLevel = 1 if CapsLock not down within 0.2 seconds.
    if ((ErrorLevel = 0) && (A_PriorKey = "CapsLock") )                 ; Is a double tap on CapsLock?
        {
        SetCapsLockState, % GetKeyState("CapsLock","T") ? "Off" : "On"  ; Toggle the state of CapsLock LED
        }
return


;================================================================================================
; Hotkeys with CapsLock modifier.  See https://autohotkey.com/docs/Hotkeys.htm#combo
;================================================================================================

#If (PowerTools_ConnectionsRootUrl != "")
CapsLock & c:: ;  <--- Connections Global Search
sSelection:= Clip_GetSelection()
Run, https://%PowerTools_ConnectionsRootUrl%/search/web/search?query=%sSelection%     ; Launch with contents of clipboard
Return

#If

CapsLock & d:: ;  <--- Get DEFINITION of selected word.
sSelection:= Clip_GetSelection()
Run, http://www.google.com/search?q=define+%sSelection%     ; Launch with contents of clipboard
Return

CapsLock & s:: ;  <--- Search in Scaledagile.com.
sSelection:= Clip_GetSelection()
Run, "https://www.google.com/search?q=site:https://www.scaledagileframework.com %sSelection%"     ; Launch with contents of clipboard
Return

CapsLock & b:: ;  <--- Bing search.
sSelection:= Clip_GetSelection()
If RegExMatch(sSelection,"(.*), (.*) <(.*)>",sMatch) ; From Outlook contact
	sSelection = %sMatch2% %sMatch1%#,Person ; Transform Firstname Lastname
Run, https://www.bing.com/search?q=%sSelection%     ; Launch with contents of clipboard
Return


CapsLock & g:: ; <--- GOOGLE the selected text.
sSelection:= Clip_GetSelection()
Run, "https://www.google.com/search?q=%sSelection%"             ; Launch with contents of clipboard
Return

;CapsLock & t:: ; <--- Do THESAURUS of selected text
;sSelection:= Clip_GetSelection()
;Run http://www.thesaurus.com/browse/%sSelection%             ; Launch with contents of clipboard
;Return


CapsLock & w:: ; <--- Do WIKIPEDIA of selected text
sSelection:= Clip_GetSelection()
Run, https://en.wikipedia.org/wiki/%sSelection%              ; Launch with contents of clipboard
Return

#If FileExist("Lib/Conti.ahk") & (Config = "Conti")
CapsLock & n:: ; <--- Open NWS Search or trigger NWS Search with selected text
FunStr = Conti_NWSSearch
%FunStr%()
return

#If

CapsLock & y:: ; <--- YouTube search of selected text
sSelection:= Clip_GetSelection()
Run, https://www.youtube.com/results?search_query=%sSelection%              ; Launch with contents of clipboard
return


; -------------------------------------------------------------------------------------------------------------------
;   All Applications
; -------------------------------------------------------------------------------------------------------------------
; Ctrl+Alt+V
^!v:: ; <--- VPN Connect
Login_VPNConnect()
return

; -------------------------------------------------------------------------------------------------------------------
; Ctrl+F12
^F12:: ; <--- Paste clean url with url decoded
PasteCleanUrl()
return	

; -------------------------------------------------------------------------------------------------------------------
; Ctrl+Ins: paste clean url without uridecode - unbroken link
; useful to paste unbroken link e.g. in Connections comments
^Ins:: ; <--- Paste Clean Url
PasteCleanUrl(true)	
return	

; -------------------------------------------------------------------------------------------------------------------
#IfWinActive, ahk_group OpenLinks
; Open in Default Browser (incl. Office applications) - see OpenLink function
; Middle Mouse Click
MButton:: ; <--- Open in Preferred Browser
; If target window is not under focus, e.g. MButton on Chrome Tab
Clip_All := ClipboardAll  ; Save the entire clipboard to a variable
Clipboard =  ; Empty the clipboard to allow ClipWait work

sleep, 200 ;(wait in ms) leave time to release the Shift 
SendEvent {RButton} ;Click Right does not work in Outlook embedded tables
sleep, 200 ;(wait in ms) give time for the menu to popup	

If WinActive("ahk_exe onenote.exe")
	SendInput i ; Copy Link
Else
	SendInput c ; Copy Link

ClipWait, 2
sURL := Clipboard

If sURL { ; Not empty
	;OpenLink(sURL)
	sURL := IntelliPaste_CleanUrl(sURL) ; convert e.g. teams links to SP links
	Run %sURL% ; Handled by BrowserSelect
} Else {
	;MsgBox OpenLinks: Empty URL/ Error
	Send {MButton}
}		

Clipboard := Clip_All ; Restore the original clipboard
return


; Ctrl+Shift+V  Paste in plain text/ removing rich-text formatting like links
; http://stackoverflow.com/a/132826/2043349
; https://lifehacker.com/better-paste-takes-the-annoyance-out-of-pasting-formatt-5388814
; Exclude for Excel for AddIn Ctrl+Shift+V: open IMS issue in document - use MS Office paste option instead
#IfWinNotActive, ahk_exe EXCEL.EXE
^+v:: ; <--- Paste in plain text format
Clip_Paste(Clipboard)
return

#If WinActive("ahk_exe Code - Insiders.exe") || WinActive("ahk_exe Code.exe")
!c:: ; Alt+C Toggle Block comment Uncomment. Block need to be selected
ClipBackup:= ClipboardAll
sSelection := Clip_GetSelection(False)
If !sSelection { ; no sSelection
    Clip_Restore(ClipBackup)
    Return
}
If RegExMatch(sSelection,"s)^/\*.*\*/$") {
	sNew := SubStr(sSelection, InStr(sSelection,"`n") + 1) ; remove first line
	sNew := SubStr(sNew, 1,InStr(sNew, "`r" , , -1) -1) ; remove last line
} Else {
	sNew = /*`n%sSelection%`n*/
}

Clip_Paste(sNew)
Clip_Restore(ClipBackup)
return

; -------------------------------------------------------------------------------------------------------------------
#If WinActive(".ahk")
!h:: ; Alt+h AHK Open Help command
kw := Clip_GrabWord()
AHK_Help(kw)
;Run,% "https://www.autohotkey.com/docs/commands/" cmd ".htm"
Return

; -------------------------------------------------------------------------------------------------------------------
;   BROWSER Group
; -------------------------------------------------------------------------------------------------------------------
#If Browser_WinActive()
/*
	; Ctrl + Alt + V - remove quotes for MySuccess copy/paste of goals from Excel
	#IfWinActive,ahk_group Browser
	^!v:: 
	ClipSaved := ClipboardAll
	sURL := clipboard
	sURL := StrReplace(sURL,"""","")
	;MsgBox Clean url:`n%sURL%
	clipboard = ; Empty the clipboard
	Clipboard := sURL
	ClipWait, 0.5
	Send ^v
	Sleep 100 ; pause necessary because of lag in browser (no problem in Notepad e.g.)- next command restore clipboard runs asynchron before paste
	; https://autohotkey.com/board/topic/37029-good-practices-with-clipboard/#entry233156	
	Clipboard := ClipSaved ; restore clipboard
	
	return
*/

; Win+F 
#f:: ; <--- [Browser] Run Quick Search (Connections, Confluence, Jira, Google)
QuickSearch()
return

; Ctrl+E - like Explorer or Edit - from Browser
; Do not use Alt key because of issue with IE
^e:: ; <--- [Browser] Edit Connections or Open SharePoint in File Explorer
;!e:: ; Alt+E because Ctrl+E is used and can not be overwritten with Windows 10 and IE/Edge Browser Universal App
sURL := Browser_GetUrl()
If Connections_IsUrl(sURL) { 
	Connections_Edit(sURL)
	return
} Else If Blogger_IsUrl(sURL) { 
	Blogger_Edit(sURL)
	return
} Else If SharePoint_IsUrl(sURL) { ; SharePoint
	newurl:= SharePoint_CleanUrl(sURL) ; returns wihout ending /
	; For o365 SharePoint check if file is synced in SPsync.ini
	If SharePoint_IsSPWithSync(newurl) { ; mspe can also offers Sync
		sFile := SharePoint_Url2Sync(sUrl)
		If (sFile=""){
			TrayTipAutoHide("NWS PowerTool","SharePoint is not Sync'ed or OneDrive SPSync.ini File is not properly configured!",3,0x3)
			Run "%sIniFile%"
			return
		}
		Run %DefExplorerExe% "%sFile%"

		
	} Else { ; SharePoints without Sync-> use Dav access
		newurl:=StrReplace(newurl,"https:","")
		newurl:=StrReplace(newurl,"+"," ") ; strange issue with blank converted to +
		newurl:=StrReplace(newurl,"/","\")
		newurl:= RegExReplace(newurl,"https?//[^/]*","$0@ssl\DavWWWroot")	; without @ssl it takes too long to open
		Run %DefExplorerExe% "%newurl%"
	}	 	
} Else {
	TrayTipAutoHide("NWS PowerTool",sURL . " did not match a Connections, SharePoint or Blogger url!",,0x2)
}
return


#F1:: ; <--- [Browser] Open NWS PowerTool Menu
; Win+F1
Menu, NWSMenu, Show
return

; IntelliCopyActiveURL Ctrl+Shift+C
#If Browser_WinActive()

^+c:: ; <--- [Browser] Intelli Copy Active Url
IntelliCopyActiveUrl:
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	Run, "https://connectionsroot/blogs/tdalon/entry/intelli_copy_active_url"
	return
}
WinGetActiveTitle, linktext
sLink := Browser_GetUrl()
sLink := IntelliPaste_CleanUrl(sLink)

; Remove trailing - containing program e.g. - Google Chrome
StringGetPos,pos,linktext,%A_space%-,R
if (pos >=0)
	linktext := SubStr(linktext,1,pos)
If Connections_IsUrl(sLink,"blog") { ; Connections Blog
	linktext := StrReplace(linktext,"Blog Blog","Blog")
}
Else If Jira_IsUrl(sLink){
	StringGetPos,pos,linktext,%A_space%-,R
	if (pos >=0)
		linktext := SubStr(linktext,1,pos)
}
sHtml =	<a href="%sLink%">%linktext%</a>   
;Clip_SetHtml(sHtml,sLink) ; WinClip.GetHtml does not work afterwards
WinClip.SetHTML(sHtml)
WinClip.SetText(sLink)
TrayTipAutoHide("NWS PowerTool","Link was copied to the clipboard!")

return

; IntelliSharebyEmailActiveURL Ctrl+Shift+M
#If Browser_WinActive()

^+m:: ; <--- [Browser] Share by eMail active url
EmailShareActiveUrl:
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	Run, "https://connectionsroot/blogs/tdalon/entry/share_link_by_email"
	return
}
sLink := Browser_GetUrl()
sLink := IntelliPaste_CleanUrl(sLink)
WinGetActiveTitle, linktext
; Remove trailing - containing program e.g. - Google Chrome
StringGetPos,pos,linktext,%A_space%-,R
if (pos != -1)
	linktext := SubStr(linktext,1,pos)

If Connections_IsUrl(sLink,"blog")  { ; Connections Blog
	linktext := StrReplace(linktext,"Blog Blog","Blog")
}
sHTMLBody = Hello<br>I thought you might be interested in this post: <a href="%sLink%">%linktext%</a>.<br>
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

return


; -------------------------------------------------------------------------------------------------------------------
;     #CHROME BROWSER
; -------------------------------------------------------------------------------------------------------------------
#IfWinActive ahk_exe chrome.exe
;https://autohotkey.com/board/topic/84792-opening-a-link-in-non-default-browser/
; Shift+ middle mouse button
+MButton:: ;  <--- [Chrome] Open Link in IE

SavedClipboard := ClipboardAll  ; Save the entire clipboard to a variable
Clipboard := ""  ; Empty the clipboard to allow ClipWait work

Click Right ; Click Right mouse button
sleep, 100 ;(wait in ms) give time for the menu to popup
Sendinput e ; Send the underlined key that copies the link from the right click menu. see https://productforums.google.com/forum/#!topic/chrome/CPi4EmhqHPE
ClipWait, 2

sURL := Clipboard
If (sURL = "") {
	Exit
}

Run, iexplore.exe %sURL%
;Sleep 1000
;WinActivate, ahk_exe, iexplore.exe
;Run, C:\Users\%A_UserName%\AppData\Local\Google\Chrome\Application\chrome.exe %clipboard% ; Open in Google Chrome
;Run, %A_ProgramFiles%\Mozilla Firefox\firefox.exe %clipboard%; open in Firefox

Clipboard := SavedClipboard ; Restore the original clipboard
return

; Shift Mouse click
+LButton:: ; <--- [Chrome] Open link in preferred browser
; Calls OpenLink
SavedClipboard := ClipboardAll  ; Save the entire clipboard to a variable
Clipboard := ""  ; Empty the clipboard to allow ClipWait work

Click Right ; Click Right mouse button
sleep, 100 ;(wait in ms) give time for the menu to popup
SendInput e ; Send the underlined key that copies the link from the right click menu. see https://productforums.google.com/forum/#!topic/chrome/CPi4EmhqHPE
ClipWait, 2

sURL := Clipboard
If (sURL = "") {
	Exit
}

sURL := IntelliPaste_CleanUrl(sURL)
OpenLink(sURL)

Clipboard := SavedClipboard ; Restore the original clipboard
return

; Ctrl+Right mouse button
^RButton:: ; <--- [Chrome] Open link in File Explorer

SavedClipboard := ClipboardAll  ; Save the entire clipboard to a variable
Clipboard := ""  ; Empty the clipboard to allow ClipWait work

Click Right ; Click Right mouse button
sleep, 100 ;(wait in ms) give time for the menu to popup
SendInput e ; Send the underlined key that copies the link from the right click menu. see https://productforums.google.com/forum/#!topic/chrome/CPi4EmhqHPE
ClipWait, 2

sURL := Clipboard
If !sURL
	Exit

If SharePoint_IsUrl(sURL) {
	newurl:=SharePoint_CleanUrl(sURL)
	newurl:=StrReplace(newurl,"https:","")
	newurl:=StrReplace(newurl,"+"," ") ; strange issue with blank converted to +
	newurl:=StrReplace(newurl,"/","\")
	newurl:= RegExReplace(newurl,"https?//[^/]*","$0@ssl\DavWWWroot")	; without @ssl it takes too long to open
	Run %DefExplorerExe% "%newurl%" 
}	

Clipboard := SavedClipboard ; Restore the original clipboard
return


; -------------------------------------------------------------------------------------------------------------------
;    EXPLORER Group
; -------------------------------------------------------------------------------------------------------------------
#ifWinActive,ahk_group Explorer              ; Set hotkeys to work in explorer only
; Open file With Notepad++ from Explorer using Alt+N hotkey
; https://autohotkey.com/board/topic/77665-open-files-with-portable-notepad/


; -------------------------------------------------------------------------------------------------------------------
; Alt+N 
!n:: ; <--- [Explorer] Open file in Notepad++
ClipSaved := ClipboardAll
Clipboard := ""
SendInput ^c
ClipWait, 0.5
file := Clipboard
Clipboard := ClipSaved
Run, notepad++.exe "%file%" 
return

; -------------------------------------------------------------------------------------------------------------------
; Override Delete key for Sync location
/*
	$Del:: ; <--- [Explorer] Safeguard Delete if in  ODB Sync location
	sFile := Explorer_GetSelection()
	; if no file selected in File Explorer
	If (!sFile) ; file empty
		return
	EnvGet, sOneDriveDir , onedrive
	sOneDriveDir := StrReplace(sOneDriveDir,"OneDrive - ","")	
	If InStr(sFile,sOneDriveDir . "\") { 
		MsgBox 0x14, Delete?,Are you sure you want to delete in your Sync location?`nIt might also delete the file in the SharePoint / not only locally for you, if sync is active.
		IfMsgBox, No
			return
	}
	Send {Delete}
	return
*/
; -------------------------------------------------------------------------------------------------------------------
; Ctrl+E Open SharePoint File from mapped Document Library or Sync location in Default Browser
; Calls: GetFileLink, GetExplorerSelection
^e:: ;	<--- [Explorer] Open SharePoint file selection in IE Browser	
sFile := Explorer_GetSelection()
; if no file selected in File Explorer
If !sFile ; empty
{
	MsgBox "You need to select a file!" 		
	return
}

; For multi-section take the last one
sFile := RegExReplace(sFile,"`n.*","")

sFile := GetFileLink(sFile)
If (!sFile) ; file empty
	return

SplitPath, sFile, OutFileName, OutDir	
If InStr(OutFileName,".")  ; then a file is selected (Last part in Path containing "." for file extension)
	Run, "%OutDir%" ; Open parent directory
Else
	Run, "%sFile%"

return
; -------------------------------------------------------------------------------------------------------------------
; Ctrl+O
; Calls: GetFileLink, GetExplorerSelection, OpenFile
^o:: ; <--- [Explorer] Open file
sFiles := Explorer_GetSelection()
; if no file selected in File Explorer
If sFiles =
{
	MsgBox "You need to select a file!" 		
	return
}

Loop, parse, sFiles, `n, `r
{
	sFile := A_LoopField	
	If InStr(sFile,".xlsx") or  InStr(sFile,".docx") or InStr(sFile,".pptx") or InStr(sFile,".xlsm") or  InStr(sFile,".docm") or InStr(sFile,".pptm") {
		sFile := GetFileLink(sFile)
		Run, iexplore.exe "%sFile%" ; BUG: Edge can not open file links
	} Else
		Run, Open "%sFile%"
}

return

; -------------------------------------------------------------------------------------------------------------------
; Ctrl+K
; Calls: GetFileLink, GetExplorerSelection, OpenFile
^k:: ; <--- [Explorer] Copy File Link (OneDrive)
Send +{F10} ; Shift+F10
Send s 
Send {Enter}
Sleep 2000 ; Time to load the UI
Send {Tab 3} 
Send {Enter}
Send {Esc}
return



; ######################################################################
NotifyTrayClick_202:   ; Left click (Button up)
Menu_Show(MenuGetHandle("Tray"), False, Menu_TrayParams()*)
Return

NotifyTrayClick_205:   ; Right click (Button up)
Menu, NWSMenu, Show
Return 

SysTrayToggleAlwaysOnTop:
SendInput, !{Esc}
WinSet, AlwaysOnTop, Toggle, A
;Tooltip("Toggle Active Window AlwaysOnTop",1000)
return

SysTrayToggleTitleBar:
SendInput, !{Esc}
WinSet, Style, ^0xC00000, A ; toggle title bar
return

#If FileExist("Lib/Conti.ahk")

SysTrayCreateTicket:
SendInput, !{Esc}
FunStr = Conti_CreateTicket
%FunStr%()
return

; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
; FUNCTIONS
; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------
;IsIELink(url)
; true if link shall be opened with Internet Explorer rather than another browser e.g. Chrome because of incompatibility
IsIELink(sURL){
	If SharePoint_IsUrl(sURL) || InStr(sURL,"file://") || InStr(sURL,"/pkit/") || InStr(sURL,"/BlobIT/") || InStr(sURL,"/openscapeuc/dial/") 
		return true
	Else 	
		return false	
}


; -------------------------------------------------------------------------------------------------------------------
PasteCleanUrl(encode:= False){
	; encode: True/False
	; calls: CleanUrl
	; called by Hotkey Ctrl+Ins [decode=false] and Ctrl+F12 [decode=true]
	ClipSaved := ClipboardAll
	sURL := clipboard

	sURL := GetFileLink(sURL)
	sURL := IntelliPaste_CleanUrl(sURL)

	If (encode) {
		sURL := uriEncode(sURL)
		sURL := StrReplace(sURL,":","%3A")
		sURL := StrReplace(sURL,"https%3A","https:")
		sURL := StrReplace(sURL,"http%3A","http:")
	}
	
	;MsgBox Clean url:`n%sURL%
	;sendInput % sURL
	
	Clip_Paste(sURL)
	return
}


; -------------------------------------------------------------------------------------------------------------------
; Function OpenLink
; Open Link in Default browser
OpenLink(sURL) {
	If IsIELink(sURL) {
		;Run, %A_ProgramFiles%\Internet Explorer\iexplore.exe %sURL%
		Run, iexplore.exe "%sURL%"
		;Sleep 1000
		;WinActivate, ahk_exe, iexplore.exe
	} Else If Confluence_IsUrl(sURL)  || Jira_IsUrl(sURL)  { ; JIRA or Confluence urls							
		;Run, C:\Users\%A_UserName%\AppData\Local\Google\Chrome\Application\chrome.exe %sURL%
		Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 3" "%sURL%"
	} Else If InStr(sURL,"https://teams.microsoft.com/") { ; Teams url=>open by default in App
		;Run, StrReplace(sURL, "https://","msteams://")
		Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 4" "%sURL%"
	} Else {
		;sURL := IntelliPaste_CleanUrl(sURL) ; No need to clean because handled by Redirector
		Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 1" "%sURL%"
	} ; End If
} ; End Function OpenLink

; ----------------------------------------------------------------------
QuickSearch(){
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	sUrl := "https://connectionsroot/wikis/home/wiki/Wc4f94c47297c_42c8_878f_525fd907cb68/page/NWS%20PowerTool?section=%28Win%2BF%29_Quick_Search_%28Connections%2C_Jira%2C_Confluence%29"
	Run, "%sUrl%"
	return
}
sUrl := Browser_GetUrl()
If !sUrl { ; empty
	MsgBox Cannot get URL ; DBG
	return
}

; Make Libraries optional
QuickSearches := "Confluence,Connections,Blogger,Jira"
Loop, parse, QuickSearches, `,
{
	FunStr := A_LoopField . "_IsUrl"
	If IsFunc(FunStr) {
		If  %FunStr%(sUrl) {
			FunStr := A_LoopField . "_Search"
			%FunStr%(sUrl)
			return
		}		
	} 
}

If RegExMatch(sUrl,"youtube\.com/(?:c|channel)/") { ; YouTube Channel Search
	; https://www.youtube.com/c/KevinStratvert/search?query=remove%20background
	;https://www.youtube.com/channel/UCfJT_eYDTmDE-ovKaxVE1ig/search?query=background
	sPat = youtube\.com/(?:c|channel)/([^/]*)/search\?query=(.*)
	sDefSearch =
	If RegExMatch(sUrl,sPat, sMatch) {
		sDefSearch := StrReplace(sMatch2,"%20"," ")
		sDefSearch := StrReplace(sDefSearch,"+"," ")
		sChannelName := sMatch1
	} Else {
		sPat = youtube\.com/(?:c|channel)/([^/]*)
		RegExMatch(sUrl,sPat, sMatch)
		sChannelName := sMatch1
	} 
	
	InputBox, sSearch , YouTube Channel Search, Enter search string:,,640,125,,,,, %sDefSearch%
	if ErrorLevel
		return
	sSearch := Trim(sSearch)
	
	sSearchUrl = https://www.youtube.com/channel/%sChannelName%/search?query=%sSearch%
	SendInput ^t^l ; close current search window
	Clip_Paste(sSearchUrl)
	SendInput {Enter}
	SendInput ^{Tab}
	Sleep 500
	SendInput ^w
} Else If InStr(sUrl,"google.com/search?q=") and !InStr(sUrl,"site:") { ; simple google search
	sPat = google.com/search\?q=([^&]*) 
	sPat := StrReplace(sPat,".","\.")
	RegExMatch(sUrl,sPat, sMatch)
	sSearch := StrReplace(sMatch1,"%20"," ")
	sSearch := StrReplace(sSearch,"+"," ")
	InputBox, sSearch , Google Search, Enter search string:,,640,125,,,,, %sSearch%
	if ErrorLevel
		return
	sSearch := Trim(sSearch)
	sSearchUrl = https://www.google.com/search?q=%sSearch%
	SendInput ^t^l 
	Clip_Paste(sSearchUrl)
	SendInput {Enter}
	; close previous search window
	SendInput ^{Tab}
	Sleep 500
	SendInput ^w

} Else {
	sPat = google.com/search\?q=site:([^`%]*)`%20 ; Chrome strips https://www. for google
	sPat := StrReplace(sPat,".","\.")
	If RegExMatch(sUrl,sPat . "(.*)", sMatch) { ; https://www.google.com/search?q=site:https://scaledagileframework.com%20pipeline
		sDefSearch := StrReplace(sMatch2,"%20"," ")
		sRootUrl := sMatch1
		;SendInput ^w ; close current search window
	} Else {
		RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
	}
	InputBox, sSearch , Google Site Search, Enter search string:,,640,125,,,,, %sDefSearch%
	if ErrorLevel
		return
	sSearch := Trim(sSearch)
	
	sSearchUrl = https://www.google.com/search?q=site:%sRootUrl% %sSearch%
	SendInput ^t^l 
	Clip_Paste(sSearchUrl)
	SendInput {Enter}
	; close previous search window
	SendInput ^{Tab}
	Sleep 500
	SendInput ^w
}	
} ; eofun
; ----------------------------------------------------------------------


; ----------------------------------------------------------------------
SetPhoneNumber(ItemName){
If GetKeyState("Ctrl") {
	sUrl := "https://connectionsroot/blogs/tdalon/entry/Connections2ticket_ahk"
	Run, "%sUrl%"
	return
}
PowerTools_SetSetting(ItemName)
}

SetJiraUserName(ItemName){
If GetKeyState("Ctrl") {
	sUrl := "https://tdalon.github.io/ahk/NWS-PowerTool"
	Run, "%sUrl%"
	return
}
PowerTools_SetSetting(ItemName)
}




; ----------------------------------------------------------------------
ODBOpenPermissions(){
If GetKeyState("Ctrl") {
	sUrl := "https://connectionsroot/blogs/tdalon/entry/onedrive_open_persmissions_settings_powertool" 
	Run, "%sUrl%"
	return
}
OfficeUid := People_GetMyOUid()
TenantName := PowerTools_GetSetting("TenantName")
Domain := People_GetDomain()
Domain := StrReplace(Domain,".","_")
Run https://%TenantName%-my.sharepoint.com/personal/%OfficeUid%_%Domain%/_layouts/15/user.aspx	
}
; ----------------------------------------------------------------------
ODBOpenDocLibClassic(){
If GetKeyState("Ctrl") {
	sUrl := "https://connectionsroot/blogs/tdalon/entry/onedrive_alert#Shortcut_/_PowerTool_way"
	Run, "%sUrl%"
	return
}
OfficeUid := People_GetMyOUid()
Domain := People_GetDomain()
Domain := StrReplace(Domain,".","_")
TenantName := PowerTools_GetSetting("TenantName")
sUrl := "https://%TenantName%-my.sharepoint.com/personal/" . OfficeUid . "_%Domain%/Documents/Forms/All.aspx?ShowRibbon=true&InitialTabId=Ribbon%2ELibrary&VisibilityContext=WSSTabPersistence"
Run, "%sUrl%"	
}

; ----------------------------------------------------------------------
TeamsShareActiveUrl:
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	Run, "https://tdalon.blogspot.com/share-to-teams"
	return
}
sLink := Browser_GetUrl()
sLink := IntelliPaste_CleanUrl(sLink)
sLink = https://teams.microsoft.com/share?href=%sLink%
Run %sLink%
return

; ---------------------------------------------------------------------- STARTUP -------------------------------------------------
MenuCb_ToggleSettingNotificationAtStartup:
If (SettingNotificationAtStartup := !SettingNotificationAtStartup) {
  Menu, SubMenuSettings, Check, Notification at Startup
}
Else {
  Menu, SubMenuSettings, UnCheck, Notification at Startup
}
PowerTools_RegWrite("NotificationAtStartup",SettingNotificationAtStartup)
return

; ----------------------------------------------------------------------



MenuCb_ToggleODMAutoStart(ItemName, ItemPos, MenuName){
If GetKeyState("Ctrl") {
    sUrl := "https://connectionsroot/blogs/tdalon/entry/OneDrive_Mapper"
    Run, "%sUrl%"
	return
}
RegRead, ODMAutoStart, HKEY_CURRENT_USER\Software\PowerTools, ODMAutoStart
ODMAutoStart := !ODMAutoStart
If (ODMAutoStart) {
 	RegRead, ODMPath, HKEY_CURRENT_USER\Software\PowerTools, ODMPath
    If !ODMPath {
		ODMPath := ODMSetPath()
		If ODMPath ; If no path entered cancel
			return
	}
	Menu,%MenuName%,Check, %ItemName%
	TrayTipAutoHide("OneDrive Mapper","OneDrive Mapper will auto-start with this script.")	

} Else {
    Menu,%MenuName%,UnCheck, %ItemName%	 
	TrayTipAutoHide("OneDrive Mapper","OneDrive Mapper auto-start was switched OFF.")
}
PowerTools_RegWrite("ODMAutoStart",ODMAutoStart)
}



ODMSetPath(){
If GetKeyState("Ctrl") {
    sUrl := "https://connectionsroot/blogs/tdalon/entry/OneDrive_Mapper"
    Run, "%sUrl%"
	return
}
	
FileSelectFile, ODMPath , 1, OneDriveMapper.ps1, Browse for your OneDriveMapper.ps1 location
If (ODMPath = "") or !InStr(ODMPath,"OneDriveMapper.ps1") {
	TrayTipAutoHide("OneDrive Mapper Setup","OneDriveMapper.ps1 wasn't selected!")
	ODMPath =
} Else {
	PowerTools_RegWrite("ODMPath",ODMPath)
}
return ODMPath
}


ODMEdit(){
If GetKeyState("Ctrl") {
    sUrl := "https://connectionsroot/blogs/tdalon/entry/OneDrive_Mapper"
    Run, "%sUrl%"
	return
}
RegRead, ODMPath, HKEY_CURRENT_USER\Software\PowerTools, ODMPath
If !ODMPath {
	TrayTipAutoHide("OneDrive Mapper Setup","OneDriveMapper.ps1 wasn't selected! Set ODM Path.")
	return
}
sCmd = Edit "%ODMPath%"
Run %sCmd%
}

ODMRun(){
If GetKeyState("Ctrl") {
    sUrl := "https://connectionsroot/blogs/tdalon/entry/OneDrive_Mapper"
    Run, "%sUrl%"
	return
}
RegRead, ODMPath, HKEY_CURRENT_USER\Software\PowerTools, ODMPath
If ODMPath {
	RunWait, PowerShell.exe -ExecutionPolicy Bypass -Command %ODMPath% ,, Hide
} Else {
	TrayTipAutoHide("OneDrive Mapper Run","OneDriveMapper.ps1 wasn't selected! Set ODM Path.")
}
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
