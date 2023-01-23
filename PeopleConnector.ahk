; Author: Thierry Dalon
; Standalone AHK File also delivered as compiled Standalone EXE file
; See help/homepage: https://tdalon.github.io/ahk/People-Connector

; Calls: ExtractEmails, TrayTipAutoHide, ToStartup
LastCompiled = 20211014103643

#SingleInstance force ; for running from editor

#Include <Clip>
#Include <People>
#Include <Connections>
#Include <Teams>
#Include <PowerTools>
#Include <Browser>

PT_Config := PowerTools_GetConfig()
RegRead, PT_TeamsOnly, HKEY_CURRENT_USER\Software\PowerTools, TeamsOnly
PowerTools_ConnectionsRootUrl := PowerTools_RegRead("ConnectionsRootUrl")
SubMenuSettings := PowerTools_MenuTray()

; -------------------------------------------------------------------------------------------------------------------
; SETTINGS
Menu, SubMenuSettings, Add, Notification at Startup, MenuCb_ToggleSettingNotificationAtStartup

RegRead, SettingNotificationAtStartup, HKEY_CURRENT_USER\Software\PowerTools, NotificationAtStartup
If (SettingNotificationAtStartup = "")
	SettingNotificationAtStartup := True ; Default value
If (SettingNotificationAtStartup) {
    Menu, SubMenuSettings, Check, Notification at Startup
}
Else {
    Menu, SubMenuSettings, UnCheck, Notification at Startup
}

Menu, SubMenuSettings, Add, Teams PowerShell, MenuCb_ToggleSettingTeamsPowerShell
RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If (TeamsPowerShell) 
  Menu,SubMenuSettings,Check, Teams PowerShell
Else 
  Menu,SubMenuSettings,UnCheck, Teams PowerShell


; -------------------------------------------------------------------------------------------------------------------
If !a_iscompiled 
	FileGetTime, LastMod , %A_ScriptFullPath%
 Else 
	LastMod := LastCompiled
FormatTime LastMod, %LastMod% D1 R


sTextMenuTip = Double Tap 'Shift' to open menu.`nClick on icon to access Help.
Menu, Tray, Tip, People Connector - %LastMod%`n%sTextMenuTip%
sText = Double Tap 'Shift' to open menu after selection.`nClick on icon to access Help/ Support/ Check for Updates/ Settings.

If (SettingNotificationAtStartup)
    TrayTip "People Connector is running!", %sText%
;TrayTipAutoHide("People Connector is running!",sText)

Menu, MainMenu, add, Teams &Chat, Teams2Chat
Menu, MainMenu, add, Teams &Pop out Chat, TeamsPop
Menu, MainMenu, add, &Teams Call, TeamsCall
Menu, MainMenu, add, Create Teams Meeting, TeamsMeet
Menu, MainMenu, add, Add to Teams &Favorites, Emails2TeamsFavs
Menu, MainMenu, add, Add Users to Team, Emails2TeamUsers
Menu, MainMenu, add, Teams Chat - Copy Link, TeamsChatCopyLink
Menu, MainMenu,Add ; Separator
If !(PT_TeamsOnly)
    Menu, MainMenu, add, &Skype Chat, SkypeChat
Menu, MainMenu, add, Soft Phone Call, Tel
Menu, MainMenu,Add ; Separator
Menu, MainMenu, add, Create &Email, MailTo
Menu, MainMenu, add, Create Outlook &Meeting, MeetTo
Menu, MainMenu, add, (Outlook) Meeting to Emails, Meeting2Emails
Menu, MainMenu, add, (Outlook) Meeting Recipients to Excel, Meeting2Excel
If (PowerTools_ConnectionsRootUrl != "") {
    Menu, MainMenu, add, Copy Connections &at-Mentions, CNEmail2Mention
    Menu, MainMenu,Add ; Separator
    Menu, MainMenu, add, &Open Connections Profile, CNOpenProfile
    Menu, MainMenu, add, Open Connections Network, CNOpenNetwork
}

Menu, MainMenu,Add ; Separator
Menu, MainMenu, add, Open &Delve, Delve
Menu, MainMenu, add, Open Delve Profile, DelveProfile

Menu, MainMenu, add, &Bing Search, BingSearch
Menu, MainMenu, add, &LinkedIn Search By Name, LinkedInSearchByName
Menu, MainMenu, add, Stream Profile, StreamProfile
If !(PowerTools_RegRead("PeopleViewCompanyId") = "")
    Menu, MainMenu, add, People&View OrgChart (MySuccess), PeopleView
Menu, MainMenu,Add ; Separator
Menu, MainMenu, add, Copy Office &Uids, Emails2Uids
Menu, MainMenu, add, Copy Windows Uids, Emails2WinUids
Menu, MainMenu, add, Copy Windows Uids with Domain, Emails2DUids
Menu, MainMenu, add, Uids to Emails, winUids2Emails
Menu, MainMenu, add, Copy Emails, CopyEmails
Menu, MainMenu,Add ; Separator
If (PowerTools_ConnectionsRootUrl != "") {
    Menu, SubMenuCNMentions, add, &Emails, ConnectionsMentions2Emails
    Menu, SubMenuCNMentions, add, &Teams Chat, ConnectionsMentions2TeamsChat
    Menu, SubMenuCNMentions, add, &Mentions (extract), ConnectionsMentions2Emails
    Menu, MainMenu, add, (Connections) Mentions to, :SubMenuCNMentions
}
Menu, MainMenu, add, Emails to Excel, Emails2Excel
return

Shift:: ; (Double Press) <--- Open People Connector Menu
If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 500) {
    sSelection := People_GetSelection()
    If (PowerTools_ConnectionsRootUrl != "") {
        If Browser_WinActive()
            Menu, MainMenu, Enable, (Connections) Mentions to
        Else
            Menu, MainMenu, Disable, (Connections) Mentions to
    }

    If WinActive("ahk_exe Outlook.exe") {
         Menu, MainMenu, Enable, (Outlook) Meeting to Emails
         Menu, MainMenu, Enable, (Outlook) Meeting Recipients to Excel
    } Else {
        Menu, MainMenu, Disable, (Outlook) Meeting to Emails
        Menu, MainMenu, Disable, (Outlook) Meeting Recipients to Excel
    }
    Menu, MainMenu, Show
}
return


; ######################################################################
NotifyTrayClick_202:   ; Left click (Button up)
Menu_Show(MenuGetHandle("Tray"), False, Menu_TrayParams()*)
Return

NotifyTrayClick_205:   ; Right click (Button up)
Return 

; ----------------------------  Menu Callbacks -------------------------------------
Teams2Chat: 
If GetKeyState("Ctrl") {
	Teamsy_Help("2c")
	return
}
Teams_Selection2Chat(sSelection)
return

; ------------------------------------------------------------------
TeamsMeet: 
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Teams_Emails2Meeting(sEmailList)
return

; ------------------------------------------------------------------
TeamsPop: 
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Teams_Pop(sEmailList)
return

; ------------------------------------------------------------------
TeamsChatCopyLink:
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Teams_Emails2ChatDeepLink(sEmailList)
return

TeamsCall:
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
If InStr(sEmailList,";") { ; multiple Emails
    TrayTipAutoHide("People Connector warning!","Feature does not work for multiple users!")   
    return
} Else {
    EnvGet, userprofile , userprofile
    Run,  %userprofile%\AppData\Local\Microsoft\Teams\current\Teams.exe callto:%sEmailList%
}
return

; ------------------------------------------------------------------
SkypeChat:
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
If InStr(sEmailList,";") { ; multiple Emails
    TrayTipAutoHide("People Connector warning!","Feature does not work for multiple users!")   
    return
} Else {
    sDomain := People_GetDomain
    sSip :=RegExReplace(sEmailList, "@.*","@" . sDomain)
    sCmd = "C:\Program Files (x86)\Microsoft Office\root\Office16\lync.exe" sip:%sSip%
    Run,  %sCmd%
}
return

; ------------------------------------------------------------------
Tel:
sEmailList := People_GetEmailList(sSelection)
If !sEmailList { ; empty
    If !RegexMatch(sSelection,"[1-9\(\)-\s]*")  {
        TrayTipAutoHide("People Connector warning!","You shall have an email or phone number selected!")   
    } Else {
        sTelNum := StrReplace(sSelection, " ","")
        sTelNum := StrReplace(sTelNum, "-","")
        ; If number starts with 0 prepend a 0
        If SubStr(sTelNum, 1, 1) = "0" {
	        sTelNum := SubStr(sTelNum, 2)
	        sTelNum = +49%sTelNum%
        }
        Run tel:%sTelNum%
    }
    return
}
If InStr(sEmailList,";") { ; multiple Emails
    TrayTipAutoHide("People Connector warning!","Feature does not work for multiple users!")   
    return
} Else {
    Run tel:%sEmailList%
}
return
; ------------------------------------------------------------------
Im:
sEmailList := People_GetEmailList(sSelection)
If !sEmailList { ; empty    
    return
}
sEmailList := StrReplace(sEmailList, ";",",")
;MsgBox %sEmailList%
Run im:%sEmailList%
return

CNEmail2Mention:
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
sHtmlMentions := Connections_Emails2Mentions(sEmailList)
Clip_SetHtml(sHtmlMentions)
TrayTipAutoHide("People Connector","Mentions were copied to clipboard in RTF!")   
return


ConnectionsMentions2Mentions:
sHtml := Connections_Mentions2Mentions(sSelection)
Clip_SetHtml(sHtml)
TrayTipAutoHide("Copy Mentions", "Mentions were copied to the clipboard in RTF!")
return

ConnectionsMentions2TeamsChat:
sEmailList := Connections_Mentions2Emails(sSelection)
Teams_Emails2ChatDeepLink(sEmailList)
return

ConnectionsMentions2Emails:
sEmailList := Connections_Mentions2Emails(sSelection)
Clip_Set(sEmailList)
TrayTipAutoHide("Copy Emails", "Emails were copied to the clipboard!")
return

; ------------------------------------------------------------------
CopyEmails:
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Clip_Set(sEmailList)
TrayTipAutoHide("Copy Emails", "Emails were copied to the clipboard!")
return

; ------------------------------------------------------------------
Emails2Uids:
If GetKeyState("Ctrl") {
	Run, "https://connectionsroot/blogs/tdalon/entry/people_connector_get_userid"
	return
}
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
sUidList := People_Emails2Uids(sEmailList)
clipboard := sUidList
TrayTipAutoHide("Copy Uid","Uids " . sUidList . " were copied to the clipboard!")   
return
; ------------------------------------------------------------------

Emails2WinUids:
If GetKeyState("Ctrl") {
	Run, "https://connectionsroot/blogs/tdalon/entry/people_connector_get_userid" ; TODO
	return
}
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
sUidList := People_Emails2Uids(sEmailList,"sAMAccountName")
clipboard := sUidList
TrayTipAutoHide("Copy Uid","Uids " . sUidList . " were copied to the clipboard!")   
return
; ------------------------------------------------------------------

Emails2DUids:
If GetKeyState("Ctrl") {
	Run, "https://connectionsroot/blogs/tdalon/entry/people_connector_get_domainuserid" ; TODO
	return
}
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
sUidList := People_Emails2DUids(sEmailList)
clipboard := sUidList
TrayTipAutoHide("Copy Uid","Domain\Uids " . sUidList . " were copied to the clipboard!")   
return


; ------------------------------------------------------------------
winUids2Emails:
sEmailList := winUids2Emails(sSelection)
clipboard := sEmailList
TrayTipAutoHide("Copy Emails","Emails " . sEmailList . " were copied to the clipboard!")   
return

; ------------------------------------------------------------------
Emails2TeamsFavs:

sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Teams_Emails2Favs(sEmailList)
return

; ------------------------------------------------------------------
Emails2TeamUsers:
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Teams_Emails2Users(sEmailList)
return
; ------------------------------------------------------------------

Emails2Excel:
People_Emails2Excel(sSelection[1])
return
; ------------------------------------------------------------------

MailTo:
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}

Try
	MailItem := ComObjActive("Outlook.Application").CreateItem(0)
Catch
	MailItem := ComObjCreate("Outlook.Application").CreateItem(0)
;MailItem.BodyFormat := 2 ; olFormatHTML

MailItem.To := sEmailList
MailItem.Display ;Make email visible
return
; ------------------------------------------------------------------

MeetTo:
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
Try
	oItem := ComObjActive("Outlook.Application").CreateItem(1)
Catch
	oItem := ComObjCreate("Outlook.Application").CreateItem(1)
;MailItem.BodyFormat := 2 ; olFormatHTML
oItem.MeetingStatus := 1
Loop, parse, sEmailList, ";"
{
    oItem.Recipients.Add(A_LoopField) 
}	
oItem.Display ;Make email visible
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
; ------------------------------------------------------------------
DelveProfile:
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
TenantName := PowerTools_RegRead("TenantName")
Loop, parse, sEmailList, ";"
{
   Run,  https://%TenantName%-my.sharepoint.com/person.aspx?user=%A_LoopField%&v=profiledetails
}	
return
; ------------------------------------------------------------------
Delve:
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}
TenantName := PowerTools_RegRead("TenantName")
Loop, parse, sEmailList, ";"
{
   Run,  https://%TenantName%-my.sharepoint.com/person.aspx?user=%A_LoopField%
}	
return

; ------------------------------------------------------------------
CNOpenProfile:
People_ConnectionsOpenProfile(sSelection[1])
return

CNOpenNetwork:
People_ConnectionsOpenNetwork(sSelection[1])
return

; ------------------------------------------------------------------
StreamProfile:
sSelection := sSelection[1]
sEmailList := People_GetEmailList(sSelection)
If (sEmailList != "") {
    Loop, parse, sEmailList, ";"
    {
    Run, https://web.microsoftstream.com/browse?q=%A_LoopField%&view=people
    }	
} Else {
    sName := People_GetName(sSelection)
    Run, https://web.microsoftstream.com/browse?q=%sName%&view=people
}
return

; ------------------------------------------------------------------
LinkedInSearch:
; NOT USED
sEmailList := People_GetEmailList(sSelection)
If (sEmailList != "") {
    Loop, parse, sEmailList, ";"
    {
        sName:= People_Email2Name(A_LoopField)
        Run, https://www.linkedin.com/search/results/people/?keywords=%sName%
    }	
} Else {
    sName := People_GetName(sSelection)
    Run, https://www.linkedin.com/search/results/people/?keywords=%sName%
}
return
; ------------------------------------------------------------------
LinkedInSearchByName:
sSelection := sSelection[2] ; plain text

sName := People_GetName(sSelection)
sName := RegExReplace(sName,"\d*","") ; remove any numbers
Run, https://www.linkedin.com/search/results/people/?keywords=%sName%
return

BingSearch:
sSelection := sSelection[2] ; plain text
sSelection := People_GetName(sSelection)
Run, http://www.bing.com/search?q=%sSelection%#,Person 
return

PeopleView:
People_PeopleView(sSelection[1])
return

; ----------------------------------------------------------------------

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

; ------------------------------- SUBFUNCTIONS ----------------------------------------------------------

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