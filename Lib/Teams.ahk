; Documentation of Lib dependencies
#Include <People>
#Include <UriDecode>
#Include <Teamsy>
#Include <Clip>
#Include <FindText>
#Include <UIA_Interface>

Teams_Launcher(){
Teamsy("-g")
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Teams_OpenBackgroundFolder(){
BackgroundDir = %A_AppData%\Microsoft\Teams\Backgrounds\Uploads
If !FileExist(BackgroundDir) {
    FileCreateDir, %BackgroundDir%
}
Run, %BackgroundDir%

} ; eofun
; -------------------------------------------------------------------------------------------------------------------


Teams_Emails2ChatDeepLink(sEmailList, askOpen:= true){
; Copy Chat Link to Clipboard and ask to open
sLink := "https://teams.microsoft.com/l/chat/0/0?users=" . StrReplace(sEmailList, ";",",") ; msteams:
If InStr(sEmailList,";") { ; Group Chat
    InputBox, sTopicName, Enter Group Chat Name,,,,100
    if (ErrorLevel=1) { ; no TopicName defined
	    sLinkDisplayText = Group Chat Link
    } Else { ; No topic defined
        sLink := sLink . "&topicName=" . StrReplace(sTopicName, ":", "")
        sLinkDisplayText = %sTopicName% (Group Chat Link)
    }
} Else {
    sName := RegExReplace(sEmailList,"@.*" ,"")
    sName := StrReplace(sName,"." ," ")
    StringUpper, sName, sName , T
    sLinkDisplayText = Chat with %sName%
    ;InputBox, sTopicName, Chat Link Display Name,,,,100,,,,, %sLinkDisplayText%
}

sHtml = <a href="%sLink%">%sLinkDisplayText%</a>
Clip_SetHtml(sHtml,sLink)
If askOpen {
	MsgBox 0x1034,People Connector, Teams link was copied to the clipboard. Do you want to open the Chat?
	IfMsgBox Yes
		Run, %sLink% 
}
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Emails2Chat(sEmailList,sTopicName :=""){
; Open Teams 1-1 Chat or Group Chat from list of Emails
; See https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-link-teams
sLink := "msteams:/l/chat/0/0?users=" . StrReplace(sEmailList, ";",",") ; msteams:
If InStr(sEmailList,";") { ; Group Chat
    InputBox, sTopicName , To Chat, Enter Group Chat Name, , , , , , , ,%sTopicName%
    if (ErrorLevel=0) { ;  TopicName defined
        sLink := sLink . "&topicName=" . StrReplace(sTopicName, ":", "")
    }
} 
Run, %sLink% 
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------

Teams_OpenChat(sSelection := "") {
; Use Selection or Outlook Object

If WinActive("ahk_exe Outlook.exe") {
    olApp := ComObjActive("Outlook.Application")
    oItem := Outlook_GetCurrentItem(olApp)
    
    Switch Outlook_Item2Type(oItem) {
        Case "Mail":
            sTopicName := "[Email] " . oItem.Subject
        Case "Meeting": ; From Inbox Meeting Request
            oItem := oItem.GetAssociatedAppointment(False)
            sTopicName := "[Meeting] " . oItem.Subject
        Case "Appointment":
            sTopicName := "[Meeting] " . oItem.Subject
    }
    sEmailList := Outlook_Item2Emails(oItem,,True)
    If (sEmailList=="") {
        return
    }
    ; TODO Remove own email address

    Teams_Emails2Chat(sEmailList,sTopicName)

} Else {
    Teams_Selection2Chat(sSelection)
}

} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Teams_Selection2Chat(sSelection:=""){
; Selection to Chat
; Called from Launcher 2c command
If (sSelection == "")
    sSelection := People_GetSelection()
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("Teams:Emails2Chat warning!","No email could be found in current selection!")   
    return
}
Teams_Emails2Chat(sEmailList)
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Emails2Meeting(sEmailList){
; Create and Open Teams Meeting from list of Emails
sLink := "https://teams.microsoft.com/l/meeting/new?attendees=" . StrReplace(sEmailList, ";",",") ; msteams:
TeamsExe = Teams_GetExe()
If FileExist(TeamsExe)
    sLink := StrReplace(sLink,"https://teams.microsoft.com","msteams:")
Run, %sLink% 
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Link2Fav(sUrl:="",FavsDir:="",sFileName :="") {
; Called by Email2TeamsFavs
; Create Shortcut file
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2021/03/teams-shortcuts-favorites.html"
	return
}

If (sUrl="") {
    sUrl := Clipboard
}

If (sUrl="") or !(sUrl ~= "https://teams.microsoft.com/*") {
	
    ; TODO Def Team name
    If RegExMatch(sLink,"https://teams.microsoft.com/l/channel/[^/]*/([^/]*)\?.*",sChannelName) 
	    linktext = %sChannelName1% (Channel)
    Else
        linktext = Team Name (Team)

    InputBox, sUrl , Teams Fav Target, Paste Teams Link:, , 640, 125,,,,, %linktext%
	If ErrorLevel
		return
}

; folder does not end with filesep
If !(FavsDir) {
    sKeyName := "TeamsFavsDir"
    RegRead, StartingFolder, HKEY_CURRENT_USER\Software\PowerTools, %sKeyName%
    FileSelectFolder, FavsDir , *%StartingFolder%, ,Select folder to store your Teams Shortcut:
    If ErrorLevel
        return
}

If !(sFileName) {
    sFileName := Teams_Link2Text(sUrl)
    InputBox, sFileName , Teams Fav File name, Enter the File name:, , 300, 125,,,,, %sFileName%
    If ErrorLevel
        return
}

; open by default in app: does not work for chat link
If Not InStr(sUrl,"?ctx=chat")
    sUrl:= StrReplace(sUrl,"https:","msteams:") 
sFile := FavsDir . "\" . sFileName . ".url"

;FileDelete %sFile%
IniWrite, %sUrl%, %sFile%, InternetShortcut, URL

; Add icon file:
TeamsExe := Teams_GetExe()
IniWrite, %TeamsExe%, %sFile%, InternetShortcut, IconFile
IniWrite, 0, %sFile%, InternetShortcut, IconIndex

; Save FavsDir to Settings in registry
SplitPath, sFile, sFileName, FavsDir
PowerTools_RegWrite("TeamsFavsDir",FavsDir)
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_FavsSetDir(){
RegRead, StartingFolder, HKEY_CURRENT_USER\Software\PowerTools, TeamsFavsDir
FileSelectFolder, sKeyValue , StartingFolder, Options, Select folder for your Teams Favorites:
If ErrorLevel
    return
PowerTools_RegWrite("TeamsFavsDir",sKeyValue)
return sKeyValue
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_FavsOpenDir(){
RegRead, FavDir, HKEY_CURRENT_USER\Software\PowerTools, TeamsFavsDir
If (FavDir ="") {
    Teams_FavsSetDir()
    return
}
Run, "%FavDir%"
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_FavsOpen(sInput) {
If (sInput="") {
    Teams_FavsOpenDir()
    Return
}
; List .lnk files in FavDir
RegRead, FavDir, HKEY_CURRENT_USER\Software\PowerTools, TeamsFavsDir
If (FavDir ="") {
    TrayTipAutoHide("Teamsy: OpenFav!","No Favorites directory!") 
    return
}
sInput := Trim(sInput)
sPat := StrReplace(sInput," ","[^\s]*\s")

Loop, Files, %FavDir%\*.url, R ; recurse in subdirectories
{
    If RegExMatch(SubStr(A_LoopFileName, 1,-4),"i)" . sPat) { ; LoopFileName contains extension; ignore case
        Run, %A_LoopFileFullPath%
        return
    }
} ; end loop
Loop, Files, %FavDir%\*.lnk, R ; recurse in subdirectories
{
    If RegExMatch(SubStr(A_LoopFileName, 1,-4),"i)" . sPat) { ; LoopFileName contains extension; ignore case
        Run, %A_LoopFileFullPath%
        return
    }
} ; end loop
sMsg := "No Favorite found matching '", . sPat . "'."
TrayTipAutoHide("Teamsy: OpenFav!",sMsg) 
    

} ; eofun


; -------------------------------------------------------------------------------------------------------------------
Teams_FavsAdd(){
; Add to Favorites: either link or Email List
; Called from Launcher f+ command
sInput := Clip_GetSelection()
If (sInput="") {
    sInput := Clipboard
}
If RegExMatch(sInput,"^http.*"){ ; Link
    Teams_Link2Fav(sInput)
} Else {
    Teams_Emails2Favs(sInput)
}
} ; eofun

; ----------------------------------------------------------------------
Teams_Emails2Favs(sInput:= ""){
; Calls: Email2TeamsFavs

If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2021/03/teams-people-favorites.html"
	return
}
If (sInput ="") {
    sInput := People_GetSelection()
    If (sInput = "") { 
        TrayTipAutoHide("Teams: Emails to Favs!","Nothing selected or Clipboard empty!")   
        return
    }
}
sEmailList := People_GetEmailList(sInput)
If (sEmailList = "") { 
    TrayTipAutoHide("Teams: Emails to Favs!","No email could be found from Selection or Clipboard!")   
    return
}
RegRead, FavsDir, HKEY_CURRENT_USER\Software\PowerTools, TeamsFavsDir
If ErrorLevel {
    FavDirs := Teams_FavsSetDir()
    If FavDirs = ""
        return    
}
FileSelectFolder, FavsDir , *%FavsDir%, ,Select folder to store your Teams Contact Shortcuts:
If ErrorLevel
    return
Loop, parse, sEmailList, ";"
{
    Email2TeamsFavs(A_LoopField,FavsDir)
}
Run %FavsDir%	; open Favorites directory
} ; eofun
; ----------------------------------------------------------------------

Email2TeamsFavs(sEmail,FavsDir){
; Calls: Teams_Link2Fav, Teams_FavsAdd

; Get Firstname
sName := RegExReplace(sEmail,"\..*" ,"")
StringUpper, sName, sName , T

; 1. Create Chat Shortcut
sUrl = https://teams.microsoft.com/l/chat/0/0?users=%sEmail% 
Teams_Link2Fav(sUrl,FavsDir,"Chat " + sName)

; 2. Create Call shortcut
sFile := FavsDir . "\Call " . sName . ".vbs"
; write code
TeamsExe := Teams_GetExe()
sText = CreateObject("Wscript.Shell").Run "%TeamsExe% callto:%sEmail%"

; Create empty File
FileDelete %sFile%
FileAppend , %sText%, %sFile%

; create shortcut
sLnk := PathX(sFile, "Ext:.lnk").Full
FileCreateShortcut, %sFile%, %sLnk%,,,, %TeamsExe%

} ; eofun

; NOT USED ########################
Emails2TeamsFavGroupChat(sEmailList){
RegRead, FavsDir, HKEY_CURRENT_USER\Software\PowerTools, TeamsFavsDir
If ErrorLevel {
    FavDirs := Teams_FavsSetDir()
    If FavDirs = ""
        return    
}

FileSelectFolder, FavsDir , *%StartingFolder%, ,Select folder to store your Teams Shortcut:
If ErrorLevel
    return

sUrl := "https://teams.microsoft.com/l/chat/0/0?users=" . StrReplace(sEmailList, ";",",")

InputBox, sName, Enter Chat Group name,,,,100
if ErrorLevel
    return

sName := StrReplace(sName, ":", "")
sLink := sLink . "&topicName=" . sName
 
Teams_Link2Fav(sUrl,FavsDir,"Group Chat -" . sName)
} ; eofun
; ----------------------------------------------------------------------

Teams_GetExe(){
fExe = C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\current\Teams.exe
return fExe
} ;eofun
; ----------------------------------------------------------------------

; ----------------------------------------------------------------------
; ----------------------------------------------------------------------
Teams_Selection2Team() {
; Calls: Teams_Emails2Team
sSelection := People_GetSelection()
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("TeamsShortcuts warning!","No email could be found in selection or clipboard!")   
    return
}
Teams_Emails2Team(sEmailList)

} ; eofun

; ----------------------------------------------------------------------

Teams_Emails2Team(EmailList,TeamLink:=""){
; Syntax: 
;     Teams_Emails2Team(EmailList,TeamLink*)
; EmailList: String of email adresses separated with a ;
; TeamLink: (String) optional. If not passed, user will be asked via inputbox
;  e.g. https://teams.microsoft.com/l/team/19%3a12d90de31c6e44759ba622f50e3782fe%40thread.skype/conversations?groupId=640b2f00-7b35-41b2-9e32-5ce9f5fcbd01&tenantId=xxx

; Requires PowerShell
; Calls: Teams_PowerShellCheck
If !Teams_PowerShellCheck()
    return

If (TeamLink=="") {
    InputBox, TeamLink , Team Link, Enter Team Link:,,640,125 
    if ErrorLevel
	    return
}
; 
sPat = \?groupId=(.*)&tenantId=(.*)
If !(RegExMatch(TeamLink,sPat,sId)) {
	MsgBox 0x10, Error, Provided Url does not match a Team Link!
	return
}
sGroupId := sId1
sTenantId := sId2

; Create csv file with list of emails
CsvFile = %A_Temp%\email_list.csv
If FileExist(CsvFile)
    FileDelete, %CsvFile%
sText := StrReplace(EmailList,";","`n")
FileAppend, %sText%,%CsvFile% 
; Create .ps1 file
PsFile = %A_Temp%\Teams_AddUsers.ps1
; Fill the file with commands
If FileExist(PsFile)
    FileDelete, %PsFile%

Domain := People_GetDomain()
If (Domain ="") {
    MsgBox 0x10, Teams Shortcuts: Error, No Domain defined!
    return
}
OfficeUid := People_GetMyOUid()

sText = Connect-MicrosoftTeams -TenantId %sTenantId% -AccountId %OfficeUid%@%Domain%
sText = %sText%`nImport-Csv -header email -Path "%CsvFile%" | foreach{Add-TeamUser -GroupId "%sGroupId%" -user $_.email}
FileAppend, %sText%,%PsFile%

; Run it
RunWait, PowerShell.exe -NoExit -ExecutionPolicy Bypass -Command %PsFile% ;,, Hide
;RunWait, PowerShell.exe -ExecutionPolicy Bypass -Command %PsFile% ,, Hide
}

; ----------------------------------------------------------------------
Teams_ExportTeams() {
; CsvFile := Teams_ExportTeams
; returns empty if not created/ failed

; Requires PowerShell

If GetKeyState("Ctrl") {
	Teamsy_Help("t2xl")
	return
}

If Not Teams_PowerShellCheck()
    return
CsvFile = %A_ScriptDir%\Teams_list.csv
If FileExist(CsvFile)
    FileDelete, %CsvFile%
; Create .ps1 file
PsFile = %A_Temp%\Teams_ExportTeams.ps1
; Fill the file with commands
If FileExist(PsFile)
    FileDelete, %PsFile%

Domain := People_GetDomain()
If (Domain ="") {
    MsgBox 0x10, Teams Shortcuts: Error, No Domain defined!
    return
}

OfficeUid := People_GetMyOUid()
sText = Connect-MicrosoftTeams -AccountId %OfficeUid%@%Domain%
sText = %sText%`nGet-Team -User %OfficeUid%@%Domain% | Select DisplayName, MailNickName, GroupId, Description | Export-Csv -Path %CsvFile% -NoTypeInformation
FileAppend, %sText%,%PsFile%

; Run it;RunWait, PowerShell.exe -NoExit -ExecutionPolicy Bypass -Command %PsFile% ;,, Hide
RunWait, PowerShell.exe -ExecutionPolicy Bypass -Command %PsFile% ,, Hide
return CsvFile
}

; ----------------------------------------------------------------------
Teams_PowerShellCheck() {
; returns True if Teams PowerShell is setup, False else
RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If (TeamsPowerShell := !TeamsPowerShell) {
    sUrl := "https://tdalon.blogspot.com/2020/08/teams-powershell-setup.html"
    Run, "%sUrl%" 
    MsgBox 0x1024,People Connector, Have you setup Teams PowerShell on your PC?
	IfMsgBox No
		return False
    OfficeUid := People_GetMyOUid()

    PowerTools_RegWrite("TeamsPowerShell",TeamsPowerShell)
    return True
} Else ; was already set
    return True 

} ; eofun


; ----------------------------------------------------------------------
Teams_GetTeamName(sInput) {
; TeamName := Teams_GetTeamName(GroupId)
; TeamName := Teams_GetTeamName(SharePointUrl)
; TeamName := Teams_GetTeamName(TeamUrl)
CsvFile = %A_ScriptDir%\Teams_list.csv
If !FileExist(CsvFile) {
    Teams_ExportTeams()
}
If RegExMatch(sInput,"/teams/(team_[^/]*)/",sMatch) {
    MailNickName := sMatch1
    TeamName := ReadCsv(CsvFile,"MailNickName",MailNickName,"DisplayName")
    return TeamName
} Else If RegExMatch(sInput,"\?groupId=([^&]*)",sMatch) 
    GroupId := sMatch1
Else
    GroupId := sInput


TeamName := ReadCsv(CsvFile,"GroupId",GroupId,"DisplayName")
return TeamName

} ; end of function
; ----------------------------------------------------------------------
TeamsLink2TeamName(TeamLink){
; Obsolete. Replaced by Teams_GetTeamName
; TeamName := TeamsLink2TeamName(TeamLink)
; Called by Teams_Link2Text
RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If !TeamsPowerShell {
    InputBox, TeamName , Team Link, Enter Team Link:,,640,125 
    if ErrorLevel
	    return
    return TeamName
}

sPat = \?groupId=(.*)&tenantId=(.*)
If !(RegExMatch(TeamLink,sPat,sId)) {
	MsgBox 0x10, Error, Provided Url does not match a Team Link!
	return
}
sGroupId := sId1
sTenantId := sId2

; Create .ps1 file
PsFile = %A_Temp%\Teams_AddUsers.ps1
; Fill the file with commands
If FileExist(PsFile)
    FileDelete, %PsFile%


Domain := People_GetDomain()
If (Domain ="") {
    MsgBox 0x10, Teams Shortcuts: Error, No Domain defined!
    return
}
OfficeUid := People_GetMyOUid()
sText = Connect-MicrosoftTeams -TenantId %sTenantId% -AccountId %OfficeUid%@%Domain%
sText = %sText%`nGet-Team -GroupId %sGroupId%
FileAppend, %sText%,%PsFile%

; Get a temporary file path
tempFile := A_Temp "\PowerShell_Output.txt"                           ; "

; Run the console program hidden, redirecting its output to
; the temp. file (with a program other than powershell.exe or cmd.exe,
; prepend %ComSpec% /c; use 2> to redirect error output), and wait for it to exit.
RunWait, powershell.exe -ExecutionPolicy Bypass -Command %PsFile% > %tempFile%,, Hide

; Read the temp file into a variable and then delete it.
FileRead, sOutput, %tempFile%
FileDelete, %tempFile%

; Display the result.
If RegExMatch(sOutput,"DisplayName.*:(.*)",sMatch)
    TeamName := sMatch1
return TeamName
}

; ----------------------------------------------------------------------
SPLink2TeamName(TeamLink){
; Called by Teams_Link2Text
RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If !TeamsPowerShell {
    InputBox, TeamName , Team Link, Enter Team Link:,,640,125 
    if ErrorLevel
	    return
    return TeamName
}

If (!TeamLink) {
    InputBox, TeamLink , SharePoint Link, Enter Team SharePoint Link:,,640,125 
    if ErrorLevel
	    return
}

; Create .ps1 file
PsFile = %A_Temp%\Teams_AddUsers.ps1
; Fill the file with commands
If FileExist(PsFile)
    FileDelete, %PsFile%

Domain := People_GetDomain()
If (Domain ="") {
    MsgBox 0x10, Teams Shortcuts: Error, No Domain defined!
    return
}
OfficeUid := People_GetMyOUid()
sText = Connect-MicrosoftTeams -AccountId %OfficeUid%@%Domain%
sText = %sText%`nGet-Team -GroupId %sGroupId%
FileAppend, %sText%,%PsFile%

; Get a temporary file path
tempFile := A_Temp "\PowerShell_Output.txt"                           ; "

; Run the console program hidden, redirecting its output to
; the temp. file (with a program other than powershell.exe or cmd.exe,
; prepend %ComSpec% /c; use 2> to redirect error output), and wait for it to exit.
RunWait, powershell.exe -ExecutionPolicy Bypass -Command %PsFile% > %tempFile%,, Hide

; Read the temp file into a variable and then delete it.
FileRead, sOutput, %tempFile%
FileDelete, %tempFile%

; Display the result.
If RegExMatch(sOutput,"DisplayName.*:(.*)",sMatch)
    sName := sMatch1
return sName
}

; ----------------------------------------------------------------------
TeamLink2Users(TeamLink){
If (!TeamLink) {
    InputBox, TeamLink , Team Link, Enter Team Link:,,640,125 
    if ErrorLevel
	    return
}

sPat = \?groupId=(.*)&tenantId=(.*)
If !(RegExMatch(TeamLink,sPat,sId)) {
	MsgBox 0x10, Error, Provided Url does not match a Team Link!
	return
}
sGroupId := sId1
sTenantId := sId2

; Create .ps1 file
PsFile = %A_Temp%\Teams_GetUser.ps1
; Fill the file with commands
If FileExist(PsFile)
    FileDelete, %PsFile%

Domain := People_GetDomain()
If (Domain ="") {
    MsgBox 0x10, Teams Shortcuts: Error, No Domain defined!
    return
}
OfficeUid := People_GetMyOUid()
sText = Connect-MicrosoftTeams -TenantId %sTenantId% -AccountId %OfficeUid%@%Domain%
sText = %sText%`nGet-TeamUser -GroupId %sGroupId%
FileAppend, %sText%,%PsFile%

; Get a temporary file path
tempFile := A_Temp "\PowerShell_Output.txt"                           ; "

; Run the console program hidden, redirecting its output to
; the temp. file (with a program other than powershell.exe or cmd.exe,
; prepend %ComSpec% /c; use 2> to redirect error output), and wait for it to exit.
RunWait, powershell.exe -ExecutionPolicy Bypass -Command %PsFile% > %tempFile%,, Hide

; Read the temp file into a variable and then delete it.
FileRead, sOutput, %tempFile%
;FileDelete, %tempFile%
Run %tempFile% 
}

; -------------------------------------------------------------------------------------------------------------------

Teams_SmartReply(doReply:=True){
If GetKeyState("Ctrl") {
    Run, "https://tdalon.blogspot.com/2020/11/teams-shortcuts-smart-reply.html"
	return
}

savedClipboard := ClipboardAll
sSelectionHtml := Clip_GetSelectionHtml(False)

If (sSelectionHtml="") { ; no selection -> default reply
    Send !+r ; Alt+Shift+r
    Clipboard := savedClipboard
    return
}

If InStr(sSelectionHtml,"data-tid=""messageBodyContent""") { ; Full thread selection e.g. Shift+Up
    ; Remove Edited block

    sSelectionHtml := StrReplace(sSelectionHtml,"<div>Edited</div>","")
    ;MsgBox % sSelectionHtml ; DBG
    ; Get Quoted Html
    ;sPat = Us)<dd><div>.*<div data-tid="messageBodyContainer">(.*)</div></div></dd>
    sPat = sU)<div data-tid="messageBodyContent"><div>\n?<div>\n?<div>(.*)</div>\n?</div>\n?</div> ; with line breaks
    If RegExMatch(sSelectionHtml,sPat,sMatch)
        sQuoteBodyHtml := sMatch1 
    Else { ; fallback
        sQuoteBodyHtml := Clip_GetSelection(False) ; removes formatting
    }
    
    sHtmlThread := sSelectionHtml
    ;MsgBox %sQuoteBodyHtml% ; DBG
} Else { ; partial thread selection

    sQuoteBodyHtml := Clip_GetSelection(False) ; removes formatting
    If (!sQuoteBodyHtml) {
        MsgBox Selection empty!
        Clipboard := savedClipboard
        return
    }

    SendInput +{Up} ; Shift + Up Arrow: select all thread 
    Sleep 200
    sHtmlThread := Clip_GetSelectionHtml(False)
}

; Extract Teams Link

; Full thread selection to get link and author
; &lt;https://teams.microsoft.com/l/message/19:b4371b8d10234ac9b4e0095ace0aae8e@thread.skype/1600430385590?tenantId=xxx&amp;amp;groupId=56bc81d8-db27-487c-8e4f-8d5ea9058663&amp;amp;parentMessageId=1600430385590&amp;amp;teamName=GUIDEs&amp;amp;channelName=Best Work Hacks&amp;amp;createdTime=1600430385590&gt;</div></div><!--EndFragment-->

; Get thread author and conversation title from parent thread
; <span>Dalon, Thierry</span><div>Quick Share link to Teams</div><div data-tid="messageBodyContainer">
sPat = U)<span>([^>]*)</span><div>(.*)</div><div data-tid="messageBodyContainer">
If (RegExMatch(sHtmlThread,sPat,sMatch)) {
    sAuthor := sMatch1
    sTitle := sMatch2
} Else { ; Sub-thread
; <span>Schmidt, Thomas</span><div data-tid="messageBodyContainer">
    sPat = U)<span>([^/]*)</span><div data-tid="messageBodyContainer">
    If (RegExMatch(sHtmlThread,sPat,sMatch)) {
        sAuthor := sMatch1
    }
}

If (sAuthor = "") { ; something went wrong
; TODO error handling
    MsgBox Something went wrong! Please retry.
    ;MsgBox %sHtmlThread% ; DBG
    Clipboard := savedClipboard
    return
}

; Get thread link
sPat := "U)<div>&lt;https://teams\.microsoft\.com/(.*;createdTime=.*)&gt;</div>.*</div><!--EndFragment-->"
If (RegExMatch(sHtmlThread,sPat,sMatch)) {
    sMsgLink = https://teams.microsoft.com/%sMatch1%
    sMsgLink := StrReplace(sMsgLink,"&amp;amp;","&")
}

If (doReply = True) { ; hotkey is buggy - will reply to original/ quoted thread-> need to click on reply manually in case of quote from another thread
    If (sMsgLink = "") ; group chat
        Send !+c ; Alt+Shift+c
    Else
        Send !+r ; Alt+Shift+r
    Sleep 1500
} Else { ; ask for continue
    Answ := ButtonBox("Paste conversation quote","Activate the place where to paste the quoted conversation and hit Continue`nor select Create New...","Continue|Create New...")
    If (Answ="Button_Cancel") {
        ; Restore Clipboard
        Clipboard := savedClipboard
        Return
    }
}

If (People_IsMe(sAuthor)) 
    sAuthor = I 

TeamsReplyWithMention := False
If WinActive("ahk_exe Teams.exe") { ; SendMention
    ; Mention Author
    If (sAuthor <> "I") {
        TeamsReplyWithMention := True
        SendInput > ; markdown for quote block
        Sleep 500  
        Teams_SendMention(sAuthor)
        sAuthor =
    }
}

If (sMsgLink = "") ; group chat
    sQuoteTitleHtml = %sAuthor% wrote:
Else
    sQuoteTitleHtml = <a href="%sMsgLink%">%sAuthor%&nbsp;wrote</a>:    

If (TeamsReplyWithMention = True)
    sQuoteHtml = %sQuoteTitleHtml%<br>%sQuoteBodyHtml%
Else
    sQuoteHtml = <blockquote>%sQuoteTitleHtml%<br>%sQuoteBodyHtml%</blockquote>

;MsgBox %sQuoteHtml% ; DBG
Clip_PasteHtml(sQuoteHtml,sQuoteBodyHtml,False)

; Escape Block quote in Teams: twice Shift+Enter
If WinActive("ahk_exe Teams.exe") {
    ;SendInput {Delete} ; remove empty div to avoid breaking quote block in two
    SendInput +{Enter} 
    SendInput +{Enter}
}

; Restore Clipboard
Clipboard := savedClipboard
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_ConversationReaction2(reaction,WinId:=""){
; Does not work / can not find menu elements
; reaction can be: Like, Heart, Laugh, Surprised, Sad, Angry
UIA := UIA_Interface()
If !WinId
    WinId := WinActive("A")

Send {Enter} ; to activate the elements

TeamsEl := UIA.ElementFromHandle(WinId) 
El:= TeamsEl.FindFirstByNameAndType(reaction,"menu")
MSgBox % El.Dump()
El.Click()
} ; eofun


; -------------------------------------------------------------------------------------------------------------------
Teams_ConversationReaction(reaction){
; reaction can be: Like, Heart, Laugh, Surprised, Sad, Angry
; hard-coded menu order

StringLower, reaction, reaction
;MsgBox % reaction
Send {Click}
Send {Enter} ; to activate the menu elements

; UIA does not work to find element

Switch reaction {
    Case "like":
    Case "heart":
        Send {Right}
    Case "laugh":
        Send {Right 2}
    Case "surprised":
        Send {Right 3}
    Case "sad":
        Send {Right 4}
    Case "angry":
        Send {Right 5}
}
Send {Enter}
} ; eofun


; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
Teams_ConversationAction(action){
; action: save|copylink|unread|sharetooutlook|delete|edit
; hard-coded menu order arrow down

StringLower, action, action
MouseClick, Left
Send {Enter} 
Send {Tab}{Enter} ; open action menu

Switch action {
    Case "save":
        Send {Alt}
    Case "edit":
        Send {Down}
    Case "delete":
        Send {Down 2}
    Case "unread":
        Send {Down 3}
    Case "copylink":
        Send {Down 4}
    Case "sharetooutlook":
        Send {Down 5}
}
Send {Enter}


} ; eofun
; -------------------------------------------------------------------------------------------------------------------


; -------------------------------------------------------------------------------------------------------------------
Teams_GetLink(){
SendInput +{Up} ; Shift + Up Arrow: select all thread
Sleep 200
sSelection := Clip_GetSelectionHtml()
SendInput {Esc} ; unselect thread
; Parse <a href=""
sPat = <a href="([^"]*)"
sep := "|"
Pos = 1 
While Pos := RegExMatch(sSelection,sPat,sMatch,Pos+StrLen(sMatch)){
    If InStr(sList . sep,sMatch1 . sep) ; skip duplicates
        continue
    sList := sList . sep . sMatch1
}
sList := SubStr(sList,2) ; Remove first sep
; Check if multiple links-> select box
If InStr(sList,sep) {
    Result := ListBox("Teams:Get Link","Choose link",sList,1)
    return Result
} 
return sList

} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Teams_ConversationGetLink(){
; Alternative to select all Thread: Double MouseClick, Left
SendInput +{Up} ; Shift + Up Arrow: select all thread
Sleep 200
sSelection := Clip_GetSelection()
; Last part between <> is the thread link
RegExMatch(sSelection,"U).*<(.*)>$",sMatch)
SendInput {Esc} ; unselect thread
sTeamLink := StrReplace(sMatch1,"&amp;amp;","&")
return sTeamLink
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Teams_SendMention(sInput, doPerso := ""){
; See notes https://tdalon.blogspot.com/teams-shortcuts-send-mention
If (doPerso = "")
    doPerso := PowerTools_RegRead("TeamsMentionPersonalize")

If InStr(sInput,"@") { ; Email
    sName := People_ADGetUserField("mail=" . sInput, "DisplayName")
    sInput := RegExReplace(sName, "\s\(.*\)") ; remove (uid) else mention completion does not work 
} Else If InStr(sInput,",") {
    sName := sInput
    sInput := RegExReplace(sName, "\s\(.*\)") 
} 
SendInput {@}
Sleep 300
SendInput %sInput%
Delay := PowerTools_GetParam("TeamsMentionDelay")
Sleep %Delay% 
SendInput +{Tab} ; use Shift+Tab because Tab will switch focus to next UI element in case no mention autocompletion can be done (e.g. name not member of Team)
; Personalize mention -> Firstname
If  (doPerso=True) {
    Teams_PersonalizeMention()
}
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Teams_PersonalizeMention() {
; Works for any Mention with format Lastname, Firstname xxx
; xxx can be (company) or company multiple words
; Runs right after mention autocomplete in Teams
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2020/11/teams-shortcuts-personalize-mentions.html"
	return
}

ClipBackup:=  ClipboardAll

; Check for (company name) format
SendInput ^+{Left} ; Ctrl+Shift+Left
sSelection := Clip_GetSelection(false) 
If (InStr(sSelection,")",0))  { ; last letter is )
    ; loop till (
    while (!InStr(sSelection,"(",1)) {
        SendInput ^+{Left} ; Ctrl+Shift+Left
        sSelection := Clip_GetSelection(false) 
    }

    ; TODO Check for (guest)
    If (InStr(sSelection,"(guest)")) {
        ; delete (guest)
        SendInput {Backspace}
        SendInput ^+{Left} ; Ctrl+Shift+Left
        sSelection := Clip_GetSelection(false) 
        ; loop till (
        while (!InStr(sSelection,"(",1)) {
            SendInput ^+{Left} ; Ctrl+Shift+Left
            sSelection := Clip_GetSelection(false) 
        }
    }

    ; delete (company name)
    SendInput {Backspace}

    ; two cases Firstname Lastname or Lastname, Firstname
    SendInput ^+{Left 2} ; Ctrl+Shift+Left

    Sleep 10
    sSelection := Clip_GetSelection(false) 
    
    
    If (InStr(sSelection,",")) {
        SendInput ^+{Right} ; Ctrl+Right/Right does not always work. Sometimes jump to the end of the name Lastname, Firstname and vice-versa -> use Shift
        
        SendInput {Left} ; remove selection
        Sleep 10
        SendInput {Backspace} 
        SendInput {delete} ; delete extra space
        ; restore cursor to the right
        SendInput ^{Right} 
    } Else { ; Assume Firstname Lastname
        SendInput ^{Right}
        Sleep 10
        SendInput {Backspace} 
    }


    Clipboard := ClipBackup
    return ; exit/ end format with (company name)
}

; For format without (company name)

; Look for , between Lastname, Firstname
; select till , is found 0; navigating wiht Ctrl pressed jump to nex words/ skip spaces
while (!InStr(sSelection,",") and A_Index < 10) {
    SendInput ^+{Left} ; Ctrl+Shift+Left
    oldLen := StrLen(sSelection)
    Sleep 10
    sSelection := Clip_GetSelection(false) 
    nwMention := A_Index + 1 ; number of words in Mention
    If (StrLen(sSelection) = oldLen) { ; beginning of line reached
        break
    }
}

If (!InStr(sSelection,",")) { ; no , found
    MsgBox % sSelection
    sSelection := Clip_GetSelectionHtml(false) 
}

; Set cursor back to end of mention
SendInput {Right}

; Delete post firstname: from right to nwMention-2
nCount := nwMention -1

Loop, %nCount% 
    SendInput ^+{Left}

return
Sleep 10 ; weird pause required
SendInput {Backspace}

; Delete Lastname
SendInput ^{Left} ; skip firstname
SendInput ^+{Left}{Backspace} 

; reposition cursor after firstname-only Mention
SendInput ^{Right} 


Clipboard := ClipBackup
return

; NOT USED
; handling - in names
SendInput +{Left} 
sLastLetter := Clip_GetSelection(false)   
If (sLastLetter = "-")
    SendInput ^{Left}{Backspace}
Else 
    SendInput {Backspace}




} ; eofun




; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
Teams_Selection2Mentions(sSelection:=""){
; Add to Favorites: either link or Email List
; Called from Launcher e2m command
If (sSelection == "")
    sSelection := People_GetSelection()
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("Teams:Emails2Chat warning!","No email could be found in current selection!")   
    return
}
Teams_Emails2Mentions(sEmailList)
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Emails2Mentions(sEmailList,doPerso :=""){
If (doPerso = "")
    doPerso := PowerTools_RegRead("TeamsMentionPersonalize")
MyEmail := People_GetMyEmail()
Loop, parse, sEmailList, ";"
{
	If (A_LoopField=MyEmail) ; Skip my email
        continue
    Teams_SendMention(A_LoopField,doPerso)
	SendInput {,}{Space} 
}	; End Loop 
SendInput {Backspace}{Backspace}{Space} ; remove final ,
} ; eofun
; -------------------------------------------------------------------------------------------------------------------



Teams_Link2Text(sLink){
sLink := StrReplace(sLink,"%2520"," ") ; spaces in Channel Link
sPat = [https|msteams]://teams.microsoft.com/[^>"]*
RegExMatch(sLink,sPat,sLink)
sLink := uriDecode(sLink)
; Link to Teams Channel
; example: https://teams.microsoft.com/l/channel/19%3a16ff462071114e31bd696aa3a4e34500%40thread.skype/DOORS%2520Attributes%2520List?groupId=cd211b48-2e8b-4b60-b5b0-e584a0cf30c0&tenantId=xxx
If (RegExMatch(sLink,"U)[https|msteams]://teams\.microsoft\.com/l/channel/[^/]*/([^/]*)\?groupId=(.*)&",sMatch)) {
    sChannelName := sMatch1
    sTeamName := Teams_GetTeamName(sMatch2)
    If (!sTeamsName)
        sDefText = %sTeamName% Team
    Else
        sDefText = Link to Teams Team
	linktext = %sDefText% - %sChannelName% (Channel)
	;linktext := StrReplace(linktext,"%2520"," ")		
; Link to Teams
; example: https://teams.microsoft.com/l/team/19%3a12d90de31c6e44759ba622f50e3782fe%40thread.skype/conversations?groupId=640b2f00-7b35-41b2-9e32-5ce9f5fcbd01&tenantId=xxx
} Else If (RegExMatch(sLink,"[https|msteams]://teams\.microsoft\.com/l/team/.*groupId=(.*)&",sMatch)) {
; https://teams.microsoft.com/l/team/19:c1471a18bae04cf692b8da7e9738df3e@thread.skype/conversations?groupId=56bc81d8-db27-487c-8e4f-8d5ea9058663&tenantId=xxx    
    sTeamName := Teams_GetTeamName(sMatch1)
    If (!sTeamsName)
        sDefText = %sTeamName% Team
    Else
        sDefText = Link to Teams Team
    InputBox, linktext , Display Link Text, Enter Team name:,,640,125,,,,, %sDefText%
	If ErrorLevel ; Cancel
		return
}
return linktext  
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_FileLinkBeautifier(sLink){
; sLink := Teams_FileLinkBeautifier(sLink)
sLink := uriDecode(sLink)

If RegExMatch(sLink,"https://teams\.microsoft\.com(?:.*)/l/file/.*&objectUrl=(.*?)&",sMatch){
    return sMatch1
}
If RegExMatch(sLink,"https://teams\.microsoft\.com(?:.*)/files/.*&rootfolder=(.*)",sMatch){
    TenantName := PowerTools_GetSetting("TenantName")
	If (TenantName="") {
		return sLink 
	}
    sLink = https://%TenantName%.sharepoint.com%sMatch1% 
    return sLink
}
return sLink
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Teams_MessageLink2Html(sLink){
sPat = https://teams.microsoft.com/[^>"]*
RegExMatch(sLink,sPat,sLink)
sLink := uriDecode(sLink)
RegExMatch(sLink,"https://teams.microsoft.com/l/message/(.*)\?tenantId=(.*)&groupId=(.*)&teamName=(.*)&channelName=(.*)&",sMatch) 
; https://teams.microsoft.com/l/message/19:fbfd482e3b544af387cbe6db65846796@thread.skype/1582627348656?tenantId=xxx&amp;groupId=56bc81d8-db27-487c-8e4f-8d5ea9058663&amp;parentMessageId=1582627348656&amp;teamName=GUIDEs&amp;channelName=Technical Questions&amp;createdTime=1582627348656  

; Prompt for Type of link: Team, Channel or Conversation
Choice := ButtonBox("Teams Link:Setting","Do you want a link to the:","Message|Channel|Team")
If ( Choice = "ButtonBox_Cancel") or ( Choice = "Timeout")
    return
Switch Choice
{
Case "Team":
    linktext = %sMatch4% (Team)
    sLink = https://teams.microsoft.com/l/team/%sMatch1%/conversations?groupId=%sMatch3%&tenantId=%sMatch2%
Case "Channel":
    linktext = %sMatch4% - %sMatch5% (Channel)
    sLink = https://teams.microsoft.com/l/channel/%sMatch1%/%sMatch5%?groupId=%sMatch3%&tenantId=%sMatch2%
    ; https://teams.microsoft.com/l/channel/19%3ab4371b8d10234ac9b4e0095ace0aae8e%40thread.skype/Best%2520Work%2520Hacks?groupId=56bc81d8-db27-487c-8e4f-8d5ea9058663&tenantId=xxx
Case "Message":
    linktext = %sMatch4% - %sMatch5% - Message
}
sHtml = <a href="%sLink%">%linktext%</a>
return sHtml
}

; -------------------------------------------------------------------------------------------------------------------
Teams_OpenSecondInstance(){
If GetKeyState("Ctrl") {
	Teamsy_Help("2")
	return
}
EnvGet, A_UserProfile, userprofile
wd = %A_UserProfile%\AppData\Local\Microsoft\Teams
sCmd = Update.exe --processStart "Teams.exe"

EnvGet, A_LocAppData, localappdata
up = %A_LocAppData%\Microsoft\Teams\CustomProfiles\Second
EnvSet, userprofile,  %up%
; Run it
TrayTipAutoHide("Teams Shortcuts", "Teams Second Instance is started...")
Run, %sCmd%,%wd%
EnvSet, userprofile,%A_UserProfile% ; might leads to troubles if not reset for further run command
}

; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------

Teams_OpenWebApp(){
Run, https://teams.microsoft.com
}

Teams_OpenWebCal(){
Run, https://teams.microsoft.com/_#/calendarv2
}

; -------------------------------------------------------------------------------------------------------------------
Teams_Members2Excel(TeamLink:=""){
; Input can be sGroupId or Team link

Domain := People_GetDomain()
If (Domain ="") {
    MsgBox 0x10, Teams Shortcuts: Error, No Domain defined!
    return
}

RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If (TeamsPowerShell := !TeamsPowerShell) {
    sUrl := "https://tdalon.blogspot.com/2020/08/teams-powershell-setup.html"
    Run, "%sUrl%" 
    MsgBox 0x1024,Teams Shortcuts, Have you setup Teams PowerShell on your PC?
	IfMsgBox No
		return
    OfficeUid := People_GetMyOUid()
    ;sWinUid := People_ADGetUserField("mail=" . sEmail, "sAMAccountName")  ;- login uid
    PowerTools_RegWrite("TeamsPowerShell",TeamsPowerShell)
   ; Menu,SubMenuSettings,Check, Teams PowerShell
}

If (TeamLink=="") {
    InputBox, TeamLink , Team Link, Enter Team Link:,,640,125 
    if ErrorLevel
	    return
}

sPat = \?groupId=(.*)&tenantId=(.*)
If (RegExMatch(TeamLink,sPat,sId)) {
	; MsgBox 0x10, Error, Provided Url does not match a Team Link!
	sGroupId := sId1
} Else
    sGroupId := TeamLink

; Create csv file with list of emails
CsvFile = %A_Temp%\Team_users_list.csv
If FileExist(CsvFile)
    FileDelete, %CsvFile%
sText := StrReplace(EmailList,";","`n")
FileAppend, %sText%,%CsvFile% 
; Create .ps1 file
PsFile = %A_Temp%\Teams_GetUsers.ps1
; Fill the file with commands
If FileExist(PsFile)
    FileDelete, %PsFile%

OfficeUid := People_GetMyOUid()
sText = Connect-MicrosoftTeams -AccountId %OfficeUid%@%Domain%
sText = %sText%`nGet-TeamUser -GroupId %sGroupId% | Export-Csv -Path %CsvFile% -NoTypeInformation
; Columns: UserId, User, Name, Role
FileAppend, %sText%,%PsFile%

; Run it
;RunWait, PowerShell.exe -NoExit -ExecutionPolicy Bypass -Command %PsFile% ;,, Hide
RunWait, PowerShell.exe -ExecutionPolicy Bypass -Command %PsFile% ,, Hide

oExcel := ComObjCreate("Excel.Application") ;handle
oExcel.Workbooks.Add ;add a new workbook
oSheet := oExcel.ActiveSheet

; Loop on Csv File

oExcel.Visible := True ;by default excel sheets are invisible
oExcel.StatusBar := "Export to Excel..."

Loop, read, %CsvFile%
; Columns: UserId, User, Name, Role
{
    RowCount := A_Index
    Loop, parse, A_LoopReadLine, CSV
    {
        ColCount := A_Index
        ; MsgBox, Field number %A_Index% is %A_LoopField%.
        Switch ColCount {
        Case 1:
            continue ;  skip UserId
        Case 2: ; User
            oSheet.Range("B" . RowCount).Value := A_LoopField  
            OfficeUid := A_LoopField 
            ;OfficeUid := RegExReplace(OfficeUid,"@.*","") ; remove @domain
        Case 3: ; Name
            oSheet.Range("A" . RowCount).Value := A_LoopField
            Name := A_LoopField 
        Case 4: ; Role
            oSheet.Range("F" . RowCount).Value := A_LoopField  
        } ; end swicth

        If (RowCount == 1)
            oSheet.Range("C" . RowCount).Value := "email"
        Else
            oSheet.Range("C" . RowCount).Value := People_oUid2Email(OfficeUid)
        
        LastName := RegexReplace(Name,",.*","")
        FirstName := RegexReplace(Name,".*,","")
        FirstName := RegExReplace(FirstName," \(.*\)","") ; Remove (uid) in firstname

        If (RowCount == 1)
            oSheet.Range("D" . RowCount).Value := "FirstName"
        Else
            oSheet.Range("D" . RowCount).Value := FirstName

        If (RowCount == 1)
            oSheet.Range("E" . RowCount).Value := "LastName"
        Else
            oSheet.Range("E" . RowCount).Value := LastName       
    }
}

; expression.Add (SourceType, Source, LinkSource, XlListObjectHasHeaders, Destination, TableStyleName)
oTable := oSheet.ListObjects.Add(1, oSheet.UsedRange,,1)
oTable.Name := "TeamUsersExport"

oTable.Range.Columns.AutoFit

oExcel.StatusBar := "READY"
} ; End of function
; -------------------------------------------------------------------------------------------------------------------

MenuCb_ToggleSettingTeamsPowerShell(ItemName, ItemPos, MenuName){
If GetKeyState("Ctrl") {
    sUrl := "https://tdalon.blogspot.com/2020/08/teams-powershell-setup.html"
    Run, "%sUrl%"
	return
}
RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
TeamsPowerShell := !TeamsPowerShell 
If (TeamsPowerShell) {
 	Menu,%MenuName%,Check, %ItemName%	 
    OfficeUid := People_GetMyOUid()
    ;sWinUid := People_ADGetUserField("mail=" . sEmail, "sAMAccountName")  ;- login uid
} Else
    Menu,%MenuName%,UnCheck, %ItemName%	 

PowerTools_RegWrite("TeamsPowerShell",TeamsPowerShell)
}

; -------------------------------------------------------------------------------------------------------------------
MenuCb_ToggleSettingTeamsMentionPersonalize(ItemName, ItemPos, MenuName){
If GetKeyState("Ctrl") {
    sUrl := "https://tdalon.blogspot.com/2020/08/teams-powershell-setup.html"
    Run, "%sUrl%"
	return
}
TeamsMentionPersonalize := PowerTools_RegRead("TeamsMentionPersonalize")
TeamsMentionPersonalize := !TeamsMentionPersonalize 
If (TeamsMentionPersonalize) {
 	Menu,%MenuName%,Check, %ItemName%	 
} Else
    Menu,%MenuName%,UnCheck, %ItemName%	 

PowerTools_RegWrite("TeamsMentionPersonalize",TeamsMentionPersonalize)
}

; -------------------------------------------------------------------------------------------------------------------

Teams_GetMainWindow(){
; See implementation explanations here: https://tdalon.blogspot.com/get-teams-window-ahk
; Syntax: hWnd := Teams_GetMainWindow()

WinGet, WinCount, Count, ahk_exe Teams.exe

If (WinCount = 0)
    GoTo, StartTeams

 ; fall-back if wrong exe found: close Teams
TeamsMainWinId := PowerTools_RegRead("TeamsMainWinId")

If WinExist("ahk_id " . TeamsMainWinId) {
    WinGet AhkExe, ProcessName, ahk_id %TeamsMainWinId% ; safe-check hWnd belongs to Teams.exe
    If (AhkExe = "Teams.exe")
        return TeamsMainWinId  
}

; when virtuawin is running Teams main window can be on another virtual desktop = hidden
Process, Exist, VirtuaWin.exe
VirtuaWinIsRunning := ErrorLevel
If (WinCount = 1) and Not (VirtuaWinIsRunning) {
    TeamsMainWinId := WinExist("ahk_exe Teams.exe")
    PowerTools_RegWrite("TeamsMainWinId",TeamsMainWinId)
    return TeamsMainWinId
}

; Get main window via Acc Window Object Name
WinGet, id, List,ahk_exe Teams.exe
Loop, %id%
{
    hWnd := id%A_Index%
    oAcc := Acc_Get("Object","4",0,"ahk_id " hWnd)
    sName := oAcc.accName(0)
    If RegExMatch(sName,".* \| Microsoft Teams, Main Window$") { ; works also for other lang
        PowerTools_RegWrite("TeamsMainWinId",hWnd)
        return hWnd
    }
}

; Fallback solution with minimize all window and run exe
If WinActive("ahk_exe Teams.exe") {
    GroupAdd, TeamsGroup, ahk_exe Teams.exe
    WinMinimize, ahk_group  TeamsGroup
} 

StartTeams: 
fTeamsExe = C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\current\Teams.exe
If !FileExist(fTeamsExe) {
    return
}
 
Run, "%fTeamsExe%""
WinWaitActive, ahk_exe Teams.exe
TeamsMainWinId := WinExist("A")
PowerTools_RegWrite("TeamsMainWinId",TeamsMainWinId)

return 

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_IsMainWindowActive() {
hWnd := WinActive("A")
hMain := Teams_GetMainWindow()
return (hWnd = hMain)

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Teams_SendCommand(sKeyword,sInput:="",Activate:= false) {
If Activate
    Teams_ActivateMainWindow()
Delay := PowerTools_GetParam("TeamsCommandDelay")
Send ^e ; Select Search bar
Send {Esc} ; Clear any preselection

If (SubStr(sKeyword,1,1) = "@") { 
    ;SendInput @
    ;sleep, 300
    ;SendInput % SubStr(sKeyword,2) 
    ;sleep, 300
    SendInput %sKeyword%
    Sleep %Delay% 
    SendInput {tab}}
} Else {
    SendInput /%sKeyword%
    
    Sleep %Delay% 
    SendInput +{enter}
}

If (sInput=="") ; empty
    return

;sLastChar := SubStr(sInput,StrLen(sInput)) 
doBreak := (SubStr(sInput,StrLen(sInput)) == "-")
If (doBreak) {
    sInput := SubStr(sInput,1,StrLen(sInput)-1) ; remove last -
}
Sleep %Delay%
SendRaw %sInput%
If (SubStr(sKeyword,1,1) = "@")
    Return
If doBreak
    return

Sleep %Delay%
SendInput +{enter}


} ; eofun


; -------------------------------------------------------------------------------------------------------------------
Teams_GetMeetingWindow(Maximize:=false, Activate:=false){
; Syntax: 
;      hWnd := Teams_GetMeetingWindow(Maximize:=true*|false,Activate:=true|false*) 
;   If window is not found, hwnd is empty
; Input Arguments:
;   Maximize: if true, minimized (aka Call in Progress) meeting window is maximized/clicked
;      This allows access for further meeting actions
;   Activate: if true, found Meeting Window will be activated
;
; See implementation explanations here: 
;   https://tdalon.blogspot.com/2022/07/ahk-get-teams-meeting-win.html

UIA := UIA_Interface()
WinGet, Win, List, ahk_exe Teams.exe
Loop %Win% {
    WinId := Win%A_Index%
    TeamsEl := UIA.ElementFromHandle(WinId)
    If Teams_IsMeetingWindow(TeamsEl)  {
        If (Maximize) {
            Lang := Teams_GetLang()
            Switch Lang 
            {
            Case "de-de":
                NavigateName := "Navigieren Sie zurück zum Anruffenster."
            Case "en-US":
                NavigateName := "Navigate back to call window."
            Default:
                NavigateName := "Navigate back to call window."
            }
            El := TeamsEl.FindFirstByNameAndType(NavigateName, "button") ; 
            
            If !Activate
                WinGet, prevWinId, ID, A
            
            If El {
                El.Click()
                Sleep 500
            }

            WinGet, WinId, ID, A ; Weird Teams Client Behavior: WinId changes after maximization
                
            If !Activate
                WinActivate, ahk_id %prevWinId%
        }

        If Activate {
            ;MsgBox Activate %WinId%
            WinActivate, ahk_id %WinId% ; sometimes does not work
        }
        return WinId

    }

    ;MsgBox % TeamsEl.DumpAll()
} ; End Loop

TrayTip, Could not find Meeting Window! , No active Teams meeting window found!,,0x2
} ; eofun



; -------------------------------------------------------------------------------------------------------------------

Teams_FindMeetingWindow(Activate:= false) {
; TeamsEl := Teams_FindMeetingWindow(Activate:= false)
; Loop through Teams.exe Window to find unminimized Meeting Window
; Does not return Minimized Window
; returns empty if not found and display a traytip message
UIA := UIA_Interface()
WinGet, Win, List, ahk_exe Teams.exe
Loop %Win% {
    WinId := Win%A_Index%
    TeamsEl := UIA.ElementFromHandle(WinId)
    
    If Teams_IsMeetingWindow(TeamsEl)  {
        if Activate
            WinActivate, ahk_id %WinId%
        return TeamsEl     
    }
}
Sleep 500

TrayTip, Could not find Meeting Window! , No unminmized active Teams meeting window!.,,0x2
} ; eofun

; ---------------------------------------------------------
Teams_IsMeetingWindow(TeamsEl,ExOnHold:=true){
; does not return true on Share / Call in progress window

; If Meeting Reactions Submenus are opened AutomationId are not visible.

Lang := Teams_GetLang()
Switch Lang 
{
Case "de-de":
    CallingControlsName := "Besprechungssteuerung"
    ResumeName := "Fortsetzen"
Case "en-US":
    CallingControlsName := "Calling controls"
    ResumeName := "Resume"
Default:
    CallingControlsName := "Calling controls"
    ResumeName := "Resume"
}

If TeamsEl.FindFirstByName(CallingControlsName) {
    ;or TeamsEl.FindFirstBy("AutomationId=meeting-apps-add-btn") or TeamsEl.FindFirstBy("AutomationId=hangup-btn") or TeamsEl.FindFirstByName("Applause"))
    If (ExOnHold) { ; exclude On Hold meeting windows
        If TeamsEl.FindFirstByName(ResumeName) ; Exclude On-hold meetings with Resume button ; TODO #lang specific
            return false
    }
    return true
}
return false
} ; eofun


; -------------------------------------------------------------------------------------------------------------------
Teams_ActivateMainWindow(){
; WinId := Teams_ActivateMainWindow()
WinId:= Teams_GetMainWindow()
WinActivate, ahk_id %WinId%
return WinId

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_NewMeeting(){
WinId := Teams_GetMainWindow()
WinActivate, ahk_id %WinId%
WinGetTitle Title, A
If ! (Title="Calendar | Microsoft Teams") {
    SendInput ^4 ; open calendar
    Sleep, 300
    While ! (Title="Calendar | Microsoft Teams") { 
        WinGetTitle Title, A
        Sleep 500
    }
}
SendInput !+n ; schedule a meeting alt+shift+n
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_MeetingOpenChat(sMeetingLink){
; Open Meeting Chat in Web browser from Teams Meeting url
; Used for Quick Join+
;sLink := UrlDecode(sMeetingLink)
RegExMatch(sMeetingLink,"https://teams.microsoft.com/l/meetup-join/([^/]*)",sId)
sChatLink := "https://teams.microsoft.com/_#/conversations/" . sId1 . "?ctx=chat"
Run, %sChatLink%
; Open in new window
;Run, chrome.exe "%sChatLink%" " --new-window --profile-directory=""Default"""
;WinWait, ahk_exe chrome.exe
;WinId := WinExist("ahk_exe chrome.exe") ; get last window
Browser_WinWaitActive()
WinId := Browser_WinExist()

; Send to second monitor
Monitor_MoveToSecondary(WinId)  
} ; eofun



; -------------------------------------------------------------------------------------------------------------------
Teams_NewConversation(){
; Using hotkeys - broken
;SendInput ^{f6} ; Activate posts tab https://support.microsoft.com/en-us/office/use-a-screen-reader-to-explore-and-navigate-microsoft-teams-47614fb0-a583-49f6-84da-6872223e74a0#picktab=windows
; workaround will flash the search bar if posts/content panel already selected but works now even if you have just selected the channel on the left navigation panel
;SendInput {Esc} ; in case expand box is already opened
; Conversation is already expanded if AutomationID=postTypeButton exists

; Alternative: UIA click on new-post-button
SendInput !+c ;  compose box alt+shift+c: necessary to get second hotkey working (regression with new conversation button)
sleep, 500
SendInput ^+x ; expand compose box ctrl+shift+x (does not work anymore immediately)
sleep, 500
SendInput +{Tab} ; move cursor back to subject line via shift+tab
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Teams_Pop(sInput){
; Pop-out chat via Teams command bar
WinId := Teams_GetMainWindow()
WinActivate, ahk_id %WinId%

Send ^e ; Select Search bar
SendInput /pop
Delay := PowerTools_GetParam("TeamsCommandDelay")
Sleep %Delay% 
SendInput {enter}
If (!sInput) ; empty
    return
sleep, 500

SendInput %sInput%
Sleep %Delay% 
SendInput {enter}

} ; eofun


; -------------------------------------------------------------------------------------------------------------------
Teams_MeetingShare(ShareMode := 2){
; ShareMode = 0 : unshare
; ShareMode = 1 : share
; ShareMode = 2: toggle share

WinId := Teams_GetMeetingWindow()
If !WinId ; empty
    return

UIA := UIA_Interface()
TeamsEl := UIA.ElementFromHandle(WinId)

ShareEl := TeamsEl.FindFirstBy("AutomationId=share-button")
;MsgBox % ShareEl.DumpAll() ; DBG
If !ShareEl {
    TrayTip TeamsShortcuts: ERROR, Share button not found!,,0x2
    return
}

IsSharing := !RegExMatch(ShareEl.Name,"^Share content")

If (ShareMode = 1) and (IsSharing) ; already sharing
    Return

If (ShareMode = 0) and !(IsSharing) ; already not sharing
    Return

;SendInput ^+e ; ctrl+shift+e - toggle share

ShareEl.Click() ; does not require Window to be active

If (ShareMode=0) or ((ShareMode=2) and IsSharing) ; unshare->done
    return 

; Wait for share tray to open
Delay := PowerTools_GetParam("TeamsShareDelay")
Sleep %Delay% 

; Include sound
El :=  TeamsEl.FindFirstByNameAndType("Include computer sound", "checkbox")
El.Click()

SendInput {Tab}{Tab}{Tab}{Enter} ; Select first screen - New Share design requires 3 {Tab}

; Move Meeting Window to secondary screen
; WinShiftRight Arrow
SysGet, MonitorCount, MonitorCount	; or try:    SysGet, var, 80
If (MonitorCount > 1) {
    ; Maximize Meeting Window by clicking on "Navigate back to call window" button
    El := TeamsEl.FindFirstByNameAndType("Navigate back to call window.", "button") ; TODO lang specific
    If El {
        El.Click()
        Sleep 500
    }
    ; Move to secondary monitor
    Monitor_MoveToSecondary(WinId,false)   ; bug: unshare on winactivate

    WinMaximize, ahk_id %WinId%
} ; end if secondary monitor


} ; eofun
; -------------------------------------------------------------------------------------------------------------------


; -------------------------------------------------------------------------------------------------------------------
Teams_ShareToTeams(sUrl:=""){
If GetKeyState("Ctrl") {
    Teamsy_Help("s2t")
	return
}
If (sUrl = "") && (Browser_WinActive()) {
    sUrl := Browser_GetActiveUrl()
}
InputBox, sUrl , Share To Teams, Enter Link to Share:, , 640, 125,,,,, %sUrl%
If ErrorLevel
    return
sUrl := "https://teams.microsoft.com/share?href=" + sUrl
Run, %sUrl%
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------
Teams_ClearCache(){
If GetKeyState("Ctrl") {
    Teamsy_Help("cl")
	return
}
Process, Exist, Teams.exe
If (ErrorLevel) {
    sCmd = taskkill /f /im "Teams.exe"
    Run %sCmd%,,Hide 
}

While WinExist("ahk_exe Teams.exe")
    Sleep 500

TeamsDir = %A_AppData%\Microsoft\Teams
FileRemoveDir, %TeamsDir%\application cache\cache, 1
FileRemoveDir, %TeamsDir%\blob_storage, 1
FileRemoveDir, %TeamsDir%\databases, 1
FileRemoveDir, %TeamsDir%\cache, 1
FileRemoveDir, %TeamsDir%\gpucache, 1
FileRemoveDir, %TeamsDir%\Indexeddb, 1
FileRemoveDir, %TeamsDir%\Local Storage, 1
FileRemoveDir, %TeamsDir%\tmp, 1

Teams_GetMainWindow()
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------
Teams_CleanRestart(){
If GetKeyState("Ctrl") {
    sUrl := "https://tdalon.blogspot.com/2021/01/teams-clear-cache.html"
    Run, "%sUrl%"
	return
}
; Warning all appdata will be deleted
MsgBox, 0x114,Teams Clean Restart, Are you sure you want to delete all Teams Client local application data?
IfMsgBox No
   return
Process, Exist, Teams.exe
If (ErrorLevel) {
    sCmd = taskkill /f /im "Teams.exe"
    Run %sCmd%,,Hide 
}
While WinExist("ahk_exe Teams.exe")
    Sleep 500

TeamsDir = %A_AppData%\Microsoft\Teams
FileRemoveDir, %TeamsDir%, 1

Teams_GetMainWindow()
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Restart(){
; Warning all appdata will be deleted
Process, Exist, Teams.exe
If (ErrorLevel) {
    sCmd = taskkill /f /im "Teams.exe"
    Run %sCmd%,,Hide 
}
While WinExist("ahk_exe Teams.exe")
    Sleep 500

Teams_GetMainWindow()
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Quit() {
sCmd = taskkill /f /im "Teams.exe"
Run %sCmd%,,Hide
} ; eofun 

; -------------------------------------------------------------------------------------------------------------------
Teams_Mute(State := 2){
; State: 
;    0: mute off, unmute
;    1: mute on
;    2*: (Default): Toggle mute state

WinId := Teams_GetMeetingWindow() 
If !WinId ; empty
    return
UIA := UIA_Interface()
TeamsEl := UIA.ElementFromHandle(WinId)

Lang := Teams_GetLang()
Switch Lang 
{
Case "de-de":
    MuteName := "Stummschalten"
    UnmuteName := "Stummschaltung"
Case "en-US":
    MuteName := "Mute"
    UnmuteName := "Unmute"
Default:
    MuteName := "Mute"
    UnmuteName := "Unmute"
}

El :=  TeamsEl.FindFirstByNameAndType(MuteName, "button",,2)
If El {
    If (State = 0) {
        Tooltip("Teams Mic is alreay on.")
        return
    } Else {
        Tooltip("Teams Mute Mic...")
        El.Click()
        return
    }
}
El :=  TeamsEl.FindFirstByNameAndType(UnmuteName, "button",,2)
If (State = 1) {
    Tooltip("Teams Mic is alreay off.")
    return
} Else {
    Tooltip("Teams Unmute Mic...")
    El.Click()
    return
}

} ; eofun


; -------------------------------------------------------------------------------------------------------------------
Teams_MeetingShortcuts(sKeyword) {
WinId := Teams_GetMeetingWindow()
If !WinId ; empty
    return
;MsgBox % WinId
WinGet, curWinId, ID, A
WinActivate, ahk_id %WinId%

Switch sKeyword
{
    Case "mu","mute":
        Tooltip("Teams Toggle Mic...") 
        SendInput ^+m ;  ctrl+shift+m 
    Case "bg","bl","blur":
        Tooltip("Teams Background Settings...") ; toggle background blur
        SendInput ^+p ;  ctrl+shift+p 
        return ; need human action to change setting; do not restore window
    Case "lobby": ; Admit people from lobby notification
        Tooltip("Teams Admit from Lobby...") ; toggle background blur
        SendInput ^+y ;  ctrl+shift+y
        return ; need human action to change setting; do not restore window
    Case "share-accept":
        Tooltip("Teams Accept Share...") 
        SendInput ^+a ;  ctrl+shift+A
    Case "share-decline":
        Tooltip("Teams Decline Share...") 
        SendInput ^+d ;  ctrl+shift+D
    Case "share":
        Tooltip("Teams Start Share...") 
        SendInput ^+e ;  Start screen share session Ctrl+Shift+E
    Case "share-toolbar":
        Tooltip("Teams Start Share...") 
        SendInput ^+e{Space} ; Go to sharing toolbar Ctrl+Shift+Spacebar
}


Sleep 500 ; pause before reactivating previous window
WinActivate, ahk_id %curWinId%


} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Leave() {
WinId := Teams_GetMeetingWindow(false,true)
If !WinId ; empty
    return
SendInput ^+h ; Ctrl+Shift+H
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_PushToTalk(){

WinId := Teams_GetMeetingWindow() 

If !WinId ; empty
    return

KeyName := A_PriorKey
If (KeyName = "m") or (KeyName = "^") or (KeyName = "!") {
    MsgBox Error: hotkey conflict with native Ctrl+Shift+M. Choose another combination.
    return
}

WinGet, curWinId, ID, A
WinActivate, ahk_id %WinId%
IsMuted := Teams_IsMuted(WindId)

ToolTip  Teams PushToTalk on... 
If (IsMuted)
    SendInput ^+m ;  ctrl+shift+m

while (GetKeyState(KeyName , "P"))
{
sleep, 100
}
WinActivate, ahk_id %WinId%
SendInput ^+m ;  ctrl+shift+m
WinActivate, ahk_id %curWinId%

Tooltip("Teams PushToTalk off...",2000)  
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------

Teams_HotkeySet(HKid){
If GetKeyState("Ctrl")  { ; exclude ctrl if use in the hotkey
	sUrl := "https://tdalon.github.io/ahk/Teams-Global-Hotkeys"
    Run, "%sUrl%"
    return
}

; For Menu callback, remove ending Hotkey and blanks and (Hotkey)

HKid := RegExReplace(HKid,"\t(.*)","")
;HKid := Trim(HKid) ; remove tab for align right hotkey in menu
HKid := RegExReplace(HKid," Hotkey$","")
HKid := StrReplace(HKid," ","")

RegRead, prevHK, HKEY_CURRENT_USER\Software\PowerTools, TeamsHotkey%HKid%
newHK := HotkeyGUI(,prevHK,,,"Teams " . HKid . " - Set Global Hotkey")

If ErrorLevel ; Cancelled
    return
If (newHK = prevHK) ; no change
    return

RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, TeamsHotkey%HKid%, %newHK%

If (newHK = "") { ; reset/ disable hotkey
    ;Turn off the new hotkey.
    Hotkey, %prevHK%, Teams_%HKid%, Off 
    TipText = Set Teams %HKid% Hotkey off!
    TrayTipAutoHide("Teams " . HKid . " Hotkey Off",TipText,2000)
    return
}

; Turn off the old Hotkey
If Not (prevHK == "")
	Hotkey, %prevHK%, Teams_%HKid%, Off

Teams_HotkeyActivate(HKid,newHK, True)

; Refresh Menus (function is only called from Script Menu)
Reload


} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Teams_HotkeyActivate(HKid,HK,showTrayTip := False) {
;Turn on the new hotkey.
Hotkey, IfWinActive, ; for all windows/ global hotkey
Hotkey, $%HK%, Teams_%HKid%, On ; use $ to avoid self-referring hotkey if Ctrl+Shift+M is used
If (showTrayTip) {
    TipText = Teams %HKid% Hotkey set to %HK%
    TrayTipAutoHide("Teams " . HKid . " Hotkey On",TipText,2000)
}
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------

Teams_IsMuted(hwnd:="") {
; IsMuted := Teams_IsMuted(hwnd*)
; hwnd: Meeting Window. If not input, Teams_GetMeetingWindow is called;
; Output IsMuted logical. Empty if no meeting window found.
if !hwnd {
    hwnd:=Teams_GetMeetingWindow()
    If !hwnd ; empty
        return
}
UIA := UIA_Interface()
UIA.AutoSetFocus := false
TeamsEl := UIA.ElementFromHandle(hwnd)
If !TeamsEl.FindFirstBy("AutomationId=microphone-button") {
    MsgBox Microphone button is not accessible!
    MsgBox % TeamsEl.DumpAll()
}
return !TeamsEl.FindFirstBy("Name=Mute (Ctrl+Shift+M)") 
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Video(State := 2){
; State: 
;    0: video off
;    1: video on
;    2*: (Default): Toggle video


WinId := Teams_GetMeetingWindow() 
If !WinId ; empty
    return
UIA := UIA_Interface()
TeamsEl := UIA.ElementFromHandle(WinId)

;El :=  TeamsEl.FindFirstByNameAndType("Turn camera on", "button",,1)
El :=  TeamsEl.FindFirstByName("Turn camera on",,1) ; menu item on meeting window; button on Call in progress
If El {
    If (State = 0) {
        Tooltip("Teams Camera is alreay off.")
        return
    } Else {
        Tooltip("Teams Camera On...")
        El.Click()
        return
    }
}
;El :=  TeamsEl.FindFirstByNameAndType("Turn camera off", "button",,1)
El :=  TeamsEl.FindFirstByName("Turn camera off",,1)
If El {
    If (State = 1) {
        Tooltip("Teams Camera is alreay on.")
        return
    } Else {
        Tooltip("Teams Camera off...")
        El.Click()
        return
    }
}

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
Teams_MuteApp(sCmd:= ""){
Switch sCmd
{
    Case "s","sw","switch":
        sCmd = /Switch
    Case "on","1":
        sCmd = /Mute
    Case "off","0":
        sCmd = /Unmute
    Default :
        sCmd = /Switch ; works even if used as menu callback
} ; end switch
SVVExe := GetSoundVolumneViewExe()
If (SVVExe = "") {
    return
}
    
;TeamsExe := Teams_GetExe()
sCmd = "%SVVExe%" %sCmd% "Teams.exe"
Run, %sCmd%
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

SetSoundVolumeViewExe(){
FileSelectFile, SVVExe , 1, SoundVolumeView.exe, Select the location of SoundVolumeView.exe, SoundVolumeView.exe
If ErrorLevel
    return
PowerTools_RegWrite("SoundVolumeViewExe",SVVExe)
return SVVExe
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

GetSoundVolumneViewExe(){
; SVVExe := Mute_GetSoundVolumneViewExe()
SVVExe := PowerTools_RegRead("SoundVolumeViewExe")  
If (SVVExe="") {
    Run, "https://www.nirsoft.net/utils/sound_volume_view.html"
    SVVExe := SetSoundVolumeViewExe()
}
return SVVExe    
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_RaiseHand() {
; Toggle Raise Hand on/off ; Default Hotkey Ctrl+Shift+K
WinGet, curWinId, ID, A
WinId := Teams_GetMeetingWindow(true,true)
If !WinId ; empty
    return
Tooltip("Teams Toggle Raise Hand...") 
SendInput ^+k ; toggle video Ctl+Shift+k
WinActivate, ahk_id %curWinId%
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_React(sReaction) {
; sReaction can be Like | Applause| Heart | Laugh
WinId := Teams_GetMeetingWindow(true)
If !WinId ; empty
    return
If  A_IsCompiled
    ImgFile = .\imgsearch\Teams_reactions.png
Else
    ImgFile = .\PowerTools\imgsearch\Teams_reactions.png
If !FileExist(ImgFile) {
    ;Tooltip("Teams Meeting Reaction: ERROR: " . ImgFile . " file not found!",1000)
    TrayTip Teams Meeting Reaction: ERROR, %ImgFile% file does not exist! 
    ; MsgBox 0x10, Teams Shortcuts: Error, Teams Meeting Reaction: ERROR: %ImgFile% file not found!
    Run, "https://tdalon.github.io/ahk/Teams-Meeting-Reactions"
    return
}

WinGet, curWinId, ID, A
MouseGetPos , MouseX, MouseY
WinActivate, ahk_id %WinId%

WinGetPos , , ,WinWidth, WinHeight, A

ImageSearch, FoundX, FoundY, 0, 0, WinWidth, WinHeight, *2 %ImgFile%
If (ErrorLevel = 0)
	Click, %FoundX%, %FoundY% Left, 1
Else {
    TrayTip Teams Meeting Reaction: ERROR, %ImgFile% image search failed! Create another screenshot!
    WinActivate, ahk_id %curWinId%
    Run, "https://tdalon.github.io/ahk/Teams-Meeting-Reactions"
    Return
}	
Sleep 200
If  A_IsCompiled 
    ImgFile = .\img\Teams_%sReaction%.png
Else
    ImgFile = .\PowerTools\img\Teams_%sReaction%.png
If !FileExist(ImgFile) {
    ;Tooltip("Teams Meeting Reaction: ERROR: " . ImgFile . " file not found!",1000)
    TrayTip Teams Meeting Reaction: ERROR, %ImgFile% file does not exist!
    WinActivate, ahk_id %curWinId%
    MouseMove, MouseX, MouseY
    Run, "https://tdalon.github.io/ahk/Teams-Meeting-Reactions"
    return
}
Retry:
ImageSearch, FoundX, FoundY, 0, 0, WinWidth, WinHeight, *2 %ImgFile%
;MsgBox %WinWidth% %WinHeight%
If (ErrorLevel = 0) {
    Click, %FoundX%, %FoundY% Left, 1
} Else {
      
    ;TrayTipAutoHide("Teams Meeting Reaction: ERROR:","ERROR: " . ImgFile . " image not found!")
    TrayTip Teams Meeting Reaction: ERROR, %ImgFile% image search failed! Create another screenshot.
    WinActivate, ahk_id %curWinId%
    MouseMove, MouseX, MouseY
    Run, "https://tdalon.github.io/ahk/Teams-Meeting-Reactions"
    return
}
WinActivate, ahk_id %curWinId%
MouseMove, MouseX, MouseY
If (ErrorLevel = 0)
    Tooltip("Teams Meeting Reaction: " . sReaction,1000) 
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Teams_IsWinActive(){
; Check if Active window is Teams client or a Browser/App window with a Teams url opened

If WinActive("ahk_exe Teams.exe")
    return True
SetTitleMatchMode, RegEx
If WinActive("\| Microsoft Teams$")
    return True
return False
}

; -------------------------------------------------------------------------------------------------------------------
Teams_MeetingReaction(Reaction) {
; Reaction can be Like | Applause| Love | Laugh | Surprised
; See documentation https://tdalon.blogspot.com/2022/07/ahk-teams-meeting-reactions-uia.html
WinId := Teams_GetMeetingWindow(true,true)
If !WinId ; empty
    return
WinGet, curWinId, ID, A

UIA := UIA_Interface()  
TeamsEl := UIA.ElementFromHandle(WinId)

; Shortcut if Reactions toolbar already available-> directly click and exit
ReactionEl := TeamsEl.FindFirstByName(Reaction)
If ReactionEl 
    Goto, React

ReactionsEl :=  TeamsEl.FindFirstBy("AutomationId=reaction-menu-button")  
If ReactionsEl
    Goto ClickReactions

; Workaround button not found: Click by position

;WinActivate, ahk_id %WinId% ; needs to be activated for UIA.ElementFromPoint to work

;MsgBox % TeamsEl.DumpAll()

; Look for Chat button
BtnEl := TeamsEl.FindFirstBy("AutomationId=chat-button") 
If BtnEl {
    BR := BtnEl.CurrentBoundingRectangle
    ReactionsEl := UIA.ElementFromPoint(BR.r+(BR.r-BR.l)//2, BR.t+20) ; Reactions button after chat
    ReactionsEl :=  TeamsEl.FindFirstBy("AutomationId=reaction-menu-button") ; Sometimes Element is made accessible by ElementFromPoint but returns only parent window element
    If ReactionsEl
        Goto ClickReactions
} 

; Look for Roster/People button
BtnEl := TeamsEl.FindFirstBy("AutomationId=roster-button") 
If BtnEl {
    BR := BtnEl.CurrentBoundingRectangle
    ReactionsEl := UIA.ElementFromPoint(BR.r+3*(BR.r-BR.l)//2, BR.t+20) ; Reactions button after chat after roster
    ReactionsEl :=  TeamsEl.FindFirstBy("AutomationId=reaction-menu-button") ; Sometimes Element is made accessible by ElementFromPoint but returns only parent window element
    If ReactionsEl
        Goto ClickReactions
} 

; Look for Controls meeting Element
ControlsEl := TeamsEl.FindFirstByName("Meeting controls")
If ControlsEl { ; Meeting controls not accessible
    br := ControlsEl.CurrentBoundingRectangle
    btnwidth := (br.r-br.l)//10 ; there are 10 buttons in the Meeting controls block
    ReactionsEl := UIA.ElementFromPoint(br.l+2.5*btnwidth, br.t+20) ; Reactions button is on 3rd position 
    ReactionsEl :=  TeamsEl.FindFirstBy("AutomationId=reaction-menu-button") ; Sometimes Element is made accessible by ElementFromPoint but returns only parent window element
    If ReactionsEl 
        Goto ClickReactions
} 

;PowerTools_ErrDlg("Meeting Reactions button not found by Id nor position!")
TrayTip TeamsShortcuts: ERROR, Meeting Reactions button not found by Id nor position!,,0x2
return

ClickReactions:
ReactionsEl.Click() ; Click element without moving the mouse
ReactionEl:=TeamsEl.WaitElementExistByName(Reaction,,,,2000) ; timeout=2s

If !ReactionEl {
    TrayTip TeamsShortcuts: ERROR, Meeting Reaction button for '%Reaction%'' not found!,,0x2
    ;MsgBox % ReactionsEl.DumpAll() ; DBG
    return
} 

React:
Tooltip("Teams Meeting Reaction: " . Reaction,1000)
ReactionEl.Click()

; Close Reactions Menu because it makes other buttons invisible
Sleep 500
SendInput {Esc}

; Restore previous window 
WinActivate, ahk_id %curWinId%
     
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Teams_Click(Id) {
If (ok := Teams_FindText(Id)) 
{
CoordMode, Mouse
X:=ok.1.x, Y:=ok.1.y, Comment:=ok.1.id
Click, %X%, %Y%
}
return ok

} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Teams_FindText(Id){
; ok := Teams_FindText(Id)
Text := Teams_GetText(Id)

WinGetPos, x,y,w,h,A ; (x,y) upper left corner
X1 := x, Y1 := y, X2 := x+w, Y2 := y+h  ; (X1,Y1): upper left corner, (X2,Y2): lower right corner
ok:=FindText(X1,Y1,X2,Y2, 0, 0, Text,,0) ; last arg FindAll
;ok:=FindText(,,,, 0, 0, Text,,0) ; last arg FindAll
 /*
MsgBox, 4096, Tip, % "Found:`t" Round(ok.MaxIndex())
   . "`n`nTime:`t" (A_TickCount-t1) " ms"
   . "`n`nPos:`t" X ", " Y
   . "`n`nResult:`t" (ok ? "Success !" : "Failed !")
*/
;t1:=A_TickCount, X:=Y:=""

;WinGetPos, x,y,w,h,A ; (x,y) upper left corner
;X1 := x, Y1 := y, X2 := x+w, Y2 := y-h  ; (X1,Y1): upper left corner, (X2,Y2): lower right corner
;X1 := 1303-150000, Y1:= 90-150000, X2 := 1303+150000, Y2 := 90+150000
;MsgBox %X1% %Y1% %X2% %Y2%
;ok:=FindText(X1,Y1,X2,Y2, 0, 0, Text,,0) ; last arg FindAll
return ok
} ;eofun
; -------------------------------------------------------------------------------------------------------------------

Teams_GetText(Id,Def:=False){
; Text := Teams_GetText(Id,Def)
; If Def = False, take value from registry

If !(Def) {
    return PowerTools_GetParam("TeamsFindText" . Id)
}

; Default values
Switch Id
{
Case "MeetingActions": ; 3 dots
    Text:="|<>*103$26.zzzzzzzzzzzzzzzzzzzzzzzzzzyC77z1VUzkMMDyC77zzzzzzzzzzzzzzzzzzzzzzzzzzs"
Case "MeetingReactions":
    Text:="|<>*100$22.zzzzwzzz0zzk1zy07zs0TzU1zy6VztyGTby8yTltxyTnnnvjaNaS3ztzzzbzN6zwknztyTzk3zzUTzzzy"
Case "MeetingReactionHeart":
    Text:="|<>*74$26.zzzzzUT1zk307s001w000D0003U000M00060001U000Q00070003s000y000Tk00Dy007zs03zz03zzs1zzz0zzzwTzzzjzzzzzy"
Case "MeetingReactionLaugh":
    Text:="|<>*126$24.zzzzzwDzzU1zy00Tw00Ds4A7sGG7k003k003k003V00VU001U001V001l00XkU13kE03s4A7s007w00Dy00TzU1zzwDzzzzzU"
Case "MeetingReactionApplause":
    Text:="|<>*116$25.zzzzxrzzzzzrzzzzzN7zzw1vjz0Qzy06Dz01bz001zU00zk00Ds007y003z001zU00zs007y003zU01zw00zzU0zzzUzzzzzzzzzU"
Case "MeetingReactionLike":
    Text:="|<>*119$20.zzzzzwzzyDzzXzzlzzkTzsDzw3zy0zzU0zk07E01U00M006001U00M006003U00w00TzUDzzzs"
Case "MeetingActionFullScreen":
    Text:="|<>*106$93.zzzzzzzzzzzzzzzy00TzUTwbzzzzzzzjzxzwzzYzzzzzzzx7sjzbzwbzzzzzzzfzpzwwtYy3123VURTyjzbbAbnH9aNYFjzxzw4tYySTAsQbBzzjzbbAbsntU04tfzpzwwtYzmTAyTbBTyjzbmAbqH9btwtcz5zwy1Yy71C3UbBzzjzzzzzzzzzzzzk03zzzzzzzzzzzzzzzzzzzzzzzzzzzzU"
Case "MeetingActionTogetherMode":
    Text:="|<>*66$18.zzzySTxhjxhjySTzzztnbqhPqhPtnbzzzV0VjSxbStlVXzvzzzzU"
Case "MeetingActionBackgrounds":
    Text:="|<>*104$18.zzztgrnThaVNhhnvAbrhBgVPtzrm0RazNizHuzLzzzU"
Case "MeetingActionShare":
    Text:="|<>*161$22.zzzzzzzw003U0060A0M1s1UDk61hUM0k1U3060A0M0k1U3060A0M001k00DzzzzzzzU"
Case "MeetingActionUnShare":
    Text:="|<>*153$22.zzzzzzzw003U006000M421U8E60G0M0k1U3060G0M241UE86000M001k00DzzzzzzzU"
Case "Muted","Unmute":
    ;Text:="|<>*113$22.zzzyzVzxw3zvU7zq0Tzg1zzM7zykTztVzzX7zy6Tz8BDwkQzv0rzbBzzDnzyA7zy7jzwzTznyzzjxzzzy" ; does not work
Case "Mute","Unmuted":
    ;Text:="|<>*111$18.zzzzzzzVzz0zz0zy0Ty0Ty0Ty0Ty0Ty0Ty0Tm0Hn0nv0rtnbwzDy0TzVzznzznzzvzzzzU" 
    Text:="|<>*112$31.zzzzzzzzzzzzzzzzzzzzzzy7zzzy1zzzy0Tzzz0Dzzz03zzzU1zzzk0zzzs0Tzzw0Dzzy07zzz03zzxU1jzys1rzzQ0vzza0NzznURzzwzwzzzDwzzzlszzzy1zzzznzzzztzzzzwzzzzyTzzzzzzzzzzzzzzzzzk"
Case "Leave":
    Text:="|<>*155$81.zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzwzzzzzzzzk0Dzzbzzzzzzzk00Dzwzzzzzzzw000zzbw70XX1z0003zwz0M2AE7k1y0DzbtlwFWQS0Ts1zwy0A2AU3k3z0Dzbk30N40S0zw1zwyDlX1Xzs7zUTzbluAMCCzrzzbzw10E3Vk7zzzzzzUA30QT0zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzU"
Case "Resume":
    Text:="|<>*138$51.zzzzzzzzw3zzzzzzzUDzzzzzzwtzzzzzzzbA64NU1kQ1423A04FUtUQNa8aAX0EXAlY1a9zUNaAbwt4Y0AlYHb4461aAkTzzzzzzzzU" ; FindText for Resume
Case "NewConversation":
    Text:="|<>*155$184.0000000000000000000000000000000zzzzzzzzzzzzzzzzzzzzzzzzzzzzzz7zzzzzzzzzzzzzzzzzzzzzzzzzzzzzyTzzzzzzzzzzzzzzzzzzzzzzzzzzzzztzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzbzzzzzzzzzzzzzzzzzzzzzzzzzzzzzyTzzzzzzzzzzzzzzzzzzzzzzzzzzzzztzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzbzzzzzzzzzzzzzzzzzzzzzzzzzzzzzyTzzzzzzzzzzzzzzzzzzzzzzzzzzzzztzzzzzvzzzzzzzzzzzzzzzzzzzzzzzzbzzzk7TzzzzzzzzzzzzzzzzzzzzzzzyTzzyTvzzllzzzzzzzzzzzzzzDzzzzztzzzvzTzz77zzzzzzzzzzzzyMzzzzzzbzzzjvzzwATzzzzzzzzzzzztzzzzzzyTzzyzRzzkFUa8y31UFW30230A60zzztzzzvvrzz94E8bk820aF4M968U83zzzbzzzjTTzwUH0WTD78mNAHbyNWQX7zzyTzzyzxzzn10E1wwQX841C21a1kATzztzzzvzrzzC4z27nlmAknwy1aMb8lzzzbzzzjzTzwsFA8z0U8n34nc4M20X7zzyTzzyTtzznlUtXy31XCS3C61UA6ATzztzzzw0DzzzzzzzzzzzzzzzzzzzzzzzzbzzzzzzzzzzzzzzzzzzzzzzzzzzzzzyTzzzzzzzzzzzzzzzzzzzzzzzzzzzzztzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzbzzzzzzzzzzzzzzzzzzzzzzzzzzzzzyTzzzzzzzzzzzzzzzzzzzzzzzzzzzzztzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzbzzzzzzzzzzzzzzzzzzzzzzzzzzzzzyTzzzzzzzzzzzzzzzzzzzzzzzzzzzzztzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzXzzzzzzzzzzzzzzzzzzzzzzzzzzzzzw0000000000000000000000000000002"
} ; eoswitch

return Text


} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_ImageSearch(img) {
If  A_IsCompiled 
    ImgFile = .\imgsearch\%img%
Else
    ImgFile = .\PowerTools\imgsearch\%img%

If !FileExist(ImgFile) {
    ;Tooltip("Teams Meeting Reaction: ERROR: " . ImgFile . " file not found!",1000)
    TrayTip Teams Shortcuts: ERROR, %ImgFile% file does not exist!
    return
}
WinGetPos , , ,WinWidth, WinHeight, A

CoordMode, Pixel, Screen
ImageSearch, x, y, 0, 0, A_ScreenWidth, A_ScreenHeight, *20 %ImgFile%

;ImageSearch, FoundX, FoundY, 0, 0, WinWidth, WinHeight, *4 %ImgFile%
If (ErrorLevel = 0) 
    return [FoundX, FoundY]
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Teams_MeetingToggleFullscreen(WinId:="",restore:=True) {
; Teams_MeetingToggleFullscreen(WinId:="",restore:=True)
; Arguments: 
;    WinId: Optional: pass Meeting Window WinId if known
;    restore: True*|False
;               If true previous current window will be restored/ activated back. 
;               If false, Meeting Window will be activated and stays activated
If !WinId
    WinId := Teams_GetMeetingWindow()
If !WinId ; empty
    return
; Needs to activate the Meeting Window because F11 Hotkey is not working even with ControlSend,,{F11}, ahk_id %WinId%
If (restore)
    WinGet, curWinId, ID, A

WinActivate, ahk_id %WinId% 
Send {F11}
; restore previous window
If (restore)
    WinActivate, ahk_id %curWinId%
}


Teams_MeetingAction(id){
; id: recording, fullscreen, device-settings, incoming-video
WinId := Teams_GetMeetingWindow() 
If !WinId ; empty
    return
UIA := UIA_Interface()
TeamsEl := UIA.ElementFromHandle(WinId)
sFindBtn := "AutomationId=" . id . "-button"

/* 
; If more Menu already clicked

actionEl :=  teamsEl.FindFirstBy(sFindBtn)  
If actionEl
    Goto ClickAction 
*/

moreEl := TeamsEl.FindFirstBy("AutomationId=callingButtons-showMoreBtn")
Sleep, 200
;moreEl.Highlight()
moreEl.Click()
Sleep, 100
actionEl:= TeamsEl.WaitElementExist(sFindBtn,,,,1000)
;actionEl:= teamsEl.WaitElementExist("AutomationId=fullscreen-button")
If !actionEl
    return

Sleep, 500
actionEl.Highlight()
Sleep, 500
actionEl.Click()
;MsgBox % actionEl.DumpAll() ; DBG
return
/* 
moreEl := teamsEl.FindFirstBy("AutomationId=callingButtons-showMoreBtn") 
Sleep, 100
moreEl.Click()
If !moreEl {
    TrayTip TeamsShortcuts: ERROR, More Actions button not found!,,0x2
    ;MsgBox % moreEl.DumpAll() ; DBG
    return
} 


actionEl:=teamsEl.WaitElementExist(sFindBtn,,,,1000) ; timeout=2s

If !actionEl {
    TrayTip TeamsShortcuts: ERROR, Meeting Action button for '%id%'' not found!,,0x2
    MsgBox % actionEl.DumpAll() ; DBG
    return
} 
*/ 

ClickAction:
Sleep, 500
actionEl.Click()
; Tooltip("Teams Meeting Action: " . id,1000) : will hide action menu
TrayTip TeamsShortcuts: Meeting Action, Button for '%id%'' clicked!,,0x1


}


Teams_Join(){
; Join a Teams Meeting from Outlook Reminder Window
If WinActive("Reminder(s) ahk_class #32770"){ ; Reminder Windows
; Keys are blocked by the UI: c,d,a,s. Alt does not work
	WinActivate
	;Send +{F10} ; Shift+F10 - Open Context Menu
	SendInput {Tab}
	Send j ; Join accelerator
	WinWaitActive, ahk_exe Teams.exe,,5
	If ErrorLevel
		Return
}
UIA := UIA_Interface()
If !WinId
    WinId := WinActive("A")
TeamsEl := UIA.ElementFromHandle(WinId) 
JoinBtn :=  TeamsEl.FindFirstBy("AutomationId=prejoin-join-button")  


}

; -------------------------------------------------------------------------------------------------------------------

Teams_MeetingActionClick(id, restore:= False){
; Based on FindText. Obsolete: replaced by Teams_MeetingAction based on UIAutomation
WinId := Teams_GetMeetingWindow()
If !WinId ; empty
    return

If (restore) {
    WinGet, curWinId, ID, A
    MouseGetPos , MouseX, MouseY
}
WinActivate, ahk_id %WinId%

ok := Teams_Click("MeetingActions")
If !(ok) {
    TrayTip Teams Meeting Action: ERROR, FindText failed!
    ;Run, "https://tdalon.github.io/ahk/Teams-Meeting-Reactions"
    return
}

Delay := PowerTools_GetParam("TeamsClickDelay")
Sleep %Delay% 

ok := Teams_Click("MeetingAction" . id)

If (restore) { ; Restore previous window and mouse position
    WinActivate, ahk_id %curWinId%
    MouseMove, MouseX, MouseY
}

If (ok)
    Tooltip("Teams Meeting Action: " . id,1000) 
}
; -------------------------------------------------------------------------------------------------------------------
Teams_ClearFlash(){
; Will activates Teams Main window and restore back to the previous window and cursor position
WinGet hcurwin
;MouseGetPos , MouseX, MouseY
MouseX := A_CaretX
MouseY := A_CaretY
; Activate Teams Main Window
hteamswin := Teams_GetMainWindow()
WinGet, MinMax , MinMax, ahk_id %hteamswin%
WinActivate, ahk_id %hteamswin%
; Restore minimize status
If (MinMax= -1)
    WinMinimize, ahk_id %hteamswin% 
; Reactivate previous window
WinActivate ahk_id %hcurwin%
; Restore cursor
Click, %MouseX%, %MouseY%
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Teams_GetLang() {
; return desktop client language
; sLang := Teams_GetLang()

; Read value of currentWebLanguage property in %AppData%\Microsoft\Teams\desktop-config.json
; e.g. "en-US"
JsonFile := A_AppData . "\Microsoft\Teams\desktop-config.json"
FileRead, Json, %JsonFile%
If ErrorLevel {
    TrayTip, Error, Reading file %JsonFile%!,,3
    return
}
oJson := Jxon_Load(Json)
Lang := oJson["currentWebLanguage"]
return Lang
} ; eofun