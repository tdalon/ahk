#Include <People>
#Include <UriDecode>
#Include <Teamsy>
#Include <Clip>
#Include <FindText>


Teams_Launcher(){
Teamsy("-g")
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Teams_OpenBackgroundFolder(){
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2021/01/teams-custom-backgrounds.html#openfolder"
	return
}
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
Teams_Emails2Chat(sEmailList){
; Open Teams 1-1 Chat or Group Chat from list of Emails
sLink := "msteams:/l/chat/0/0?users=" . StrReplace(sEmailList, ";",",") ; msteams:
If InStr(sEmailList,";") { ; Group Chat
    InputBox, sTopicName, Enter Group Chat Name,,,,100
    if (ErrorLevel=0) { ;  TopicName defined
        sLink := sLink . "&topicName=" . StrReplace(sTopicName, ":", "")
    }
} 
Run, %sLink% 
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Emails2Meeting(sEmailList){
; Open Teams 1-1 Chat or Group Chat from list of Emails
sLink := "https://teams.microsoft.com/l/meeting/new?attendees=" . StrReplace(sEmailList, ";",",") ; msteams:
TeamsExe = C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\current\Teams.exe
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

sUrl:= StrReplace(sUrl,"https:","msteams:") ; open by default in app
sFile := FavsDir . "\" . sFileName . ".url"

;FileDelete %sFile%
IniWrite, %sUrl%, %sFile%, InternetShortcut, URL

; Add icon file:
TeamsExe = Teams_GetExe()
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

; ----------------------------------------------------------------------
Teams_Emails2Favs(sEmailList:= ""){
; Calls: Email2TeamsFavs

If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2021/03/teams-people-favorites.html"
	return
}
If (sEmailList ="") {
    sSelection := People_GetSelection()
    If (sSelection = "") { 
        TrayTipAutoHide("Teams: Email to Fav!","Nothing selected!")   
        return
    }
    sEmailList := People_GetEmailList(sSelection)
    If (sEmailList = "") { 
        TrayTipAutoHide("Teams: Email to Fav!","No email could be found from Selection!")   
        return
    }
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
; Calls: Teams_Link2Fav

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

Teams_Emails2Users(EmailList,TeamLink:=""){
; Syntax: 
;     Emails2TeamMembers(EmailList,TeamLink*)
; EmailList: String of email adressess separated with a ;
; TeamLink: (String) optional. If not passed, user will be asked via inputbox
;  e.g. https://teams.microsoft.com/l/team/19%3a12d90de31c6e44759ba622f50e3782fe%40thread.skype/conversations?groupId=640b2f00-7b35-41b2-9e32-5ce9f5fcbd01&tenantId=xxx

If GetKeyState("Ctrl") {
	Run, "https://connectionsroot/blogs/tdalon/entry/emails2teammembers"
	return
}

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

If GetKeyState("Ctrl") {
	Run, "https://connectionsroot/blogs/tdalon/entry/emails2teammembers" ; TODO
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

    ;SendInput +{Up} ; Shift + Up Arrow: select all thread -> not needed anymore. covered by Ctrl+A
    SendInput ^a ; Select all thread
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
Clip_PasteHtml(sQuoteHtml,,False)

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
TeamsCopyLink(){
Send {Enter}
Send {Tab}
Send {Enter}
;Sleep 500
;Send {Down 2}

MsgBox 0x1001, Wait..., Wait for Copy Link. Press OK to continue. 
IfMsgBox Cancel
    return False
return true
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_CopyLink(){
SendInput +{Up} ; Shift + Up Arrow: select all thread
Sleep 200
SendInput ^a
sSelection := Clip_GetSelection()
; Last part between <> is the thread link
RegExMatch(sSelection,"U).*<(.*)>$",sMatch)
SendInput {Esc}
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
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2020/11/teams-shortcuts-personalize-mentions.html"
	return
}
SendInput +{Left}
sLastLetter := Clip_GetSelectionHtml()
IsRealMention := InStr(sLastLetter,"http://schema.skype.com/Mention") 
sLastLetter := Clip_GetSelection()    
SendInput {Right}

If (IsRealMention) {
    If (sLastLetter = ")")
        SendInput {Backspace}^{Left}
    Else 
        SendInput ^{Left}

    SendInput +{Left} 
    sLastLetter := Clip_GetSelection()   
    If (sLastLetter = "-")
        SendInput ^{Left}{Backspace}
    Else 
        SendInput {Backspace}
} Else {
    ; do not personalize if no match to avoid ambiguity which name was mentioned if shorten to firstname and reprompt for matching
    return
    If (sLastLetter = ")") 
        SendInput {Backspace}^{Backspace}{Backspace} ; remove ()

    SendInput ^{Left}^{Backspace}^{Backspace}^{Right}
}
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
	Run, "https://tdalon.blogspot.com/2020/12/open-multiple-microsoft-teams-instances.html"
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
Teams_Users2Excel(TeamLink:=""){
; Input can be sGroupId or Team link
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2020/08/teams-users2excel.html"
	return
}

Domain := People_GetDomain()
If (Domain ="") {
    MsgBox 0x10, Teams Shortcuts: Error, No Domain defined!
    return
}

RegRead, TeamsPowerShell, HKEY_CURRENT_USER\Software\PowerTools, TeamsPowerShell
If (TeamsPowerShell := !TeamsPowerShell) {
    sUrl := "https://tdalon.blogspot.com/2020/08/teams-powershell-setup.html"
    Run, "%sUrl%" 
    MsgBox 0x1024,People Connector, Have you setup Teams PowerShell on your PC?
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
    If RegExMatch(sName,".* \| Microsoft Teams, Main Window$") {
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
return TeamsMainWinId

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_GetMeetingWindow(useFindText:="" , restore := True){
; Syntax: Teams_GetMeetingWindow(useFindText:="" , restore := True)
;   restore is only used if FindText is used. (without FindText Teams windows are not activated.)
;   useFindText is taken from Settings PowerTools_GetParam("TeamsMeetingWinUseFindText") by default
; See implementation explanations here: 
;   https://tdalon.blogspot.com/2021/04/ahk-get-teams-meeting-window.html
;   https://tdalon.blogspot.com/2020/10/get-teams-window-ahk.html

If (useFindText="")
    useFindText := PowerTools_GetParam("TeamsMeetingWinUseFindText") ; Default 1

If (useFindText) {
    If (restore)
        WinGet, curWinId, ID, A
Else
    restore := False ; no need to restore if FindText is not used
}

WinGet, Win, List, ahk_exe Teams.exe
TeamsMainWinId := Teams_GetMainWindow()
TeamsMeetingWinId := PowerTools_RegRead("TeamsMeetingWinId")

; Shortcut if same Meeting as previous function call
If (useFindText) and WinExist("ahk_id " TeamsMeetingWinId) {
    If !(Teams_FindText("Resume")) and (Teams_FindText("Leave")) {
        return TeamsMeetingWinId
    } 
}

WinCount := 0

Loop %Win% {
    WinId := Win%A_Index%
    If (WinId = TeamsMainWinId) { ; Exclude Main Teams Window 
        ;WinGetTitle, Title, % "ahk_id " WinId
        ;MsgBox %Title%
        Continue
    }
    WinGetTitle, Title, % "ahk_id " WinId  
    
    IfEqual, Title,, Continue
    Title := StrReplace(Title," | Microsoft Teams","")
    If RegExMatch(Title,"^[^\s]*\s?[^\s]*,[^\s]*\s?[^\s]*$") or RegExMatch(Title,"^[^\s]*\s?[^\s]*,[^\s]*\s?[^\s]*\([^\s\(\)]*\)$") ; Exclude windows with , in the title (Popped-out 1-1 chat) and max two words before , Name, Firstname               
        Continue
    If RegExMatch(Title,"^Microsoft Teams Call in progress*") or RegExMatch(Title,"^Microsoft Teams Notification*") or RegExMatch(Title,"^Screen sharing toolbar*")
        Continue
    
    WinList .= ( (WinList<>"") ? "|" : "" ) Title "  {" WinId "}"
    WinCount++
 
} ; End Loop

If (WinCount = 0)
    return

If (WinCount = 1) { ; only one other window
    RegExMatch(WinList,"\{([^}]*)\}$",WinId)
    PowerTools_RegWrite("TeamsMeetingWinId",WinId1)
    return WinId1
}

If (useFindText) {
    WinCount1 := 0
    WinCount2 := 0
    Loop, Parse, WinList, | 
    {
        RegExMatch(A_LoopField,"\{([^}]*)\}$",WinId)
    
        WinActivate, ahk_id %WinId1%
        
        ; Final check - exclude window with Resume element = On hold meetings
        If (ok:=Teams_FindText("Resume")) {
            Continue
        } 
        WinListFT1 .= ( (WinListFT1<>"") ? "|" : "" ) A_LoopField
        WinCount1++
        ; Exclude window with no Leave element
        If !(ok:=Teams_FindText("Leave")) {
            Continue
        } 
        WinCount2++
        WinListFT2 .= ( (WinListFT2<>"") ? "|" : "" ) A_LoopField
    } ; End Loop

    If (WinCount2 = 1) { ; only one other window
        PowerTools_RegWrite("TeamsMeetingWinId",WinId1)
        If (restore)
            WinActivate, ahk_id %curWinId%
        return WinId1
    }

    If (WinCount1 = 1) { ; only one other window
        RegExMatch(WinListFT1,"\{([^}]*)\}$",WinId)
        PowerTools_RegWrite("TeamsMeetingWinId",WinId1)
        If (restore)
            WinActivate, ahk_id %curWinId%
        return WinId1
    }
    TrayTip, Could not find Meeting Window! , FindText failed to identify a meeting Window.,,0x2

    WinList := WinListFT1
} ; eo if useFindText

LB := WinListBox("Teams: Meeting Window", "Select your current Teams Meeting Window:" , WinList)
RegExMatch(LB,"\{([^}]*)\}$",WinId)
PowerTools_RegWrite("TeamsMeetingWinId",WinId1)
If (restore)
    WinActivate, ahk_id %curWinId%
return WinId1

} ; eofun
; -------------------------------------------------------------------------------------------------------------------


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
Teams_NewConversation(){
; Using hotkeys - broken
;SendInput ^{f6} ; Activate posts tab https://support.microsoft.com/en-us/office/use-a-screen-reader-to-explore-and-navigate-microsoft-teams-47614fb0-a583-49f6-84da-6872223e74a0#picktab=windows
; workaround will flash the search bar if posts/content panel already selected but works now even if you have just selected the channel on the left navigation panel
;SendInput {Esc} ; in case expand box is already opened
SendInput !+c ;  compose box alt+shift+c: necessary to get second hotkey working (regression with new conversation button)
sleep, 500
SendInput ^+x ; expand compose box ctrl+shift+x (does not work anymore immediately)
sleep, 500
SendInput +{Tab} ; move cursor back to subject line via shift+tab
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------
Teams_NewConversation2(){
; Using FindText
ok := Teams_Click("NewConversation")
If !(ok) {
    TrayTip Teams Shortcuts: ERROR, FindText failed!
    Run, "https://tdalon.github.io/ahk/Teams-Meeting-Reactions"
    return
}
Delay := PowerTools_GetParam("TeamsClickDelay")
Sleep %Delay% 

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
Teams_Share(ShareMode := 2){
WinId := Teams_GetMeetingWindow()
If !WinId ; empty
    return
SysGet, MonitorCount, MonitorCount	; or try:    SysGet, var, 80
If (MonitorCount > 1)
    IsActive := WinActive("ahk_id " . WinId)
Else 
    IsActive := False ; Set to False if only one monitor is used
WinActivate, ahk_id %WinId%


ok := Teams_FindText("MeetingActionShare")
IsSharing := !(ok)
If (ShareMode = 1) and (IsSharing) ; already sharing
    Return

SendInput ^+e ; ctrl+shift+e 
; Wait for share tray to open
Delay := PowerTools_GetParam("TeamsShareDelay")
Sleep %Delay% 
;while !Teams_ImageSearch("teams_include_computer_sound.png") ; ImageSearch does not work for this element
;    Sleep, 500

If !(IsSharing)
    SendInput {Tab}{Enter} ; Select first screen
If (IsActive) and !(IsSharing) { ; only if IsActive and not sharing
; Bring back meeting window (multiple screen setup) - it is else minimized while sharing by Teams
    WinWaitNotActive, ahk_id %WinId%
    
    WinActivate, ahk_id %WinId% ; requires activate to move
    MonIndex := Monitor_GetMonitorIndex(WinId)
    If (MonIndex=1) {
        SendInput #+{Right}; Win+Shift+Right Arrow
        ;Monitor_WinMove(WinId)    
    }   
}
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------
Teams_ClearCache(){
If GetKeyState("Ctrl") {
    sUrl := "https://tdalon.blogspot.com/2021/01/teams-clear-cache.html"
    Run, "%sUrl%"
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
Teams_Mute(State := 2){
; State: 
;    0: mute off, unmute
;    1: mute on
;    2*: (Default): Toggle mute state

; N.B.: ControlSend does not work for Teams window e.g. ControlSend, ahk_parent, ^+m, ahk_id %WinId% - Teams window must be active
; hotkey does not work anymore from main window
/* If (State = 2) ; Toggle mute
    WinId := Teams_GetMainWindow() ; mute hotkey can be run from Main window - prefer main window because it is easier and more robust - 
Else {
    WinId := Teams_GetMeetingWindow() ; need to get meeting window to check mute state
} 
*/

WinId := Teams_GetMeetingWindow()
If !WinId ; empty
    return
;MsgBox % WinId
WinGet, curWinId, ID, A
WinActivate, ahk_id %WinId%

If (State <> 2)
    IsMuted := !(Teams_FindText("Mute"))
MsgBox %IsMuted%
Switch State 
{
    Case 0:
        
        Tooltip("Teams Unmute Mic...") 
        If !IsMuted
            return
        
    Case 1:
        If IsMuted
            return
        Tooltip("Teams Mute Mic...") 
    Case 2:
        Tooltip("Teams Toggle Mute Mic...") 
}

SendInput ^+m ;  ctrl+shift+m 
Sleep 500 ; pause before reactivating previous window
WinActivate, ahk_id %curWinId%

} ; eofun


Teams_Leave() {
    WinId := Teams_GetMeetingWindow()
    If !WinId ; empty
        return
    WinActivate, ahk_id %WinId%
    SendInput ^+b ; ctrl+shift+b
    return
}

; -------------------------------------------------------------------------------------------------------------------
Teams_PushToTalk(){

WinId := Teams_GetMeetingWindow() ; mute can be run from Main window

If !WinId ; empty
    return

KeyName := A_PriorKey
If (KeyName = "m") or (KeyName = "^") or (KeyName = "!") {
    MsgBox Error: hotkey conflict with native Ctrl+Shift+M. Choose another combination.
    return
}

WinGet, curWinId, ID, A
WinActivate, ahk_id %WinId%
IsMuted := !(Teams_FindText("Mute"))

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
}

; For Menu callback, remove ending Hotkey and blanks
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


Teams_Video(){
; Toggle Video on/off
WinId := Teams_GetMeetingWindow()
If !WinId ; empty
    return
WinGet, curWinId, ID, A
WinActivate, ahk_id %WinId%
SendInput ^+o ; toggle video Ctl+Shift+o
Sleep 500
;SendInput ^+p ; toggle background blur
WinActivate, ahk_id %curWinId%
Tooltip("Teams Toggle Video...") 
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
WinId := Teams_GetMeetingWindow()
If !WinId ; empty
    return
Tooltip("Teams Toggle Raise Hand...") 
WinGet, curWinId, ID, A
WinActivate, ahk_id %WinId%
SendInput ^+k ; toggle video Ctl+Shift+k
WinActivate, ahk_id %curWinId%
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_React(sReaction) {
; sReaction can be Like | Applause| Heart | Laugh
WinId := Teams_GetMeetingWindow()
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

; -------------------------------------------------------------------------------------------------------------------
Teams_MeetingReaction(Reaction) {
; Reaction can be Like | Applause| Heart | Laugh
WinId := Teams_GetMeetingWindow()
If !WinId ; empty
    return

WinGet, curWinId, ID, A
MouseGetPos , MouseX, MouseY
WinActivate, ahk_id %WinId%

WinGetPos , , ,WinWidth, WinHeight, A

ok := Teams_Click("MeetingReactions")
If !(ok) {
    TrayTip Teams Meeting Reaction: ERROR, FindText failed!
    Run, "https://tdalon.github.io/ahk/Teams-Meeting-Reactions"
    return
}

Delay := PowerTools_GetParam("TeamsClickDelay")
Sleep %Delay% 
ok := Teams_Click("MeetingReaction" . Reaction)

; Restore previous window and mouse position
WinActivate, ahk_id %curWinId%
MouseMove, MouseX, MouseY
If (ok)
    Tooltip("Teams Meeting Reaction: " . Reaction,1000) 
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
Teams_MeetingAction(id, restore:= False){
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
    Tooltip("Teams Meeting Reaction: " . Reaction,1000) 
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
