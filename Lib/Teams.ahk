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

Teams_GetVer() { ; NOT WORKING
; Get Client Version information
If Teams_IsNew() {
    Loop, Reg, HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages\ 
        {
            MsgBox % A_LoopRegName
            If RegExMatch(A_LoopRegName,"U)^MSTeams_(.*)_",sMatch)
                return sMatch1
        }
}

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Teams_BackgroundOpenLibrary() {
    If GetKeyState("Ctrl") {
        Teamsy_Help("cbg")
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
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Teams_BackgroundGetFolder(doOpen := true){
; BackgroundDir := Teams_BackgroundGetFolder(doOpen := true)
    If GetKeyState("Ctrl") {
        Teamsy_Help("cbg")
        return
    }

    If Teams_IsNew() { ; Teams New
    ; https://techcommunity.microsoft.com/t5/microsoft-teams-public-preview/backgrounds/m-p/3782690
        EnvGet, LOCALAPPDATA, LOCALAPPDATA
        BackgroundDir = %LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\Uploads
    } Else { ; Teams Classic
        BackgroundDir = %A_AppData%\Microsoft\Teams\Backgrounds\Uploads
    }

    If !FileExist(BackgroundDir)
        FileCreateDir, %BackgroundDir%
    If (doOpen)
        Run, %BackgroundDir%
    return BackgroundDir

} ; eofun
; -------------------------------------------------------------------------------------------------------------------


; -------------------------------------------------------------------------------------------------------------------
Teams_BackgroundImport(srcDir:=""){
    ; https://techcommunity.microsoft.com/t5/microsoft-teams-public-preview/backgrounds/m-p/3782690
    ; Default will migrate Classic Client location to New Teams Background location
    
    If GetKeyState("Ctrl") {
        Teamsy_Help("bgi")
        return
    }
    If !Teams_IsNew() 
        {
            return
        }
        
    If (srcDir = "") { ; prompt user to select directory
        srcDir = %A_AppData%\Microsoft\Teams\Backgrounds\Uploads
        ;FileSelectFolder, srcDir , ,0 , Select folder from which to import background images:
        sf := SelectFolder(srcDir,"Select folder from which to import background images:") 
        If ErrorLevel
            return
        srcDir := sf.SelectedDir
    }

    EnvGet, LOCALAPPDATA, LOCALAPPDATA
    NewBackgroundDir = %LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\Uploads
    If !FileExist(NewBackgroundDir)
        FileCreateDir, %NewBackgroundDir%

    ; Generate array of MD5 checksums for background files
    ArrayCount := 0
    MD5Array := []
    Loop, Files, %NewBackgroundDir%\*
    {
        If InStr(A_LoopFileName,"_thumb.")
            Continue
        ArrayCount += 1
        MD5Array[ArrayCount] := HashFile(A_LoopFileFullPath,"MD5")
    }
        
    If !pToken := Gdip_Startup()
        {
            MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
            return
        }

    BgCount := 0
    BgCopiedCount :=0
    Loop, Files, %srcDir%\*
    {
        ; Skip files ending with _thumb.
        If InStr(A_LoopFileName,"_thumb.")
            Continue
        BgCount +=1
        ; Check if file already exist
        srcMD5 := HashFile(A_LoopFileFullPath,"MD5")
        
        If HasVal(MD5Array,srcMD5) 
            Continue
        BgCopiedCount +=1    
        ; Generate GUID
        GUID := CreateGUID()
        GUID := SubStr(GUID,2,-1)
        StringLower, GUID, GUID

        RegExMatch(A_LoopFileName,"\.([^\.]*$)",FileExt)
        FileCopy, %A_LoopFileFullPath%, %NewBackgroundDir%\%GUID%.%FileExt1%

        ; Save Thumbnail 
        destThumbFile = %NewBackgroundDir%\%GUID%_thumb.%FileExt1%
        srcThumbFile := RegExReplace(A_LoopFileFullPath,"\.*$") . "_thumb." . FileExt1
        If FileExist(srcThmbFile) {
            FileCopy, %srcThumbFile%, %destThumbFile%
        } Else {
            ; Create Thumbnails using Gdip 
            pBitmap := Gdip_CreateBitmapFromFile(A_LoopFileFullPath)
            pThumbnail := Gdip_GetImageThumbnail(pBitmap, 278, 159)
            Gdip_SaveBitmapToFile(pThumbnail, destThumbFile)
        }
    }

    TrayTip Backgrounds imported! , %BgCopiedCount%/%BgCount% backgrounds were copied!
    Run, %NewBackgroundDir%

    If !(pBitmapt = "") {
        Gdip_DisposeImage(pBitmap) 
        Gdip_DisposeImage(pThumbnail) 
    }
    Gdip_Shutdown(pToken) 


    ; Merge/save BackgroundNames in PowerTools.ini
    


        
} ; eofun
; -------------------------------------------------------------------------------------------------------------------


Teams_BackgroundName2File(Name:=""){
    ; Name := Teams_BackgroundFile2Name(File:="",mode:="r")
    ; If no File input, user will be prompted to select an image file
    ; If mode is "w" (write) user can change the name value if already existing

    If GetKeyState("Ctrl") {
        Teamsy_Help("bg")
        return
    }

    ; Get Background File
    If (Name = "") {
        BackgroundDir := Teams_BackgroundGetFolder(false)
        FileSelectFile, File, 1, %BackgroundDir%,Select a background image
        If ErrorLevel
            return
        Px := PathX(File)
        File := Px.File
        File := StrReplace(File, "_thumb.",".") ; in case thumbnail was selected
    } Else If (Name ="no") {
        File := "No background effect" ; TODO lang
        return File
    } Else {
        If !FileExist("PowerTools.ini") {
            PowerTools_ErrDlg("No PowerTools.ini file found!")
            return
        }
    
        IniRead, TeamsBackgroundNames, PowerTools.ini,Teams,TeamsBackgroundNames
        If (TeamsBackgroundNames="ERROR") { 
            PowerTools_ErrDlg("TeamsBackgroundNames key not found in PowerTools.ini file [Teams] section!")
            return
        }
      
        JsonObj := Jxon_Load(TeamsBackgroundNames)
        For i, bg in JsonObj 
        {
            name_i := bg["name"]
            If RegExMatch(name_i,"^" . StrReplace(Name,"*",".*")) {
                File := bg["file"]
                break
            }
        }
    }

    return File

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Teams_MeetingBackgroundSet(File:="") {

    WinId := Teams_GetMeetingWindow() 
    If !WinId ; empty
        return

    If (File = "") or !InStr(File,".") ; no file name with extension bug Background simple name
        File := Teams_BackgroundName2File(File)

    If (File = "")
        return

    WinGet, curWinId, ID, A
    WinActivate, ahk_id %WinId%

    UIA := UIA_Interface()  
    TeamsEl := UIA.ElementFromHandle(WinId)

    ; Check if rail pane already opened - if yes close it
    El :=  TeamsEl.FindFirstBy("AutomationId=rail-header-close-button")  
    El.Click()
    
    TeamsEl.WaitElementNotExist("AutomationId=rail-header-close-button")
    El.Click()
    

    ; Click on More -> Video Effects
    El :=  TeamsEl.FindFirstBy("AutomationId=callingButtons-showMoreBtn")  
    El.Click() 
    El :=  TeamsEl.WaitElementExist("AutomationId=video-effects-and-avatar-button",,,,1000)  
    El.Click() 

    Name := Teams_GetLangName("ShowAllBackgrounds","Show all background effects")
    El :=  TeamsEl.WaitElementExistByName(Name,,1,false,1000)  ; 1 match mode start with
    If !El {
        ;error notification
        PowerTools_ErrDlg("UIA Element 'Show all backgrounds' not found!")
        WinActivate, ahk_id %curWinId%
        return
    }
    El.Click()

    ;AutomationId=video-effects-and-avatar-button, Name: Hide or Show
    
    El :=  TeamsEl.WaitElementExistByNameAndType(File,"Checkbox",,2,,1000) ; TODO Lang  ; matchmode contains

    ;Type: Checkbox, Name <FileName> background effect
    If !El {
        ;error notification
        PowerTools_ErrDlg("UIA Element not found!")
        WinActivate, ahk_id %curWinId%
        return
    }

    El.Click()
    ; Button Name=Apply
    El :=  TeamsEl.FindFirstByNameAndType("Apply","Button",,3) ; TODO Lang ; matchmode exact
    If El
        El.Click()

    ; Close video effects
    El :=  TeamsEl.FindFirstBy("AutomationId=rail-header-close-button")  
    El.Click()

    WinActivate, ahk_id %curWinId%

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Teams_MeetingBackgroundSettings() {
; Open Background Settings

    WinId := Teams_GetMeetingWindow() 
    If !WinId ; empty
        return
    
    WinActivate, ahk_id %WinId%

    UIA := UIA_Interface()  
    TeamsEl := UIA.ElementFromHandle(WinId)

    ; Check if rail pane already opened - if yes close it
    El :=  TeamsEl.FindFirstBy("AutomationId=rail-header-close-button")  
    El.Click()
    
    TeamsEl.WaitElementNotExist("AutomationId=rail-header-close-button")
    El.Click()
    

    ; Click on More -> Video Effects
    El :=  TeamsEl.FindFirstBy("AutomationId=callingButtons-showMoreBtn")  
    El.Click() 
    El :=  TeamsEl.WaitElementExist("AutomationId=video-effects-and-avatar-button",,,,1000)  
    El.Click() 

    Name := Teams_GetLangName("ShowAllBackgrounds","Show all background effects")
    El :=  TeamsEl.WaitElementExistByName(Name,,1,false,1000)  ; 1 match mode start with
    If !El {
        ;error notification
        PowerTools_ErrDlg("UIA Element 'Show all backgrounds' not found!")
        WinActivate, ahk_id %curWinId%
        return
    }
    El.Click()

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Teams_BackgroundFile2Name(File:="",mode:="r"){
    ; Name := Teams_BackgroundFile2Name(File:="",mode:="r")
    ; If no File input, user will be prompted to select an image file
    ; If mode is "w" (write) user can change the name value if already existing

    If GetKeyState("Ctrl") {
        Teamsy_Help("bgs")
        return
    }

    If (File = "") { ; Select File 
        BackgroundDir := Teams_BackgroundGetFolder(false)
        FileSelectFile, File, 1, %BackgroundDir%,Select a background image
        If ErrorLevel
            return
        Px := PathX(File)
        File := Px.File
        File := StrReplace(File, "_thumb.",".") ; in case thumbnail was selected
    }


    If !FileExist("PowerTools.ini") {
        ;PowerTools_ErrDlg("No PowerTools.ini file found!")
        return
    }

    IniRead, TeamsBackgroundNames, PowerTools.ini,Teams,TeamsBackgroundNames
    ;If (TeamsBackgroundNames="ERROR") { 
    ;    PowerTools_ErrDlg("TeamsBackgroundNames key not found in PowerTools.ini file [Teams] section!")
    ;    return
    ;}

    If (TeamsBackgroundNames="ERROR") or (TeamsBackgroundNames="[]") or (TeamsBackgroundNames="")
        Goto InputName

    JsonObj := Jxon_Load(TeamsBackgroundNames)
    For i, bg in JsonObj 
    {
        file_i := bg["file"]
        
        If (File = file_i) {
            Name := bg["name"]
            break
        }
    }
    If !(Name = "") and (mode = "r")
        return Name

    ; Input name
    InputName:
    ;InputBox, OutputVar , Title, Prompt, Hide, Width, Height, X, Y, Locale, Timeout, Default
    InputBox, NewName , Background Name, Enter name:,,200,125,,,,,%Name%
    If ErrorLevel
        return
    If (NewName = Name) ; no changes
        return Name
    Name := NewName
    If (TeamsBackgroundNames="ERROR") or (TeamsBackgroundNames="[]") or (TeamsBackgroundNames="") 
        NewTeamsBackgroundNames = [{"file":"%File%","name":"%Name%"}]
    Else {
        sPat = "file":"%File%","name":"([^"]*)"
        If RegExMatch(TeamsBackgroundNames,sPat,sMatch) { ; replace
            sNeedle = "file":"%File%","name":"%sMatch1%"
            sRep = "file":"%File%","name":"%Name%"
            NewTeamsBackgroundNames := StrReplace(TeamsBackgroundNames,sNeedle,sRep)
        } Else { ; append
            NewTeamsBackgroundNames := SubStr(TeamsBackgroundNames,1,-1) 
            NewTeamsBackgroundNames = %NewTeamsBackgroundNames%,{"file":"%File%","name":"%Name%"}]
        }
    }
    
    ; Save to Ini file
    IniWrite, %NewTeamsBackgroundNames%, PowerTools.ini,Teams,TeamsBackgroundNames
    return Name

} ; eofun
; -------------------------------------------------------------------------------------------------------------------


CreateGUID()
;https://www.autohotkey.com/boards/viewtopic.php?f=6&t=4732
{
    VarSetCapacity(pguid, 16, 0)
    if !(DllCall("ole32.dll\CoCreateGuid", "ptr", &pguid)) {
        size := VarSetCapacity(sguid, (38 << !!A_IsUnicode) + 1, 0)
        if (DllCall("ole32.dll\StringFromGUID2", "ptr", &pguid, "ptr", &sguid, "int", size))
            return StrGet(&sguid)
    }
    return ""
}
; -------------------------------------------------------------------------------------------------------------------


Teams_Emails2ChatDeepLink(sEmailList, askOpen:= true){
; Copy Chat Link to Clipboard and ask to open
sLink := "https://teams.microsoft.com/l/chat/0/0?users=" . StrReplace(sEmailList, ";",",") 
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
sLink := "https://teams.microsoft.com/l/chat/0/0?users=" . StrReplace(sEmailList, ";",",") 
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
sLink := "https://teams.microsoft.com/l/meeting/new?attendees=" . StrReplace(sEmailList, ";",",") 
Teams_OpenUrl(sLink) 
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Link2Fav(sUrl:="",FavsDir:="",sFileName :="") { ; @fun_teams_link2fav@
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
    If RegExMatch(sUrl,"https://teams.microsoft.com/l/channel/[^/]*/([^/]*)\?.*",sChannelName) 
	    linktext = %sChannelName1% (Channel)
    Else
        linktext = Team Name (Team)

    InputBox, sUrl , Teams Fav Target, Paste Teams Link:, , 640, 125,,,,, %linktext%
	If ErrorLevel
		return
}

; FavsDir (folder does not end with filesep)
If !(FavsDir) {
    sKeyName := "TeamsFavsDir"
    RegRead, StartingFolder, HKEY_CURRENT_USER\Software\PowerTools, %sKeyName%
    FileSelectFolder, FavsDir , *%StartingFolder%, ,Select folder to store your Teams Shortcut:
    If ErrorLevel
        return
}

; For Message Link, choose if you want to link to the message or the containing Chat
If (RegExMatch(sUrl,"(https|msteams)://teams\.microsoft\.com/l/message/(.*)@thread\.v2/",sMatch)) {
    ; https://teams.microsoft.com/l/team/19:c1471a18bae04cf692b8da7e9738df3e@thread.skype/conversations?groupId=56bc81d8-db27-487c-8e4f-8d5ea9058663&tenantId=xxx    
    OnMessage(0x44, "OnTeamsLinkTypeMessageMsgBox")
		MsgBox 0x24, Message Link, Select to what you want to link:
		OnMessage(0x44, "")
		IfMsgBox, No
			sUrl := StrReplace(sMatch,"/l/message/","/l/chat/")
}

; FileName
If !(sFileName) {
    sFileName := Teams_Link2Text(sUrl,silent:=true)
    InputBox, sFileName , Teams Fav File name, Enter the File name:, , 300, 125,,,,, %sFileName%
    If ErrorLevel
        return
}

sFile := FavsDir . "\" . sFileName . ".url"


;FileDelete %sFile%
IniWrite, %sUrl%, %sFile%, InternetShortcut, URL

; Add icon file:
If Teams_IsNew()
    TeamsExe := Teams_GetExe()
Else
    TeamsExe = C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\Update.exe
IniWrite, %TeamsExe%, %sFile%, InternetShortcut, IconFile
IniWrite, 0, %sFile%, InternetShortcut, IconIndex

; Save FavsDir to Settings in registry
SplitPath, sFile, sFileName, FavsDir
PowerTools_RegWrite("TeamsFavsDir",FavsDir)
} ; eofun


; -------------------------------------------------------------------------------------------------------------------
OnTeamsLinkTypeMessageMsgBox() {
    DetectHiddenWindows, On
    Process, Exist
    If (WinExist("ahk_class #32770 ahk_pid " . ErrorLevel)) {
		ControlSetText Button1, Message
		ControlSetText Button2, Parent Chat
    }
}

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
        ;Run, %A_LoopFileFullPath%
        OpenFavUrl(A_LoopFileFullPath)
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

OpenFavUrl(File) {
    ; Opens Internet Shortcuts .url files and close left-over browser window
    ; Assumes Edge Browser is installed. Window History won't be changed
    
    ; Extract url from Shortcut
    IniRead, sUrl, %File%, InternetShortcut, URL
    Teams_OpenUrl(sUrl)
    
} ; eofun

Teams_OpenUrl(sUrl) {
; Open Teams Url without Leftover
    TeamsExe := Teams_GetExeName()
    EdgeWinId := WinActive("ahk_exe msedge.exe") 
    Run, msedge.exe "%sUrl%" " --new-window"
    If EdgeWinId   
        WinWaitNotActive, ahk_id %EdgeWinId%
    Else
        WinWaitActive, ahk_exe msedge.exe
    NewEdgeWinId := WinExist("ahk_exe msedge.exe")
    WinWaitActive, ahk_exe %TeamsExe%,,2
    If !ErrorLevel {    
        If WinExist("ahk_id " . NewEdgeWinId) {
            WinActivate
            Send ^w ; Close leftover browser window
        }
    }
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
Teams_Link2Fav(sUrl,FavsDir,"Chat " . sName)

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

Teams_IsNew(){ ; @fun_teams_isnew@
; IsNew := Teams_IsNew()
; return true or false depending if Teams New Client is installed and Classic Teams running
    static IsNew
    If !(IsNew = "") {
        ;MsgBox IsNew:%IsNew%
        return IsNew
    }
        
    fExe = C:\Users\%A_UserName%\AppData\Local\Microsoft\WindowsApps\ms-teams.exe
    If !(FileExist(fExe)) { ; New Teams not installed
        IsNew := False
        return IsNew
    }

    Process, Exist, ms-teams.exe
	If Not (ErrorLevel= 0) {
		;Classic Teams
		IsNew:=True
        return IsNew
	}

     ; Check if a Teams process is running ; need to disable Microsoft Teams classic in Startup to work at startup
     Process, Exist, Teams.exe
     If Not (ErrorLevel= 0) {
         ;Classic Teams
         IsNew := False
         return IsNew
     }

    ; Possibility to overwrite in ini file if Teams is not started
    If FileExist("PowerTools.ini") {
        IniRead, IniIsNew, PowerTools.ini,Teams,TeamsIsNew
        If !(IniIsNew="ERROR") {
            IsNew := IniIsNew
            return IsNew
        } 
    }
    IsNew:=True
    return IsNew
} ;eofun

; ----------------------------------------------------------------------
Teams_GetExe() {
    If Teams_IsNew()
        ; New Teams in WindowsApp C:\Users\%A_UserName%\AppData\Local\Microsoft\WindowsApps
        fExe = C:\Users\%A_UserName%\AppData\Local\Microsoft\WindowsApps\ms-teams.exe
    Else { ; Classic Team Client
        ;EnvGet, userprofile , userprofile
        ;fExe = %userprofile%\AppData\Local\Microsoft\Teams\current\Teams.exe ;C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\Update.exe
        fExe = C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\current\Teams.exe
    }
    return fExe
}

Teams_RunExe(){
If Teams_IsNew() {
    ; New Teams in WindowsApp C:\Users\%A_UserName%\AppData\Local\Microsoft\WindowsApps
    fExe := "ms-teams.exe"
    Run, %fExe%
} Else { ; Classic Team Client
    ;fExe = C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\current\Teams.exe
    ;fExe = C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\Update.exe --processStart "Teams.exe"
    Run, "C:\Users\%A_UserName%\AppData\Local\Microsoft\Teams\Update.exe" --processStart "Teams.exe"
}
} ;eofun
; ----------------------------------------------------------------------
; ----------------------------------------------------------------------
Teams_GetExeName(IsNew:=""){
If Teams_IsNew()
    return "ms-teams.exe"
Else
    return "Teams.exe"
} ;eofun
; ----------------------------------------------------------------------

Teams_SharingControlBar(mode:="-") {
    ; Teams_SharingControlBar(mode:="-")
    ; mode is '+' pr '-'. Empty will toggle
    If GetKeyState("Ctrl") {
        Teamsy_Help("sb")
        return
    }
    
    Lang := Teams_GetLang()
    Prop := "SharingControlBar"
    Name := Teams_GetLangName(Prop,"Sharing control bar",Lang)
    If (Name="") 
        return
    
    TeamsExe := Teams_GetExeName()
    wTitle = %Name% ahk_exe %TeamsExe% 
    Switch mode {
        Case "-":
            WinHide, %wTitle%
            return
        Case "+":
            WinShow, %wTitle%
            return
        Default:
            If WinExist(wTitle)
                WinHide, %wTitle%
            Else
                WinShow, %wTitle% 
    }
} ; eofun

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
MsgBox % sSelectionHtml

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
TeamsExe := Teams_GetExeName()
If WinActive("ahk_exe " . TeamsExe) { ; SendMention
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
If WinActive("ahk_exe " . TeamsExe) {
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
; Shift+F10 Open Context Menu->Copy Link

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



Teams_Link2Text(sLink,silent:= false){

; Link clean-up
;sLink := StrReplace(sLink,"%2520"," ") ; spaces in Channel Link
;sPat = (?:https|msteams)://teams.microsoft.com/[^>"]*
;RegExMatch(sLink,sPat,sLink)
;sLink := uriDecode(sLink)

; Link to Teams Channel
; example: https://teams.microsoft.com/l/channel/19%3a16ff462071114e31bd696aa3a4e34500%40thread.skype/DOORS%2520Attributes%2520List?groupId=cd211b48-2e8b-4b60-b5b0-e584a0cf30c0&tenantId=xxx
If (RegExMatch(sLink,"U)(?:https|msteams)://teams\.microsoft\.com/l/channel/[^/]*/([^/]*)\?groupId=(.*)&",sMatch)) {
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
} Else If (RegExMatch(sLink,"(?:https|msteams)://teams\.microsoft\.com/l/team/.*groupId=([^&]*)",sMatch)) {
; https://teams.microsoft.com/l/team/19:c1471a18bae04cf692b8da7e9738df3e@thread.skype/conversations?groupId=56bc81d8-db27-487c-8e4f-8d5ea9058663&tenantId=xxx    
    sTeamName := Teams_GetTeamName(sMatch1)
    If (!sTeamsName)
        sDefText = %sTeamName% (Team)
    Else
        sDefText = Link to Teams Team
    If silent
        return sDefText  
    InputBox, linktext , Display Link Text, Enter Team name:,,640,125,,,,, %sDefText%
	If ErrorLevel ; Cancel
		return

} Else If (RegExMatch(sLink,"(https|msteams)://teams\.microsoft\.com/l/message/",sMatch)) {
; https://teams.microsoft.com/l/team/19:c1471a18bae04cf692b8da7e9738df3e@thread.skype/conversations?groupId=56bc81d8-db27-487c-8e4f-8d5ea9058663&tenantId=xxx    
    linktext = Message (Message)
    If silent
        return linktext 
    InputBox, linktext , Display Link Text, Enter Text:,,640,125,,,,, %linktext%
    If ErrorLevel ; Cancel
        return

} Else If (RegExMatch(sLink,"(?:https|msteams)://teams\.microsoft\.com/l/chat/(.*)",sMatch)) {
    ; https://teams.microsoft.com/l/team/19:c1471a18bae04cf692b8da7e9738df3e@thread.skype/conversations?groupId=56bc81d8-db27-487c-8e4f-8d5ea9058663&tenantId=xxx    
        If InStr(sMatch1,"19:meeting")
            sDefText = MeetingName (Meeting Chat)
        Else
            sDefText = GroupName (Group Chat)
        If silent
            return sDefText 
        InputBox, linktext , Display Link Text, Enter Text:,,640,125,,,,, %sDefText%
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
TeamsExe := Teams_GetExeName()
sCmd = Update.exe --processStart "%TeamsExe%"

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

Teams_GetMainWindow(){ ; @fun_teams_getmainwindow@
; See implementation explanations here: https://tdalon.blogspot.com/get-teams-window-ahk
; Syntax: hWnd := Teams_GetMainWindow()


TeamsMainWinId := PowerTools_RegRead("TeamsMainWinId") ; use registry vs. static to make it persistent to tool restart
If WinExist("ahk_id " . TeamsMainWinId) 
    return TeamsMainWinId
TeamsExe := Teams_GetExeName()

WinGet, WinCount, Count, ahk_exe %TeamsExe%

If (WinCount = 0)
    GoTo, StartTeams

; Get main window via Acc Window Object Name - only for Classic Client
If !Teams_IsNew() { ; Classic Teams

    ; If only one window opened- main window (no true for New Teams where main window can be closed)
    ; when virtuawin is running Teams main window can be on another virtual desktop = hidden
    Process, Exist, VirtuaWin.exe
    VirtuaWinIsRunning := ErrorLevel
    If (WinCount = 1) and Not (VirtuaWinIsRunning) {
        TeamsMainWinId := WinExist("ahk_exe " . TeamsExe)
        RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, TeamsMainWinId, %TeamsMainWinId%
        return TeamsMainWinId
    }

    WinGet, WinId, List,ahk_exe %TeamsExe%
    Loop, %WinId%
    {
        hWnd := WinId%A_Index%
        oAcc := Acc_Get("Object","4",0,"ahk_id " hWnd)
        sName := oAcc.accName(0)
        If RegExMatch(sName,".* \| Microsoft Teams, Main Window$") { ; works also for other lang - does not work with Teams New
            TeamsMainWinId := hWnd
            RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, TeamsMainWinId, %TeamsMainWinId%
            return TeamsMainWinId
        }
    }
} ; Teams Classic

; Fallback solution with minimize all windows and run exe

WinGet, curWinId, ID, A

If WinActive("ahk_exe " . TeamsExe) {
    GroupAdd, TeamsGroup, ahk_exe %TeamsExe%
    WinMinimize, ahk_group  TeamsGroup
} 

StartTeams: 
Teams_RunExe()
WinWaitActive, ahk_exe %TeamsExe%,,5
If ErrorLevel {
    TrayTip, Could not find Teams Main Window! , Timeout exceeded!,,0x2
    return
}
TeamsMainWinId := WinExist("A")

; restore previous active window
WinActivate, ahk_id %curWinId%
RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, TeamsMainWinId, %TeamsMainWinId%
return TeamsMainWinId


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
    SendInput {tab}
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
Teams_GetMeetingWindow(Minimize:=false,showTrayTip:=true){ ; @fun_teams_getmeetingwindow@
; Get active Teams Meeting window
; Syntax: 
;      hWnd := Teams_GetMeetingWindow(Minimize:=false,showTrayTip:=true) 
;   If window is not found, hwnd is empty
; Input Arguments:
;   Minimize: if true, minimized (aka Call in Progress) meeting window can be returned
;

; See implementation explanations here: 
;   https://tdalon.blogspot.com/2022/07/ahk-get-teams-meeting-win.html
; Does not require window to be activated
UIA := UIA_Interface()
TeamsExe := Teams_GetExeName()
WinGet, Win, List, ahk_exe %TeamsExe%
/* 
If restoreWin
    WinGet, curWinId, ID, A 
*/
Loop %Win% {
    WinId := Win%A_Index%
   ; WinActivate, ahk_id %WinId%
    TeamsEl := UIA.ElementFromHandle(WinId)

    If Teams_IsMeetingWindow(TeamsEl)  {
        If (!Minimize)
            If Teams_IsMinMeetingWindow(TeamsEl)
                Continue
        return WinId
    }
} ; End Loop

/* 
If (restoreWin)
    WinActivate, ahk_id %curWinId% ; restore win 
*/
If (showTrayTip)
    TrayTip, Could not find Meeting Window! , No active Teams meeting window found!,,0x2
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Teams_IsMeetingWindow(TeamsEl,ExOnHold:=true){
; Check if UIA Window is an active* meeting window
; By default exclude On-hold meetings (with Resume button)
If (ExOnHold) {
    Name := Teams_GetLangName("Resume","Resume")
    If (Name="") 
        return
}

If TeamsEl.FindFirstBy("AutomationId=microphone-button") {
   If (ExOnHold) { ; exclude On Hold meeting windows
        If TeamsEl.FindFirstByName(Name) ; Exclude On-hold meetings with Resume button 
            return false
    }
    return true
}
return false
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_IsMinMeetingWindow(TeamsEl) {
; Return true if the window is a minimized meeting window
; Check for button "Navigate back"
    Name := Teams_GetLangName("NavigateBack","Navigate back to call window.")
    If (Name="") 
        return 
    El := TeamsEl.FindFirstByNameAndType(Name, "button") ; 
    If El 
        return true
    Else 
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
Teams_ActivateMeetingWindow(){
; WinId := Teams_ActivateMainWindow()
    WinId:= Teams_GetMeetingWindow()
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
Teams_MeetingRecord(Mode := 2){
    ; Mode = 0 : stop
    ; Mode = 1 : start
    ; Mode = 2: toggle

    If GetKeyState("Ctrl") {
        Teamsy_Help("rec")
        return
    }
    WinId := Teams_GetMeetingWindow()
    If !WinId ; empty
        return

    UIA := UIA_Interface()
    TeamsEl := UIA.ElementFromHandle(WinId)


    WinGet, curWinId, ID, A
    WinActivate, ahk_id %WinId%


    ; Click on More 
    El :=  TeamsEl.FindFirstBy("AutomationId=callingButtons-showMoreBtn")  
     

    ;MsgBox % ShareEl.DumpAll() ; DBG
    If !El {
        TrayTip TeamsShortcuts: ERROR, More button not found!,,0x2
        return
    }
    El.Click()
   

    El := TeamsEl.WaitElementExist("AutomationId=RecordingMenuControl-id",,,,1000)
    If !El {
        TrayTip TeamsShortcuts: ERROR, Record control not found!,,0x2
        return
    }

    Send +{SPACE} ; Shift+Space will expand the first menu

    
    El := TeamsEl.WaitElementExist("AutomationId=recording-button",,,,1000)
    If !El {
        TrayTip TeamsShortcuts: ERROR, Record button not found!,,0x2
        return
    }

    Name := Teams_GetLangName("RecordStart","Start")
    If (Name="") 
        return
    
    IsRecording := !RegExMatch(El.Name,"^" . Name) 

    If (Mode = 1) and (IsRecording) ; already recording
        Return

    If (Mode = 0) and !(IsRecording) ; already not recording
        Return

    If (IsRecording)  { ; Stop recording
        SendInput {Enter}
        El := TeamsEl.WaitElementExistByNameAndType("Cancel","button",,,,1000) ; TODO Lang specific
        SendInput {Tab}{Enter}

    } Else { ; Start recording
        SendInput {Enter}
    }
    

   
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Teams_MeetingShare(ShareMode := 2){
    ; ShareMode = 0 : unshare
    ; ShareMode = 1 : share
    ; ShareMode = 2: toggle share

    If GetKeyState("Ctrl") {
        Teamsy_Help("sh")
        return
    }
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

    Lang := Teams_GetLang()
        
    Name := Teams_GetLangName("Share","Share",Lang)
    If (Name="") 
        return
    
    IsSharing := !RegExMatch(ShareEl.Name,"^" . Name) 
    If (ShareMode = 1) and (IsSharing) ; already sharing
        Return
    If (ShareMode = 0) and !(IsSharing) ; already not sharing
        Return

    ;SendInput ^+e ; ctrl+shift+e - toggle share

    ShareEl.Click() ; does not require Window to be active
    
    If (ShareMode=0) or ((ShareMode=2) and IsSharing) { ; unshare->done
        FocusAssist("-") ; deactivate focusassist 
        return 
    }

    ; Include sound
    Name := Teams_GetLangName("ComputerAudio","Include computer sound",Lang)
    If !(Name="") {
        El :=  TeamsEl.WaitElementExistByName(Name,,,,1000) 
        El.Click()
    }

    SendInput {Tab}{Tab}{Tab}{Enter} ; Select first screen - New Share design requires 3 {Tab}
    
    ; Move Meeting Window to secondary screen
    SysGet, MonitorCount, MonitorCount	; or try:    SysGet, var, 80
    If (MonitorCount > 1) {
       ; Wait for Window to be minimized
       Sleep 500
        ; Move to secondary monitor (activates window)
        Monitor_MoveToSecondary(WinId,false)   ; bug: unshare on winactivate
        Sleep 500 ; Wait for move to Maximize
        WinMaximize, ahk_id %WinId%
        Sleep 500 ; Wait for maximize to switch screen
    } ; end if secondary monitor


    ; Hide Sharing Control Bar
    Name := Teams_GetLangName("SharingControlBar","Sharing control bar",Lang)
    If !(Name="") {
        TeamsExe := Teams_GetExeName()
        wTitle =  %Name% ahk_exe %TeamsExe%
        WinWait, %wTitle%,,2
        WinHide, %wTitle%
    }

     ; Activate FocusAssistant
     FocusAssist("+")

} ; eofun
; -------------------------------------------------------------------------------------------------------------------


; -------------------------------------------------------------------------------------------------------------------
Teams_ShareToTeams(sUrl:=""){
If GetKeyState("Ctrl") {
    Teamsy_Help("s2t")
	return
}
If (sUrl = "") && (Browser_WinActive()) {
    sUrl := Browser_GetUrl()
}
InputBox, sUrl , Share To Teams, Enter Link to Share:, , 640, 125,,,,, %sUrl%
If ErrorLevel
    return
sUrl := "https://teams.microsoft.com/share?href=" + sUrl
Run, %sUrl%
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------
Teams_GetCacheDir(isNew := "") {
; CacheDir := Teams_GetCacheDir()
If (isNew="")
    isNew := Teams_IsNew()
If isNew {
    ; %localappdata%\packages\MSTeams_8wekyb3d8bbwe\Localcache\Microsoft\MSTeams
    CacheDir := RegExReplace(A_AppData,"\\[^\\]*$") . "\Local\packages\MSTeams_8wekyb3d8bbwe\Localcache\Microsoft\MSTeams"
} Else {
    CacheDir = %A_AppData%\Microsoft\Teams
}
return CacheDir
} ; eofun


Teams_ClearCache(){ ; @fun_teams_clearcache@
; See https://learn.microsoft.com/en-us/microsoftteams/troubleshoot/teams-administration/clear-teams-cache
    If GetKeyState("Ctrl") {
        Teamsy_Help("clc")
        return
    }

    IsNew := Teams_IsNew()
    If IsNew
        TeamsExe := "ms-teams.exe"
    Else
        TeamsExe := "Teams.exe"

    Process, Exist, %TeamsExe%
    If (ErrorLevel) {
        sCmd = taskkill /f /im "%TeamsExe%"
        Run %sCmd%,,Hide 
    }

    While WinExist("ahk_exe " . TeamsExe)
        Sleep 500

    ; Delete Folders and Files in Cache Directory
    If IsNew { ; https://microsoft365pro.co.uk/2023/07/31/teams-real-simple-with-pictures-clear-cache-in-teams-2-1-client/ @Microsoft365Pro
        TeamsDir := RegExReplace(A_AppData,"\\[^\\]*$") . "\Local\packages\MSTeams_8wekyb3d8bbwe\Localcache\Microsoft\MSTeams" ; AppData env variable points to Roaming        
        FileRecycle, %TeamsDir%\EBWebView
        FileRecycle, %TeamsDir%\Logs 
        FileRecycle, %TeamsDir%\*.dat64
        FileRecycle, %TeamsDir%\*.json
    } Else { ; Teams Classic ; https://learn.microsoft.com/en-us/microsoftteams/troubleshoot/teams-administration/clear-teams-cache
        TeamsDir = %A_AppData%\Microsoft\Teams
        FileRemoveDir, %TeamsDir%\application cache\cache, 1
        FileRemoveDir, %TeamsDir%\blob_storage, 1
        FileRemoveDir, %TeamsDir%\databases, 1
        FileRemoveDir, %TeamsDir%\cache, 1
        FileRemoveDir, %TeamsDir%\gpucache, 1
        FileRemoveDir, %TeamsDir%\Indexeddb, 1
        FileRemoveDir, %TeamsDir%\Local Storage, 1
        FileRemoveDir, %TeamsDir%\tmp, 1 
    }
    Teams_GetMainWindow()
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------
Teams_CleanRestart(){ ; @fun_teams_cleanrestart@
If GetKeyState("Ctrl") {
    sUrl := "https://tdalon.blogspot.com/2021/01/teams-clear-cache.html"
    Run, "%sUrl%"
	return
}
; Warning all appdata will be deleted
MsgBox, 0x114,Teams Clean Restart, Are you sure you want to delete all Teams Client local application data? (incl. custom Backgrounds)
IfMsgBox No
   return

TeamsExe := Teams_GetExeName()
Process, Exist, %TeamsExe%
If (ErrorLevel) {
    sCmd = taskkill /f /im "%TeamsExe%"
    Run %sCmd%,,Hide 
}
While WinExist("ahk_exe " . TeamsExe)
    Sleep 500

TeamsDir = Teams_GetCacheDir(Teams_IsNew())
;FileRemoveDir, %TeamsDir%, 1
FileRecycle, %TeamsDir%\*
Teams_GetMainWindow() ; will restart Teams
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Restart(){

TeamsExe := Teams_GetExeName()
Process, Exist, %TeamsExe%
If (ErrorLevel) {
    sCmd = taskkill /f /im "%TeamsExe%"
    Run %sCmd%,,Hide 
}
While WinExist("ahk_exe " . TeamsExe)
    Sleep 500

Teams_GetMainWindow()
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Quit() {
TeamsExe := Teams_GetExeName()
sCmd = taskkill /f /im "%TeamsExe%"
Run %sCmd%,,Hide
} ; eofun 

; -------------------------------------------------------------------------------------------------------------------
Teams_Mute(State := 2,showInfo:=true,restoreWin:=true){ ; @fun_teams_mute@
; State: 
;    0: mute off, unmute
;    1: mute on
;    2*: (Default): Toggle mute state

WinId := Teams_GetMeetingWindow() 
If !WinId ; empty
    return

If (restoreWin)
    WinGet, curWinId, ID, A

If (showInfo) {
    displayTime := 2000
    Tray_Icon_On := "HBITMAP:*" . Create_Mic_On_ico()
    Tray_Icon_Off := "HBITMAP:*" . Create_Mic_Off_ico()
}
UIA := UIA_Interface()
TeamsEl := UIA.ElementFromHandle(WinId)

Lang := Teams_GetLang()

MuteName := Teams_GetLangName("Mute","Mute",Lang) 
If (MuteName="") 
    return
UnmuteName := Teams_GetLangName("Unmute","Unmute",Lang)
If (UnmuteName ="")
    return

El:=TeamsEl.FindFirstBy("AutomationId=microphone-button")
If !El {
    TrayTip TeamsShortcuts: ERROR, Microphone button UIA Element not found!,,0x2
    Return
}

If RegExMatch(El.FullDescription,"^" . MuteName) {
    If (State = 0) {
        If (showInfo)
            Tooltip("Teams Mic is already on.")
        return
    } Else {
        If (showInfo) {
            Tooltip("Teams Mute Mic...",displayTime)
            TrayIcon_Mic_Off := "HBITMAP:*" . Create_Mic_Off_ico()
            TrayIcon(TrayIcon_Mic_Off,displayTime)
        }
        El.Click() ; activates the window
        If (restoreWin) 
            WinActivate, ahk_id %curWinId%
        return
    }
}

If RegExMatch(El.FullDescription,"^" . UnmuteName) {
    If (State = 1) {
        If (showTooltip)
            Tooltip("Teams Mic is already off.")
        return
    } Else {
        If (showInfo) {
            Tooltip("Teams Unmute Mic...",displayTime)
            TrayIcon_Mic_On := "HBITMAP:*" . Create_Mic_On_ico()
            TrayIcon(TrayIcon_Mic_On,displayTime)
        }            
        El.Click()
        If (restoreWin) 
            WinActivate, ahk_id %curWinId%
        return
    }
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
        Tooltip("Teams Admit from Lobby...") 
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
        SendInput ^+{Space} ; Go to sharing toolbar Ctrl+Shift+Spacebar
    Case "leave":
        Tooltip("Leave meeting...") 
        SendInput ^+h ; Ctrl+Shift+H
}


Sleep 500 ; pause before reactivating previous window
WinActivate, ahk_id %curWinId%


} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_Leave() {
WinId := Teams_GetMeetingWindow(true)
If !WinId ; empty
    return

WinActivate ahk_id %WinId%
SendInput ^+h ; Ctrl+Shift+H

; Reset FocusAssistant
FocusAssist("-")
} ; eofun


; -------------------------------------------------------------------------------------------------------------------
Teams_MeetingLeave(mode:="?") { ; @fun_teams_meetingleave@
; mode : "e" : end meeting "?" ask if you want to end
    WinId := Teams_GetMeetingWindow(true)
    If !WinId ; empty
        return
    
    WinGet, curWinId, ID, A
    WinActivate ahk_id %WinId% ; activate window to click on element
    
    UIA := UIA_Interface()  
    TeamsEl := UIA.ElementFromHandle(WinId)

    El := TeamsEl.FindFirstBy("AutomationId=hangup-button")
    If El {
        El.Click() ; SendInput ^+h ; Ctrl+Shift+H
        GoTo Finish
    }

    Els :=  TeamsEl.FindAllByNameAndType("More options","Button")  ; Name More options ; AutomationId changing menuddd
    for i, El in Els   
        If InStr(El.ClassName,"SplitButton") { ;   
            ; ClassName =fui-Button rlr4yyk fui-MenuButton fui-SplitButton__menuButton toggleButton ___z3kncc0 f1sbtcvk fwiuce9 fdghr9 f15vdbe4 fwbmr0d f44c6la frrbwxo f1um7c6d f6pwzcr fdl5y0r f1p3nwhy f11589ue f1q5o8ev f1pdflbu fonrgv7 f5k7jbu f1s2uweq fr80ssc f1ukrpxl fecsdlb fhhy8jn fqwlww5 f1bbhs8t f1i3by9 fbb8suj f1sw78cg f1vj5d8e f1udo9fm f1hoz3np 
                El.Click() ; Click element without moving the mouse
            EndEl:=TeamsEl.WaitElementExistByName("End meeting",,,,2000)
            If (EndEl) and (mode="e")
                GoTo EndMeeting
            Else If (mode="?") {
                OnMessage(0x44, "OnLeaveMsgBox")
                MsgBox, 0x23, Leave End Meeting,Do you want to End or Leave the meeting?
                OnMessage(0x44, "")
                IfMsgBox Cancel 
                    {
                        El.Click()
                        WinActivate, ahk_id %curWinId%
                        return
                    }

                IfMsgBox Yes ; End
                    GoTo EndMeeting
                Else { ; Leave
                    TeamsEl.FindFirstBy("AutomationId=splitButton-ddd__primaryActionButton").Click()
                    GoTo Finish
                }   
        }
    } 
    
    TrayTip TeamsShortcuts: ERROR, Leave button UIA Element not found!,,0x2
    return
    


    EndMeeting:
    EndEl.Click()
    EndEl:=TeamsEl.WaitElementExistByNameAndType("End","Button",,3,,2000) ; exact match ; TODO Lang specific
    EndEl.Click()

    
    Finish:
    ; Reset FocusAssistant
    FocusAssist("-")
    ; Restore previous window
    WinActivate, ahk_id %curWinId%

    ; Dismiss Call Quality Window
    El:=TeamsEl.WaitElementExist("AutomationId=cqf-dismiss-button",,,,1000) 
    If El
        El.Click()

    
} ; eofun


OnLeaveMsgBox() {
    DetectHiddenWindows, On
    Process, Exist
    If (WinExist("ahk_class #32770 ahk_pid " . ErrorLevel)) {
        ControlSetText Button1, End
		ControlSetText Button2, Leave
    }
}


; -------------------------------------------------------------------------------------------------------------------
Teams_PushToTalk(KeyName:="MButton"){

Cnt := 0
MinCnt := 2

while (GetKeyState(KeyName , "P"))
{
    sleep, 100
    Cnt += 1
    If (Cnt=MinCnt) {
        Teams_Mute(0,false)
        ToolTip("Teams PushToTalk on...",2000) 
        Tray_Icon_On := "HBITMAP:*" . Create_Mic_On_ico()
        ;Tray_Icon_Off := "HBITMAP:*" . Create_Mic_Off_ico()
        Menu, Tray, Icon, %Tray_Icon_On%
        
    }
}

If (Cnt>MinCnt) {
    Teams_Mute(1,false)
    Tooltip("Teams PushToTalk off...",2000)
    IcoFile  := PathX(A_ScriptFullPath, "Ext:.ico").Full
    If (FileExist(IcoFile)) 
	    Menu,Tray,Icon, %IcoFile%
} Else
    Teams_Mute() 
  
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
newHK := Hotkey_GUI(,prevHK,,,"Teams " . HKid . " - Set Global Hotkey")

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

MuteName := Teams_GetLangName("Mute","Mute",Lang) 
If (MuteName="") 
    return
return !RegExMatch(TeamsEl.FullDescription,"^" . MuteName)
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

WinGet, curWinId, ID, A

UIA := UIA_Interface()
TeamsEl := UIA.ElementFromHandle(WinId)

;El :=  TeamsEl.FindFirstByNameAndType("Turn camera on", "button",,1)
El :=  TeamsEl.FindFirstByName("Turn camera on",,1) ; menu item on meeting window; button on Call in progress
If El {
    If (State = 0) {
        Tooltip("Teams Camera is already off.")
        return
    } Else {
        Tooltip("Teams Camera On...")
        El.Click()
        WinActivate, ahk_id %curWinId%
        return
    }
}
;El :=  TeamsEl.FindFirstByNameAndType("Turn camera off", "button",,1)
El :=  TeamsEl.FindFirstByName("Turn camera off",,1)
If El {
    If (State = 1) {
        Tooltip("Teams Camera is already on.")
        return
    } Else {
        Tooltip("Teams Camera off...")
        El.Click()
        WinActivate, ahk_id %curWinId%
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
    
TeamsExe := Teams_GetExeName()
sCmd = "%SVVExe%" %sCmd% "%TeamsExe%"
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
Teams_RaiseHand2() {
; Toggle Raise Hand on/off ; Default Hotkey Ctrl+Shift+K

WinId := Teams_GetMeetingWindow()
If !WinId ; empty
    return
WinGet, curWinId, ID, A
WinActivate, ahk_id %WinId%
Tooltip("Teams Toggle Raise Hand...") 
SendInput ^+k ; toggle video Ctl+Shift+k
WinActivate, ahk_id %curWinId%
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_RaiseHand(State := 2,showInfo:=true,restoreWin:=true){
; State: 
;    0: lower hand
;    1: raise hand
;    2*: (Default): Toggle raise hand

WinId := Teams_GetMeetingWindow() 
If !WinId ; empty
    return

If (restoreWin)
    WinGet, curWinId, ID, A

If (showInfo) {
    displayTime := 2000
}
    
UIA := UIA_Interface()
TeamsEl := UIA.ElementFromHandle(WinId)

Lang := Teams_GetLang()
Prop := "RaiseHand"
If InStr(Lang,"en-") or (Lang="") {
    RaiseHandName := "Raise"
    LowerHandName := "Lower"
} Else
    RaiseHandName := Teams_GetLangName(Prop,Lang)
If (RaiseHandName="") {
    Text := "Language " . Lang . " not implemented!"
    sUrl := Teamsy_Help("lang",false)
    PowerTools_ErrDlg(Text,sUrl:="")
}

If (LowerHandName ="") {
    LowerHandName := Teams_GetLangName("LowerHand",Lang)
}

El:=TeamsEl.FindFirstBy("AutomationId=raisehands-button")
;MsgBox % El.DumpAll()

If RegExMatch(El.FullDescription,"^" . LowerHandName) {
    If (State = 1) {
        If (showInfo)
            Tooltip("Hand is already raised.")
        return
    } Else {
        If (showInfo) {
            Tooltip("Lower hand...",displayTime)
            TrayIcon:= "HBITMAP:*" . Create_LowerHand_ico()
            TrayIcon(TrayIcon,displayTime)
        }
        El.Click()
        If (restoreWin) 
            WinActivate, ahk_id %curWinId%
        return
    }
}

If RegExMatch(El.FullDescription,"^" . RaiseHandName) {
    If (State = 0) {
        If (showTooltip)
            Tooltip("Hand is already lowered.")
        return
    } Else {
        If (showInfo) {
            Tooltip("Raise hand...",displayTime)
            TrayIcon := "HBITMAP:*" . Create_RaiseHand_ico()
            TrayIcon(TrayIcon,displayTime)
        }
            
        El.Click()
        If (restoreWin) 
            WinActivate, ahk_id %curWinId%
        return
    }
}

    
} ; eofun



; -------------------------------------------------------------------------------------------------------------------

Teams_IsWinActive(){
; Check if Active window is Teams client or a Browser/App window with a Teams url opened
TeamsExe := Teams_GetExeName()
If WinActive("ahk_exe " . TeamsExe)
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


WinId := Teams_GetMeetingWindow() 
If !WinId ; empty
    return

UIA := UIA_Interface()  
TeamsEl := UIA.ElementFromHandle(WinId)

; Language specific implementation
Lang := Teams_GetLang()
Name := Teams_GetLangName(Reaction,Reaction,Lang)
If (Name="") 
    return

WinGet, curWinId, ID, A

; Shortcut if Reactions toolbar already available-> directly click and exit
ReactionEl := TeamsEl.FindFirstByName(Name)
If ReactionEl {
    WinActivate, ahk_id %WinId% ; window must be active for click 
    Goto, React
}

ReactionsEl :=  TeamsEl.FindFirstBy("AutomationId=reaction-menu-button")  
If ReactionsEl
    ReactionsEl.Click() ; Click element without moving the mouse - will activate window
Else
    WinActivate, ahk_id %WinId% ; window must be active for click to open menus

ReactionEl:=TeamsEl.WaitElementExistByName(Name,,,,2000) ; timeout=2s

If !ReactionEl {
    TrayTip TeamsShortcuts: ERROR, Meeting Reaction button for '%Reaction%'' not found!,,0x2
    ;MsgBox % ReactionsEl.DumpAll() ; DBG
    return
} 

React:
ReactionEl.Click()

;WinGetTitle, T, ahk_id %curWinId%
;MsgBox % T

; Restore previous window 
WinActivate, ahk_id %curWinId%
Tooltip("Teams Meeting Reaction: " . Reaction,1000)
IcoFile := "HBITMAP:*" . Create_mr_%Reaction%_ico()
TrayIcon(IcoFile,2000)
     
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

If Teams_IsNew() {
    UIA := UIA_Interface()  
    TeamsEl := UIA.ElementFromHandle(WinId)

    El :=  TeamsEl.FindFirstBy("AutomationId=view-mode-button")
    El.Click() 
    El :=  TeamsEl.WaitElementExist("AutomationId=ViewModeMoreOptionsMenuControl-id",,,,1000)  
    El.ControlClick() ; Click does not work see https://github.com/Descolada/UIAutomation/issues/45
    El :=  TeamsEl.WaitElementExist("AutomationId=fullscreen-button",,,,1000)  
    El.Click()
    return

}

; For Legacy/ Classic Teams
; Needs to activate the Meeting Window because F11 Hotkey is not working even with ControlSend,,{F11}, ahk_id %WinId%
If (restore)
    WinGet, curWinId, ID, A

WinActivate, ahk_id %WinId% 
Send {F11}


 ; Click on View -> Video Effects
 El :=  TeamsEl.FindFirstBy("AutomationId=callingButtons-showMoreBtn")  
 El.Click() 
 El :=  TeamsEl.WaitElementExist("AutomationId=video-effects-and-avatar-button",,,,1000)  
 El.Click() 



; restore previous window
If (restore)
    WinActivate, ahk_id %curWinId%
}

; -------------------------------------------------------------------------------------------------------------------

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
        TeamsExe := Teams_GetExeName()
        WinWaitActive, ahk_exe %TeamsExe%,,5
        If ErrorLevel
            Return
    }
    UIA := UIA_Interface()
    If !WinId
        WinId := WinActive("A")
    TeamsEl := UIA.ElementFromHandle(WinId) 
    JoinBtn :=  TeamsEl.FindFirstBy("AutomationId=prejoin-join-button")  
} ; eofun

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
} ; eofun
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

Teams_GetLang() { ; @fun_teams_getlang@
; return desktop client language
; sLang := Teams_GetLang()

; For Classic Teams:
;   Read value in %AppData%\Microsoft\Teams\desktop-config.json -> currentWebLanguage
; For New Teams: 
;           %LOCALAPPDATA%\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\app_settings.json -> language
; e.g. "en-us"
static Lang
If !(Lang = "")
    return Lang

If Teams_IsNew() {
    EnvGet, LOCALAPPDATA, LOCALAPPDATA
    JsonFile := LOCALAPPDATA . "\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\app_settings.json"
    FileRead, Json, %JsonFile%
    If ErrorLevel {
        TrayTip, Error, Reading file %JsonFile%!,,3
        return
    }
    oJson := Jxon_Load(Json)
    Lang := oJson["language"]
    If (Lang="")
        {

        }
    return Lang
} Else {
    JsonFile := A_AppData . "\Microsoft\Teams\desktop-config.json"
    FileRead, Json, %JsonFile%
    If ErrorLevel {
        TrayTip, Error, Reading file %JsonFile%!,,3
        return
    }
    oJson := Jxon_Load(Json)
    Lang := oJson["currentWebLanguage"]
    return Lang
}

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Teams_GetLangName(Prop,Def,Lang:="") { ; @fun_teams_getlangname@
; Name := Teams_GetLangName(Prop,Def,Lang:="")
; Def: Default value for English/unspecified language
    If (Lang="")
        Lang:=Teams_GetLang()

    If InStr(Lang,"en-") or (Lang="")
        return Def
    
    sHelpUrl := Teamsy_Help("lang",false)
    
    
    If !FileExist("PowerTools.ini") {
        PowerTools_ErrDlg("No PowerTools.ini file found! Language specific setting not implemented!")
        return
    }

    IniRead, TeamsLangNames, PowerTools.ini,Teams,TeamsLangNames
    If (TeamsLangNames="ERROR") { ; Section [Jira] Key JiraAuth not found
        PowerTools_ErrDlg("TeamsLangNames key not found in PowerTools.ini file [Teams] section!",sHelpUrl)
        return
    }

    JsonObj := Jxon_Load(TeamsLangNames)
    For i, langName in JsonObj 
    {
        lang_i := langName["lang"]
        If InStr(Lang,lang_i) {
            Name := langName[Prop]
            If (Name="") { 
                PowerTools_ErrDlg("Property '" . Prop . "'' is not defined in PowerTools.ini file [Teams] section, TeamsLangNames key for lang '" . Lang . "'!",sHelpUrl)
                return
            }
            return Name 
        }
    }
    ; Lang could not be found take first entry as default value
    If (Name="") {
        Text := "Language " . Lang . " for '" . Prop . "' not implemented!"
        PowerTools_ErrDlg(Text,sHelpUrl)
    }

}
; -------------------------------------------------------------------------------------------------------------------




Teams_SetStatusMessage() {
    If GetKeyState("Ctrl")  { ; exclude ctrl if use in the hotkey
        Teamsy_Help("st")
        return
    }
    WinId := Teams_GetMainWindow()
    If !WinId ; empty
        return
    UIA := UIA_Interface()
    TeamsEl := UIA.ElementFromHandle(WinId)
    El := TeamsEl.FindFirstBy("AutomationId=idna-me-control-set-status-message-trigger")
    If !El { ; menu not opened 
        ; Click on avatar
        MeCtrl :=  TeamsEl.FindFirstBy("AutomationId=idna-me-control-avatar-trigger")
        MeCtrl.Click()
        El:= TeamsEl.WaitElementExist("AutomationId=idna-me-control-set-status-message-trigger")
    }
    
    El.Click()

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Teams_SwitchTenant(sTenant) {
    If GetKeyState("Ctrl")  { ; exclude ctrl if use in the hotkey
        Teamsy_Help("sw")
        return
    }
    WinId := Teams_GetMainWindow()
    If !WinId ; empty
        return
    UIA := UIA_Interface()
    TeamsEl := UIA.ElementFromHandle(WinId)

    If !TeamsEl.FindFirstBy("AutomationId=idna-me-control-set-status-message-trigger") { ; menu not opened 
        ; Click on avatar
        MeCtrl :=  TeamsEl.FindFirstBy("AutomationId=idna-me-control-avatar-trigger")
        MeCtrl.Click()
        El:= TeamsEl.WaitElementExistByNameAndType("Switch to " . sTenant,"MenuItem",,1,False,1000)
    } Else { ; menu already opened
        El:= TeamsEl.FindFirstByNameAndType("Switch to " . sTenant,"MenuItem",,1,False)
    }
    
    If (El="") {
        TrayTipAutoHide("Switch Tenant","Tenant name starting with '" . sTenant . "' not found!")
    } Else
        El.Click()
} ; eofun
; -------------------------------------------------------------------------------------------------------------------



; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_Mic_On_ico(NewHandle := False) {
	Static hBitmap := 0
	If (NewHandle)
	   hBitmap := 0
	If (hBitmap)
	   Return hBitmap
	VarSetCapacity(B64, 884 << !!A_IsUnicode)
	B64 := "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAMAAAAoLQ9TAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAAnFBMVEUAAAAAxgAA0gwA0w0A0g0A0w0A0w0A3wAA0w0A1Q4A1g4A1AsAzgwA0w0A0wkAzAkA2Q0AqgAA0w0A0w0A0w0A1Q0A0w0A/wAA0w4A0w0A0w8A0g0A1A4A0w0A0g4A1w0A0w0A0w0A1w0A1A0A2A0A0w0A0w0A1Q4A0QkA0g0A0w0A0w0A2AoA0A0A0g8A1A0A0w0A0w0A0w0AAACNi8PxAAAAMnRSTlMACY/mjp2cCPokJS8V/B0eFAPvOtM87gKWxTTlNcaUE96zJk0ntN0SHLfMtRomIp/49wBmOSwAAAABYktHRACIBR1IAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAAB3RJTUUH5QEdDQIQMhdn1gAAAIBJREFUGNNjYAABRiZmZhZGBgRgNQICNgSfnQMkwIFQwmkEBpxwAS6IABfxAtwwAW4In4eXDyLAzysAFhAUEhYB8UVExcQhSiQkpaSNjKRlZOWgZsgrKCopK6uoqqnDTNXQ5NDS4tDWQfKMrqIiFxKXQUdPX18PSQGPAcgWQ7ClALhDEXOeykbUAAAAJXRFWHRkYXRlOmNyZWF0ZQAyMDIxLTAxLTI5VDEzOjAyOjE2KzAwOjAwTlt0UgAAACV0RVh0ZGF0ZTptb2RpZnkAMjAyMS0wMS0yOVQxMzowMjoxNiswMDowMD8GzO4AAAAZdEVYdFNvZnR3YXJlAHd3dy5pbmtzY2FwZS5vcmeb7jwaAAAAAElFTkSuQmCC"
	If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
	   Return False
	VarSetCapacity(Dec, DecLen, 0)
	If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
	   Return False
	; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
	; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
	hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
	pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
	DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
	DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
	DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
	hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
	VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
	DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
	DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
	DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
	DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
	DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
	DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
	DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
	Return hBitmap
} ; eofun
	
; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_Mic_Off_ico(NewHandle := False) {
	Static hBitmap := 0
	If (NewHandle)
	   hBitmap := 0
	If (hBitmap)
	   Return hBitmap
	VarSetCapacity(B64, 920 << !!A_IsUnicode)
	B64 := "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAMAAAAoLQ9TAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAAolBMVEUAAAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAD/AAAAAACyecVlAAAANHRSTlMAA4Pm5YJ0c07jSMf5+jna9VBiGlbrT+GkeMxc4LOmq0r4Rv78o1k4dpue/XI9tkAUM7ET+6bC8AAAAAFiS0dEAIgFHUgAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAAHdElNRQflAR0NAQrkWM1vAAAAkUlEQVQY01XP2RaCIBCA4dGMLCXArcXMFttXbN7/2YLgHGiu+L8zFwOAniAcRMMA3BBUM4J4PLGQaEjilE4tMA08pQJ8QNcGvP6B11nOdBd5ZqGsmGrCqtLCbE6pWOCyXlloUDeG69b0ZkvFbt8djp0566TuOV/q4nq724VIwKPh/PlyZ7yl7HspP/+fRST6/QXOExB226vKAgAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAyMS0wMS0yOVQxMzowMToxMCswMDowMMa8+msAAAAldEVYdGRhdGU6bW9kaWZ5ADIwMjEtMDEtMjlUMTM6MDE6MTArMDA6MDC34ULXAAAAGXRFWHRTb2Z0d2FyZQB3d3cuaW5rc2NhcGUub3Jnm+48GgAAAABJRU5ErkJggg=="
	If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
	   Return False
	VarSetCapacity(Dec, DecLen, 0)
	If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
	   Return False
	; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
	; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
	hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
	pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
	DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
	DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
	DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
	hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
	VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
	DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
	DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
	DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
	DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
	DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
	DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
	DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
	Return hBitmap
} ;eofun




; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_mr_Laugh_ico(NewHandle := False) {
    Static hBitmap := 0
    If (NewHandle)
       hBitmap := 0
    If (hBitmap)
       Return hBitmap
    VarSetCapacity(B64, 2800 << !!A_IsUnicode)
    B64 := "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"
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    VarSetCapacity(Dec, DecLen, 0)
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    ; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
    ; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
    hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
    pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
    DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
    DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
    DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
    hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
    VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
    DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
    DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
    DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
    DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
    DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
    DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
    DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
    Return hBitmap
    }


    ; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_mr_Like_ico(NewHandle := False) {
    Static hBitmap := 0
    If (NewHandle)
       hBitmap := 0
    If (hBitmap)
       Return hBitmap
    VarSetCapacity(B64, 2044 << !!A_IsUnicode)
    B64 := "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"
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    VarSetCapacity(Dec, DecLen, 0)
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    ; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
    ; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
    hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
    pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
    DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
    DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
    DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
    hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
    VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
    DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
    DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
    DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
    DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
    DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
    DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
    DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
    Return hBitmap
    }

    ; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_mr_Love_ico(NewHandle := False) {
    Static hBitmap := 0
    If (NewHandle)
       hBitmap := 0
    If (hBitmap)
       Return hBitmap
    VarSetCapacity(B64, 2148 << !!A_IsUnicode)
    B64 := "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"
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    VarSetCapacity(Dec, DecLen, 0)
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    ; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
    ; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
    hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
    pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
    DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
    DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
    DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
    hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
    VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
    DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
    DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
    DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
    DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
    DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
    DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
    DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
    Return hBitmap
    }


    ; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_mr_Surprised_ico(NewHandle := False) {
    Static hBitmap := 0
    If (NewHandle)
       hBitmap := 0
    If (hBitmap)
       Return hBitmap
    VarSetCapacity(B64, 2896 << !!A_IsUnicode)
    B64 := "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"
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    VarSetCapacity(Dec, DecLen, 0)
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    ; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
    ; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
    hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
    pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
    DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
    DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
    DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
    hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
    VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
    DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
    DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
    DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
    DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
    DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
    DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
    DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
    Return hBitmap
    }

    ; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_mr_Applause_ico(NewHandle := False) {
    Static hBitmap := 0
    If (NewHandle)
       hBitmap := 0
    If (hBitmap)
       Return hBitmap
    VarSetCapacity(B64, 2440 << !!A_IsUnicode)
    B64 := "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAACXBIWXMAABYlAAAWJQFJUiTwAAAG1klEQVRYhcWXe2xXZxnHP+97Lr/zu9NCC22hpYWCIJdk4MAxVBI3gS2Ni9s08YLZxYRNjRlx84/hnCZGdG4SFxJZjLrpNhOmri77YzHZAixZDNM5oeNWCgJtf5e2v/vv3F//6K/QRWhxEfckb/LknOc83895znve93mFUooP0+RsAWLzA01CPD5r3Ac1/arCN+3sYGnfU6zd2UnrP+4ARv4vAKJnUxdzOj3ceB21sINUbxtJ7Osh/h8A4pO7e/n0T15Eqb+qZzbtFOv23Uf9/P3IEzasvv4ASGs+dGrE9T6xdc/T6u1HjgEPwfbrIv4+ANG9xVJDrx8WYu0mbtuznaqqTw985GaRXNGL9tVfqQJA33Ix72M3sGLNWtrzNU7e+3319we3iqXf2MG2YsDwxi/zslLKnw1ATP2G4nO/+wELPrKM0Vd2qZceuzA96M9fELffvJldcj7uobM8/eavee/e29mvddGxeBX1UpzRV//C77sDvrhxG3NpI//WIfo33ql+PhuABBBCSOS8DbSsu5H5dz0ptjx+qTL714vYpoU82hSlNW2S+EQ3X//Sp9jfGWdxRBGERZx0GnXXdu5Z0clqNQSME350JVv37hbLr6kCQgjB+h/fhNQl7d3n+dMd51SjNEKIePZrHGpZQIwEddqxvWHahjPEw/nU2+aRsTopswS3fpoe7SLKbGeE9ZT/eYS/rb5FPXZNn+BK1v8Vsb67hx0tglub66QIKKoUbrEVEWnHkDoRMUwQT5ARPVRVCtN+l8UW5MRSCnYC97f9fPe+Xert7m4x5/7bWFfJc/6HL6qTUxpXXYi+s16kv7mZn7Z304xP2cmilQdJ1nzKWiuixaIouii4Ogv9c6SNHJ5IUdc7mfBPkTZGqVkbcLZs4LO3dIjc8w+xb+0Guuwo/+r/jdjTt0MdvDQHrmSeQyrp08opCM+iaVB1dWStQEQV8IIRPPJgLiKrkliMESWLYcwjTxOKEgnOEFnUzsqHH+CJNR3cEJP4zSsJu+ZztxBCzgjw5FEyF0qcCGvE6+Po5TyGilANFJZbo9kpkiCDiY8tuyh6HmmymFQJ9C7GfEiSwTJDolu20RlIdHxCSkhTIwJoVwV4/k7Rc/phnl26miVVA6tSJl7NIv1mvI4+Si034vomTcE4FhkMvZkx1YoICyQZxRARSrKLCoIo45h6M8VYOwUEAg1l29SVUt5VAVakeXDJIlYbHqERY8IRWI6H0NPIhEE12cmpaC9Fz6WZHAYllNZFPkgTpYZFAU3OI08bLhIDD6W3cZEkNhoE6vLeckUAv4rFKL4/SuAXUTKKZ/tEwxy2P4JLDmEuIMdc9HCCJBkMTVI2esmTQsfGoIYgQh2LAB8NDZs0FUxUGODOCDB0gQOVYQQ2ypkAQjwERvkMspYBNYJJDc/oImdrNKk8UcbRETgkcRqeJCAEfAIUPiDx0VFSw58R4O7D6vWBLId1IKzjUyWMRPGDgFQpg7Sz6IxiaHEKWid1X9BEGYsaEoGPRgAofFTDAw9JAAiUUogZAQBeO83eCwWClIYvHVwT/IiFYReJlzOIIIPBGDLSxqjejgMYeAgUITpBw6MxJkF8ROOOMSvAo8fViWNjHIgYKBNczcY1Q0LdJVHJotUzSIYxKKCERZkYLooQH0WIarz3pLgCAiBA4iF1nfisAAC/OM4vzxQozI0R6D6O5uIYChmWSJay4GXQyaJTm0rdkJuUng4x6flADW1OnDm9vSIyK8AfLqixIxmeCRTEDTzTx9YCHNMnameJVHIQ5NGooHAICRtS4TSEyyiT5feQUZ3ox3smt4FZu93Pv8Ufj2YZSMQITIlnhDiGIDCqpCqj4I4DRaA+bdKF08YURAiAQCJKRWrPvYZzTQBKKe+NIZ7KVyFhERgCTwuxDYUejhGrZPHVBFAG7EaZQ7g0+6f8SQiBQIYlzk51S9fU73/rpHrn3Yu8ZEicqMQzFI4RUo+4xJyLYI/ik0dRRuGgcKdVYxJE4COQSIaJnDrFwanc13zg2HuSfUfPMRA1qUXBNQJqVoAbKRArncb2hwnJAaVGJRzAg8Z3n1wdoojBdxjo+94HAHg5p8p7jrB7YJD3LEU5pajFfMoJH2VliFUGqKmz+IygGG+AVBFUkThomITn32DihX5+NL1ZnbEjupItEc3pF5ZN3LO8ic+kLQIEGgEGPmWilFiMwRIizMUggUYURQ0xdI7BA6/ys2+/qQan5/uvAabsiaRYtaWFjYFkWVqSqguCwOa45jJY72BNool2UlA3uViscOTWVzg4tQX/TwDel0SsMuFYoJQKLl/rjUBEKXXUnfHZD/t4/m9ptC7GzVbVYQAAAABJRU5ErkJggg=="
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    VarSetCapacity(Dec, DecLen, 0)
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    ; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
    ; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
    hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
    pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
    DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
    DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
    DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
    hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
    VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
    DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
    DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
    DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
    DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
    DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
    DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
    DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
    Return hBitmap
    }


; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_RaiseHand_ico(NewHandle := False) {
    Static hBitmap := 0
    If (NewHandle)
       hBitmap := 0
    If (hBitmap)
       Return hBitmap
    VarSetCapacity(B64, 2104 << !!A_IsUnicode)
    B64 := "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"
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    VarSetCapacity(Dec, DecLen, 0)
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    ; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
    ; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
    hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
    pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
    DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
    DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
    DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
    hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
    VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
    DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
    DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
    DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
    DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
    DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
    DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
    DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
    Return hBitmap
    }


; ##################################################################################
; # This #Include file was generated by Image2Include.ahk, you must not change it! #
; ##################################################################################
Create_LowerHand_ico(NewHandle := False) {
    Static hBitmap := 0
    If (NewHandle)
       hBitmap := 0
    If (hBitmap)
       Return hBitmap
    VarSetCapacity(B64, 752 << !!A_IsUnicode)
    B64 := "iVBORw0KGgoAAAANSUhEUgAAABkAAAAaCAYAAABCfffNAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAASdEVYdFNvZnR3YXJlAEdyZWVuc2hvdF5VCAUAAAGpSURBVEhL7ZU9rwFBFIbvb0OIApWvpUKERK1TicYfECJUSoVSFHQqiQjiKxp6GoXGe3POHWN9zazkXrfxJJvZc3azj3dndnzhDXwkL/E/ktPphHg8jvV6LToXSqUSPB4PyuWy6FjjYZLlcoloNCpF1WoVdrudD+rZbDY4nU7UajW+ruPp65rNZjAMgx/a7XZ5dDgcfI0k9ENcLhfXOpRzQpJgMCgTURKCJOZRh1JCDxkMBvD5fCyq1+uybx51aCXnMRAIyER/JqFEfr+fRa1WS/atYElCy5YYDod3iQqFArxeL1KpFMbjMfdvsSQxMxqNEAqFZKJer8f9TqeDWCzG57e8LCEmkwnC4bBMtNvtkMlkkMvluL5FKXG73Tgej6K6ZjqdIhKJsIi+o0ajwbvFI5QSir/ZbER1T7vd5vk4J3qGUkJ7VLPZFNU1+/2e56ZSqcid4RlKCW0diURCVNfk83kUi0U+n8/n8tU9Qikh0uk0+v2+qH6gpUzfzOFwEJ37TdWMVrLdbpFMJnmlnQ/6K1gsFuKOC6vVCtlsVlQXtJLf4CN5iTdIgG9QfYFVYky38QAAAABJRU5ErkJggg=="
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", 0, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    VarSetCapacity(Dec, DecLen, 0)
    If !DllCall("Crypt32.dll\CryptStringToBinary", "Ptr", &B64, "UInt", 0, "UInt", 0x01, "Ptr", &Dec, "UIntP", DecLen, "Ptr", 0, "Ptr", 0)
       Return False
    ; Bitmap creation adopted from "How to convert Image data (JPEG/PNG/GIF) to hBITMAP?" by SKAN
    ; -> http://www.autohotkey.com/board/topic/21213-how-to-convert-image-data-jpegpnggif-to-hbitmap/?p=139257
    hData := DllCall("Kernel32.dll\GlobalAlloc", "UInt", 2, "UPtr", DecLen, "UPtr")
    pData := DllCall("Kernel32.dll\GlobalLock", "Ptr", hData, "UPtr")
    DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", pData, "Ptr", &Dec, "UPtr", DecLen)
    DllCall("Kernel32.dll\GlobalUnlock", "Ptr", hData)
    DllCall("Ole32.dll\CreateStreamOnHGlobal", "Ptr", hData, "Int", True, "PtrP", pStream)
    hGdip := DllCall("Kernel32.dll\LoadLibrary", "Str", "Gdiplus.dll", "UPtr")
    VarSetCapacity(SI, 16, 0), NumPut(1, SI, 0, "UChar")
    DllCall("Gdiplus.dll\GdiplusStartup", "PtrP", pToken, "Ptr", &SI, "Ptr", 0)
    DllCall("Gdiplus.dll\GdipCreateBitmapFromStream",  "Ptr", pStream, "PtrP", pBitmap)
    DllCall("Gdiplus.dll\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "PtrP", hBitmap, "UInt", 0)
    DllCall("Gdiplus.dll\GdipDisposeImage", "Ptr", pBitmap)
    DllCall("Gdiplus.dll\GdiplusShutdown", "Ptr", pToken)
    DllCall("Kernel32.dll\FreeLibrary", "Ptr", hGdip)
    DllCall(NumGet(NumGet(pStream + 0, 0, "UPtr") + (A_PtrSize * 2), 0, "UPtr"), "Ptr", pStream)
    Return hBitmap
    }