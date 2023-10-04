#Include <Jira>
#Include <Confluence>

; Launcher for Atlassian related Tools: Jira, Confluence, Bitbucket, BigPicture, R4J, Xray 
; Example of commands
; j : jira
; jc: jira cloud
; c : confluence
; cc : confluence cloud
; bp : BigPicture
; r : r4j
; d|g doc or guidelines (PMT space)

; s : search
; r : recent
; -bm: by me
; -l : last modified
; j c projectKey: create issue


Atlasy(sInput:="-g"){
    
    FoundPos := InStr(sInput," ")  

    If FoundPos {
        sKeyword := SubStr(sInput,1,FoundPos-1)
        sInput := SubStr(sInput,FoundPos+1)
    } Else {
        sKeyword := sInput
        sInput := ""
    }

    Switch sKeyword
    {
    Case "-g": ; gui/ launcher    
        sCmd := AtlasyInputBox()
        if ErrorLevel
            return
        sCmd := Trim(sCmd) 
        Atlasy(sCmd)
        return
    Case "r": ; r4j
        If (sInput = "") {
            sIssueKey := R4J_GetIssueKey()
            R4J_OpenIssue(sIssueKey)
            return
        } 
        StringUpper, sInput, sInput ; R4J is case sensitive for keys->convert to uppercase
        If (RegExMatch(sInput,"\d")) ; Issue Key
            R4J_OpenIssue(sInput)
        Else
            R4J_OpenProject(sInput)
    Case "h","-h","help":
        Atlasy_Help(sInput)
        return
    Case "c": ; Confluence
        FoundPos := InStr(sInput," ")  
        If FoundPos {
            sSpace := SubStr(sInput,1,FoundPos-1)
            sQuery:= SubStr(sInput,FoundPos+1)
        } Else
            sSpace := sInput
        Confluence_SearchSpace(sSpace,sQuery)
    Case "bp":  ; BigPicture
        ; %JiraRootUrl%/plugins/servlet/ac/eu.softwareplant.bigpicture/bigpicture
        
        return
    } ; end case keyword


    
} ; End function     


AtlasyInputBox(){
    static
    ButtonOK:=ButtonCancel:= false
	Gui GuiAtlasy:New,, Atlasy
    ;Gui, add, Text, ;w600, % Text
    Gui, add, Edit, w190 vAtlasyEdit
    Gui, add, Button, w60 gAtlasyOK Default, &OK
    Gui, add, Button, w60 x+10 gAtlasyHelp, &Help

    Gui +AlwaysOnTop -MinimizeBox ; no minimize button, always on top-> modal window
	Gui, Show
    while !(ButtonOK||ButtonCancel)
        continue
    if ButtonCancel {
        ErrorLevel := 1
        return
    }
    Gui, Submit
    ErrorLevel := 0
    return AtlasyEdit
    ;----------------------
    AtlasyOK:
    ButtonOK:= true
    return
    ;---------------------- 
    AtlasyHelp:
    Atlasy_Help()

    GuiAtlasyGuiEscape:
	GuiAtlasyGuiClose:
    
    ButtonCancel:= true
    
    Gui, Cancel
    return
}



Atlasy_Fav(){
    ; b|f : bookmark search use \ for folder match case insensitive

    ;Browser="C:\Program Files\Google\Chrome\Application\chrome.exe" --profile-directory="Profile 2"
    ;C:\Users\thierry.dalon\AppData\Local\Google\Chrome\User Data\Profile 1

    If !FileExist("PowerTools.ini") {
        MsgBox 0x1010, Error, PowerTools.ini file not found!
        return
    }    
    IniRead, Browser, PowerTools.ini,General,Browser
    If (Browser="ERROR") { ; Section [General] Key Browser not found
        MsgBox 0x1010, Error, Browser Entry in PowerTools.ini not found!
        return
    }
    Browser = "%Browser%" ; bug in IniRead if between ""

    sPat = --profile-directory="(.*)"
    If RegExMatch(Browser,sPat,sMatch)
        sProfile := sMatch1
    Else
        sProfile := "Default"
    BookmarksFile := RegExReplace(A_AppData,"[^\\]*$","") . "Local\Google\Chrome\User Data\" . sProfile . "\Bookmarks"
    MsgBox % BookmarksFile ; DBG
    If !FileExist(BookmarksFile) {
        MsgBox 0x1010, Error, Bookmarks file %BookmarksFile% not found!
        return
    }

    FileRead, Bookmarks, %BookmarksFile%
    MsgBox %Bookmarks%
    BookmarksJson := Jxon_Load(Bookmarks)


} ; eofun

; -----------------------------------------
Atlasy_OpenUrl(sUrl) {
    If Browser_WinActive() {
        sCurUrl := Browser_GetUrl()
        If Confluence_IsUrl(sCurUrl) or Jira_IsUrl(sCurUrl) {
            Send ^t ; Open new Tab
            Sleep 100
            Clip_Paste(sUrl)
            SendInput {enter}
            return
        }
    }

    ; Get Browser from PowerTools.ini Settings [Atlasy] section with optional keys Browser and BrowserCloud
    If FileExist("PowerTools.ini") {
        If InStr(sUrl,".atlassian.net") { ; 
            IniRead, BrowserCloud, PowerTools.ini,Atlasy,BrowserCloud
            If !(BrowserCloud="ERROR") { ; Section [General] Key Browser not found
                BrowserCmd = "%BrowserCloud%" ; bug in IniRead if between ""
            }
        }
        If (BrowserCmd = "") {
            IniRead, Browser, PowerTools.ini,Atlasy,Browser
            If !(Browser="ERROR") { ; Section [General] Key Browser not found
                BrowserCmd = "%Browser%" ; bug in IniRead if between ""
            }
        }
    }
    If (BrowserCmd = "")
        Run %sUrl%
    Else {
        sCmd = %BrowserCmd% "%sUrl%" ; Reading String in Ini files removes trailing quotes
	    Run, %sCmd%
    }



}
; -----------------------------------------

Atlasy_Help(sKeyword:=""){
    Switch sKeyword 
    {
    Case "":
        sUrl := "https://tdalon.github.io/ahk/Atlasy"
    Case "2c","oc":
        sUrl := ""
    Case "f","fav","f+","of": ; favorites
        sUrl := ""
    Default:
        sUrl := "https://tdalon.github.io/ahk/Atlasy"
    } ; end switch
    Run, "%sUrl%"
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Atlasy_Launcher(){
    Atlasy("-g")
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Atlasy_HotkeySet(HKid){
    If GetKeyState("Ctrl")  { ; exclude ctrl if use in the hotkey
        sUrl := "https://tdalon.github.io/ahk/Atlasy-Global-Hotkeys"
        Run, "%sUrl%"
        return
    }
    
    ; For Menu callback, remove ending Hotkey and blanks and (Hotkey)
    
    HKid := RegExReplace(HKid,"\t(.*)","")
    ;HKid := Trim(HKid) ; remove tab for align right hotkey in menu
    HKid := RegExReplace(HKid," Hotkey$","")
    HKid := StrReplace(HKid," ","")
    
    RegRead, prevHK, HKEY_CURRENT_USER\Software\PowerTools, AtlasyHotkey%HKid%
    newHK := Hotkey_GUI(,prevHK,,,"Atlasy " . HKid . " - Set Global Hotkey")
    
    If ErrorLevel ; Cancelled
        return
    If (newHK = prevHK) ; no change
        return
    
    RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, AtlasyHotkey%HKid%, %newHK%
    
    If (newHK = "") { ; reset/ disable hotkey
        ;Turn off the new hotkey.
        Hotkey, %prevHK%, Atlasy_%HKid%, Off 
        TipText = Set Atlasy %HKid% Hotkey off!
        TrayTipAutoHide("Atlasy " . HKid . " Hotkey Off",TipText,2000)
        return
    }
    
    ; Turn off the old Hotkey
    If Not (prevHK == "")
        Hotkey, %prevHK%, Atlasy_%HKid%, Off
    
    Atlasy_HotkeyActivate(HKid,newHK, True)
    
    ; Refresh Menus (function is only called from Script Menu)
    Reload
    
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Atlasy_HotkeyActivate(HKid,HK,showTrayTip := False) {
    ;Turn on the new hotkey.
    Hotkey, IfWinActive, ; for all windows/ global hotkey
    Hotkey, $%HK%, Atlasy_%HKid%, On ; use $ to avoid self-referring hotkey if Ctrl+Shift+M is used
    If (showTrayTip) {
        TipText = Atlasy %HKid% Hotkey set to %HK%
        TrayTipAutoHide("Atlasy " . HKid . " Hotkey On",TipText,2000)
    }
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Atlasy_OpenIssueDoc() {
; ^+v:: ; <--- Open Issue in R4J Document
    If GetKeyState("Ctrl") and !GetKeyState("Shift") {
        PowerTools_OpenDoc("r4j_openissuedoc") 
        return
    }
    If WinActive("ahk_exe EXCEL.EXE") {
        sKey := Jira_Excel_GetIssueKeys()
        If (sKey="")
            return
        R4J_OpenIssues(sKey)
    } Else {
        R4J_OpenIssueSelection()
    }
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Atlasy_OpenIssue() {
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	PowerTools_OpenDoc("jira_openissue") 
	return
}
If Browser_WinActive() { ; switch ServiceDesk Requester <-> Jira Agent
	sUrl := Browser_GetUrl()
	If Jira_IsUrl(sUrl) {
		IssueKey := Jira_Url2IssueKey(sUrl)
		If InStr(IssueKey,"ECOSYS-") {
			If InStr(sUrl,"/portal/") {
				sUrl := "https://instartconsult.atlassian.net/browse/" . IssueKey
			} Else 
				sUrl := "https://instartconsult.atlassian.net/servicedesk/customer/portal/15/" . IssueKey

			Run, %sUrl%
			return
		}
	}
} Else If WinActive("ahk_exe EXCEL.EXE") {
	sKey := Jira_Excel_GetIssueKeys()
	If (sKey="")
		return
	Jira_OpenIssues(sKey)
} Else {
	Jira_OpenIssueSelection()
}
} ; eofun