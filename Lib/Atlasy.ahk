;#Include <Jira>
;#Include <Confluence>

; Launcher for Atlassian related Tools: Jira, Confluence, Bitbucket, BigPicture, R4J, Xray 
; See documentation of commands in Atlasy.md

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
        hWin := WinExist("A")
        sCmd := AtlasyInputBox()
        WinActivate, ahk_id %hWin% ; required because active window looses focus after the Gui closes
        if ErrorLevel
            return
        sCmd := Trim(sCmd) 
        Atlasy(sCmd)
        return
    Case "h","-h","help":
        Atlasy_Help(sInput)
        return
    Case "?":
        sFile:= A_ScriptDir . "\doc\Atlasy.md"
        Run, %sFile%
    Case "j": ;#jira
        If (sInput = "") {
            Jira_OpenIssues()
            return
        } 
        JiraRootUrl := Jira_GetRootUrl()
        IsCloud := Jira_IsCloud(JiraRootUrl)

        ; Short navigation keys
        If (sInput = "-b") or (sInput = "b") or RegExMatch(sInput,"^\-?b\s(.*)",sMatch) {
            If IsCloud
                sUrl := JiraRootUrl . "/jira/boards?contains=" . sMatch1
            Else
                sUrl := JiraRootUrl . "/secure/ManageRapidViews.jspa"
            Atlasy_OpenUrl(sUrl)
            return
        }

        If RegExMatch(sInput,"^\-?s\s(.*)",sMatch) {
            Jira_QuickSearch(sMatch1)
            Return
        }

        If (sInput = "-l") or (sInput = "l") { ; AddLink
            ;IssueKeys := Jira_GetIssueKeys()
            sLog := Jira_AddLink("",Jira_GetIssueKeys()) ; user will be prompted for link name and target issues
            If !(sLog = "") {
                ;TrayTipAutoHide("e.CoSys PowerTool",Text)
                ;TrayTip, e.CoSys PowerTool, %Text%
                OSDTIP_Pop("PowerTool: Link(s) added!",sLog)
            }
            return
        }

        If (sInput = "-c") or (sInput = "c") or RegExMatch(sInput,"^\-?c\s(.*)",sMatch) { ; create issue
            Jira_CreateIssue("",sMatch1)
            return
        }

        If (sInput = "b")  { ; bulk edit
            Jira_BulkEdit()
            return
        }

        If (sInput = "n")  { ; open issues in issue navigator
            Jira_OpenIssuesNav()
            return
        }
        
        ; issue Key
        sKeyPat := "i)^([A-Z\d]{3,}\-[\d]{1,})$"
        If RegExMatch(sInput,sKeyPat,sMatch)  {
            sUrl := JiraRootUrl . "/browse/" . sMatch1
            Atlasy_OpenUrl(sUrl)
            Return
        }

         ; Project Key
         sKeyPat := "i)^([A-Z\d]{3,})$"
         If RegExMatch(sInput,sKeyPat,sMatch)  {
             sUrl := JiraRootUrl . "/browse/" . sMatch1
             Atlasy_OpenUrl(sUrl)
             Return
         }

        If (sInput = "-i") or (sInput = "i") or RegExMatch(sInput,"^\-?i\s(.*)",sMatch) {
            sUrl := JiraRootUrl . "/issues/?jql=" . sMatch1
            Atlasy_OpenUrl(sUrl)
            Return
        }

        If (sInput = "-f") or (sInput = "f") or RegExMatch(sInput,"^\-?f\s(.*)",sMatch) {
            If IsCloud
                sUrl := JiraRootUrl . "/jira/filters?name=""" . sMatch1 . """"
            Else
                sUrl := JiraRootUrl . "/secure/ManageFilters.jspa"
            Atlasy_OpenUrl(sUrl)
            Return
        }
      
        ; dp; default project  
        If (sInput = "-dp") or (sInput = "dp") {
            PowerTools_SetSetting("JiraProject")
            return
        } 
        
        If RegExMatch(sInput,"^\-?dp\s(.*)",sMatch) {
            StringUpper, pj, sMatch1
            PowerTools_RegWrite("JiraProject",pj)
            return
        } 
        
        If (sInput = "-pl") or (sInput = "pl")  { ; Project List
            SettingName := "JiraProjects"
            Section := "Jira"
            Projects := PowerTools_IniRead(Section,SettingName)
	        If (Projects="ERROR")  
                Projects =""
            InputBox, Projects, PowerTools Setting, Enter %SettingName%:,, 250, 125, , , , ,%Projects%
            If ErrorLevel
                return
            StringUpper, Projects, Projects
            PowerTools_IniWrite(Projects,Section,SettingName)
            return
        } 

    Case "e": ;#eCoSys
       If (sInput = "-p") or (sInput = "p") {
            PowerTools_SetSetting("ECProject")
            return
        } Else If RegExMatch(sInput,"^\-?p\s(.*)",sMatch) {
            StringUpper, pj, sMatch1
            PowerTools_RegWrite("ECProject",pj)
        } Else If (sInput = "-pl") or (sInput = "pl") { ; Project List
            SettingName := "ECProjects"
            Section := "e.CoSys"
            Projects := PowerTools_IniRead(Section,SettingName)
	        If (Projects="ERROR")  
                Projects =""
            InputBox, Projects, PowerTools Setting, Enter %SettingName%:,, 250, 125, , , , ,%Projects%
            If ErrorLevel
                return
            StringUpper, Projects, Projects
            PowerTools_IniWrite(Projects,Section,SettingName)
        }

        return
    ;---------------------- 
    Case "r","r4j": ;#r4j
        If (sInput = "") {
            sIssueKey := R4J_GetIssueKey()
            If !(sIssueKey = "") {
                R4J_OpenIssue(sIssueKey)
                Return
            }   
        }
        
        view := "d" ; default
        Loop, Parse, % Trim(sInput), %A_Space%
        {
            If RegExMatch(A_LoopField,"^\-") { ; commands start with - e.g. -cv -cp
                cmd :=  A_LoopField 
                continue
            }

            Switch A_LoopField {                    
                Case "d","c","t":
                    view := A_LoopField
                Default:
                    If (cmd != "") {
                        cmd2 := A_LoopField
                    } Else {
                        pj := A_LoopField
                        StringUpper, pj, pj ; R4J is case sensitive for keys->convert to uppercase
                    }
                    
            } ; end sw
        }     

        If (cmd != "") {
            Switch cmd {
                Case "-cp","cp": ; Copy Path Jql
                    R4J_CopyPathJql()
                    return
                Case "-cc","cc": ; Copy Children Jql
                    R4J_CopyChildrenJql()
                    return
                Case "-p","p": ; Paste Migrated Jql
                    Jql := Clipboard    
                    NewJql := R4J_Migrate_Jql(Jql)
                    Clip_Paste(NewJql)
                    return
                Case "-n","n": ; Open in issue navigator r4jPath filter
                    R4J_OpenPathJql()
                    return
                Case "-cv","-c","-t","-tv":
                    If InStr(cmd,"c")
                        viewtype := "c"
                    Else
                        viewtype := "t"
                    Switch cmd2 {
                        Case "c","cp","mv":
                            R4J_View_Copy(viewtype)
                            return
                        Case "i":
                            R4J_View_Import(viewtype)
                            return
                        Case "d":
                            R4J_View_Delete(viewtype)
                            return
                        Case "e","x":
                            R4J_View_Export(viewtype)
                            return
                        Default:
                            R4J_View_Copy(viewtype)
                            return
                    }
            }
        }


        If (pj="") {
            pj := R4J_GetProjectDef()
        }

        If (RegExMatch(pj,"\d")) ; Issue Key
            R4J_OpenIssue(pj)
        Else
            R4J_OpenProject(pj,view)

    ;---------------------- 
    Case "x": ; #xray
        view := "r" ; default repository
        Loop, Parse, % Trim(sInput), %A_Space%
        {
            Switch A_LoopField
            {
                case "r","p", "e","trace","m","te","tp","tr","ts","t","c","cov":
                    view := A_LoopField
                case "gs":
                    Xray_Open("gs")
                    return
                case "doc":
                    Xray_Open(A_LoopField,RegExReplace(sInput,"doc "))
                    return
                Default:
                    pj := A_LoopField
                    StringUpper, pj, pj                    
            } ; end switch
        }    ; end loop
        Xray_Open(view,pj)
        return
    
    Case "c": ; #Confluence

        If (sInput = "") {
            Atlasy_OpenUrl(Confluence_GetRootUrl())
            return
        }
        
        If InStr(sInput,"order") {
            Confluence_Reorder()
            return
        }
            
        If RegExMatch(sInput,"^\-?s?\s(.*)",sMatch)  {
            Confluence_QuickSearch(sMatch1)
            Return
        }

        If RegExMatch(sInput,"^\-?o?\s(.*)",sMatch)  {
            Confluence_QuickOpen(sMatch1)
            Return
        }
        
        FoundPos := InStr(sInput," ")  
        If FoundPos {
            sSpace := SubStr(sInput,1,FoundPos-1)
            sQuery:= SubStr(sInput,FoundPos+1)
        } Else
            sSpace := sInput
        Confluence_SearchSpace(sSpace,sQuery)
    Case "bp":  ; #BigPicture
        JiraRootUrl := Jira_GetRootUrl()
        If InStr(JiraRootUrl,".atlassian,net") { ; cloud
           sUrl := JiraRootUrl . "/plugins/servlet/ac/eu.softwareplant.bigpicture/bigpicture" ;#!box/ROOT/o/hierarchy"
        } Else { ; server/dc
            sUrl := JiraRootUrl . "/plugins/servlet/ac/eu.softwareplant.bigpicture/bigpicture"
        }
        Atlasy_OpenUrl(sUrl)
        return

    Case "bb":  ; #BitBucket
        Url := PowerTools_IniRead("BitBucket","BitBucketRootUrl")
        If (BUrl="ERROR") {
            MsgBox 0x1010, Error, BitBucketRootUrl key not defined in BitBucket section in PowerTools.ini file!
	        return
        }
        If (sInput = "")
            defProject := PowerTools_GetSetting("ECProject")
        Else    
            defProject := sInput
        If !(defProject ="")
            Url := Url . "/projects/" . defProject
        
        Atlasy_OpenUrl(Url)
        return

    Case "sw": ; switch server <->cloud
        If !Browser_IsWinActive()   {
            TrayTipAutoHide("e.CoSys PowerTool","Switch only possible from browser window!")
            return
        } 
        sUrl := Browser_GetUrl()
        If (sUrl="") {
            return
        }
        If R4J_IsUrl(sUrl)
            sUrl := R4J_Redirect(sUrl)
        Else If BigPicture_IsUrl(sUrl)
            sUrl := BigPicture_Redirect(sUrl)
        Else If Confluence_IsUrl(sUrl)
            sUrl := Confluence_Redirect(sUrl)
        Else If Jira_IsUrl(sUrl)
            sUrl := Jira_Redirect(sUrl)

        Atlasy_OpenUrl(sUrl)
        return
    } ; end switch/case keyword
    
} ; eofun  
; -----------------------------------------


AtlasyInputBox(){
    static
    ButtonOK:=ButtonCancel:= false
	Gui GuiAtlasy:New ,, Atlasy
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
; -----------------------------------------



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
Atlasy_IsUrl(sUrl) {

If Jira_IsUrl(sUrl)
    return true
If Confluence_IsUrl(sUrl)
    return true
return false

} ; eofun
; -----------------------------------------

Atlasy_OpenUrl(sUrl) {
    ;MsgBox % sUrl ; DBG
    If (sUrl="")
        return
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

    sQuote = "
    sUrl := StrReplace(sUrl,sQuote,"%22") ; encode quotes for run cmd
    If (BrowserCmd = "")
        Run %sUrl%
    Else {
        sCmd = %BrowserCmd% "%sUrl%" ; Reading String in Ini files removes trailing quotes
	    Run, %sCmd%
    }

}
; -----------------------------------------

Atlasy_Help(sKeyword:=""){
    sUrl := PowerTools_OpenDoc("atlasy")
    
    Switch sKeyword 
    {
    Case "2c","oc":
        sUrl := ""
    Case "f","fav","f+","of": ; favorites
        sUrl := ""
    Case "cmd":
        
    } ; end switch

    Run, %sUrl%
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Atlasy_Launcher(){
    Atlasy("-g")
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Atlasy_HotkeySet(HKid){
    If GetKeyState("Ctrl")  { ; exclude ctrl if use in the hotkey
        dockey := "Atlasy_" . HKid
        StringLower dockey, dockey 
        PowerTools_OpenDoc(dockey)
        return
    }
    
    ; For Menu callback, remove ending Hotkey and blanks and (Hotkey)
    
    HKid := RegExReplace(HKid,"\t(.*)","")
    ;HKid := Trim(HKid) ; remove tab for align right hotkey in menu
    HKid := RegExReplace(HKid," Hotkey$","")
    HKid := StrReplace(HKid," ","")
    
    RegRead, prevHK, HKEY_CURRENT_USER\Software\PowerTools, AtlasyHotkey%HKid%
    newHK := Hotkey_GUI(,prevHK,,,"Atlasy " . HKid . " - Set Hotkey")
    
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
        HKrev := Hotkey_ParseRev(HK)
        TipText = Atlasy %HKid% Hotkey set to %HKrev%
        TrayTipAutoHide("Atlasy " . HKid . " Hotkey On",TipText,2000)
    }
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Atlasy_OpenIssueDoc() {
; ^+v:: ; <--- Open Issue in R4J Document
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
    PowerTools_OpenDoc("atlasy_openissuedoc") 
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
Jira_OpenIssues()
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Atlasy_DatePicker() { ; @fun_Atlassy_DatePicker@
    static date
    DatePicker(date)
    If !date ; empty= cancel
        return
    If Browser_IsWinActive() {
        Url := Browser_GetUrl()
        If Confluence_IsUrl(Url) {
            FormatTime, date, %date%, M/dd/yyyy
            SendInput /
            Sleep 200
            SendInput /
            Sleep 200
            SendInput {Enter}
            Sleep 200
            SendInput %date%
            SendInput {Esc}
            return
        } Else If Jira_IsUrl(Url) {
            FormatTime, date, %date%, M/dd/yyyy
            Clipboard := date
            TrayTipAutoHide("Atlasy","Date copied to clipboard!")
            return
        } 
    } ; eif browser
   
    DateFormat := PowerTools_IniRead("General","DateFormat")
    If (DateFormat="ERROR") 
        DateFormat := ""
    FormatTime, date, %date%, %DateFormat%
    SendInput %date%

} ; eofun
; -------------------------------------
DatePicker(ByRef DatePicker){
    ;static DatePicker
    Gui, +LastFound 
    gui_hwnd := WinExist()
    Gui, Add, MonthCal, 4 vDatePicker
    Gui, Add, Button, Default , &OK
    Gui Add, Button, x+0, Cancel
    Gui, Show , , Date Picker
    WinWaitClose, AHK_ID %gui_hwnd%
    return
    
    ButtonOK:
    Gui, submit ;, nohide
    Gui, Destroy
    ;Gui, Hide
    return 
    
    GuiEscape:
    ButtonCancel:
    GuiClose:
    Gui, Destroy
    return
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Atlasy_CurrentDate() { ; @fun_Atlassy_CurrentDate@
If Browser_IsWinActive() {
    Url := Browser_GetUrl()
    If Confluence_IsUrl(Url) {
        SendInput /
        Sleep 200
        SendInput /
        Sleep 200
        SendInput {Enter} 
        Sleep 200
        SendInput {Esc}
        return
    } Else If Jira_IsUrl(Url) {
        FormatTime, date, , M/dd/yyyy
        SendInput %date%
        return
    } 
} ; eif browser

DateFormat := PowerTools_IniRead("General","DateFormat")
If (DateFormat="ERROR") 
    DateFormat := ""
FormatTime, date, %date%, %DateFormat%
SendInput %date%

} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Atlasy_Redirect(sUrl) {
    If R4J_IsUrl(sUrl)
		sUrl := R4J_Redirect(sUrl)
	Else If BigPicture_IsUrl(sUrl)
		sUrl := BigPicture_Redirect(sUrl)
	Else If Jira_IsUrl(sUrl)
		sUrl := Jira_Redirect(sUrl)	
	Else If Confluence_IsUrl(sUrl)
		sUrl := Confluence_Redirect(sUrl)
	
	Atlasy_OpenUrl(sUrl)
} ;eofun