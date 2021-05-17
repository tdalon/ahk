; You can compile it via running the Ahk2Exe command e.g. D:\Programs\AutoHotkey\Compiler\Ahk2Exe.exe /in "Teamsy.ahk" /icon "icons\Teams.ico"
;#Include <Teams>
;#SingleInstance force ; for running from editor

Outlooky(A_Args[1])
ExitApp

Outlooky(sInput){
    
If (!sInput) { ; empty
    WinId := GetMainWindow()
    WinActivate, ahk_id %WinId%
    return
}

FoundPos := InStr(sInput," ")  
If FoundPos {
    sKeyword := SubStr(sInput,1,FoundPos-1)
    sInput := SubStr(sInput,FoundPos+1)
} Else {
    sKeyword := sInput
    sInput =
}

Switch sKeyword
{
Case "w": ; Web App
    Switch sInput
    {
    Case "m": ; mail
        Run, https://outlook.office.com/mail/
    Case "c","cal":
        Run, https://outlook.office.com/calendar/
    Default:
        Run, https://outlook.office.com/
    }
    return
}

WinId := GetMainWindow()
WinActivate, ahk_id %WinId%

Switch sKeyword
{
Case "n": ; new
    Switch sInput
    {
    Case "e": ; email
        WinGetTitle Title, A
        If ! RegExMatch(Title,"^Inbox.* - Outlook$") {
            SendInput ^1 
            Sleep 500
        }
        SendInput ^n
    Case "m": ; meeting
        WinGetTitle Title, A
        If ! RegExMatch(Title,"^Calendar.* - Outlook$") {
             SendInput ^2 ; open calendar
             While !RegExMatch(Title,"^Calendar.* - Outlook$") {
                WinGetTitle Title, A
                Sleep 500
            }
        }
        SendInput ^n
    Case "t": ; new task
        WinGetTitle Title, A
        If ! RegExMatch(Title,"^To-Do.* Outlook$") {
            SendInput ^3 ; open tasks
            While !RegExMatch(Title,"^To-Do.* Outlook$") {
                WinGetTitle Title, A
                Sleep 500
            }
        }
        SendInput ^n
    Default: ; new item
        SendInput ^n
    }
    return
Case "c","cal":
    WinGetTitle Title, A
    If ! RegExMatch(Title,"^Calendar.* - Outlook$") {
        SendInput ^2 ; open calendar
    }
    return
Case "t": ; new Teams meeting
    WinGetTitle Title, A
    If ! RegExMatch(Title,"^Calendar.* - Outlook$") {
            SendInput ^2 ; open calendar
            While !RegExMatch(Title,"^Calendar.* - Outlook$") {
                WinGetTitle Title, A
                Sleep 500
            }
    }
    SendInput !h ; Home Tab
    Sleep 500
    SendInput y1
    return
Case "q","quit": ; quit
    sCmd = taskkill /f /im "outlook.exe"
    Run %sCmd%,,Hide 
    return
Case "r","restart": ; restart
    sCmd = taskkill /f /im "outlook.exe"
    Run %sCmd%,,Hide 
    Run, outlook.exe
    return
} ; End Switch
 
} ; End function     



; -------------------------------------------------------------------------------------------------------------------

GetMainWindow(Exe:="outlook.exe"){
ExeName := StrReplace(Exe,".exe","")
StringUpper, ExeName, ExeName , T ; uppercase first letter
RegProp = %ExeName%MainWinId
WinGet, WinCount, Count, ahk_exe %Exe%
;MsgBox %RegProp% : %WinCount% ; DBG
If (WinCount = 0) {
    Run, %Exe%
    WinWaitActive, ahk_exe %Exe%
    MainWinId := WinExist("A")
} Else If (WinCount = 1) {
    MainWinId := WinExist("ahk_exe " . Exe)
} Else If (WinCount > 0) { ; fall-back multiple windows
    MainWinId := PowerTools_RegRead(%RegProp%)
    If WinExist("ahk_id " . MainWinId) 
        return MainWinId
    SetTitleMatchMode, RegEx  
    WinGet,MainWinId,ID,.* - Outlook$
}
PowerTools_RegWrite(%RegProp%,MainWinId)
return MainWinId

} ; eofun

