; See documentation https://tdalon.blogspot.com/2020/12/chromy.html
LastCompiled = 
#Include <Clip>
#SingleInstance force ; for running from editor

SetTitleMatchMode, 1 ; start with

If (A_Args.Length() = 0)  {
    ;RefreshProfiles()
    PowerTools_MenuTray()

    ; Tooltip
    If !a_iscompiled 
        FileGetTime, LastMod , %A_ScriptFullPath%
    Else 
        LastMod := LastCompiled
    FormatTime LastMod, %LastMod% D1 R

    sTooltip = Chromy %LastMod%`nRight-Click on icon to access help/support.
    Menu, Tray, Tip, %sTooltip%
    return
} ; end icon tray

If (A_Args.Length() > 0) 
    Chromy(A_Args[1],A_Args[2])

ExitApp



;  ##################  FUNCTIONS ###########################

Chromy(sInput,sProfileName:=""){
FoundPos := InStr(sInput," ")   
sKeyword := SubStr(sInput,1,FoundPos-1)
sInput := SubStr(sInput,FoundPos+1)
;SendInput, !{Esc} ; focus is lost from launcher
;sInput := StrReplace(sInput, "#", "{#}")
If (sProfileName == "" ){ ; no Profile
    If WinActive("ahk_exe chrome.exe") {   
        SendInput , ^t
        WinWaitActive, New Tab ahk_exe chrome.exe
    } Else {
        Run, chrome.exe
        WinWaitActive, New Tab ahk_exe chrome.exe,,2000 ; 
    }
    SendBar(sKeyword,sInput)
   
} Else { ; Profile passed as argument
       
    hWnd := Chrome_WinGet(sProfileName)
    ;MsgBox %hWnd% %sProfileName%
    If (!hWnd) { ; empty - first instance
        Chrome_NewTab(sProfileName)
        While !(Chrome_GetProfile() == sProfileName )
            Sleep 500
        SendInput ^t^{tab}^w ; close previous home window
        WinWaitActive, New Tab ahk_exe chrome.exe
        SendBar(sKeyword,sInput)
    } Else {
        WinActivate, ahk_id %hWnd%
        WinWaitActive ahk_id %hWnd%
        SendInput , ^t
        SendBar(sKeyword,sInput)
    }

} ; end of else Profile
} ; eofun


; ---------------------------------------------
SendBar(sKeyword,sInput){

If !(sInput)
    return
If !(sKeyword="")
    SendInput %sKeyword%{tab}
Clip_Paste(sInput)
SendInput {enter}
}


; ------------------------------------------------------------

Chrome_Close(sProfile :=""){
; Close all Chrome Windows matching input Profile.
; If no Profile input, close all Chrome Windows

WinGet, Win, List, ahk_exe Chrome.exe

Loop %Win% {
    WinId := Win%A_Index%
    If !(sProfile ="") {
        WinProfile := Chrome_GetProfile(WinId) 
        If Not (WinProfile = sProfile) 
            Continue
    }

    WinActivate, ahk_id %WinId%
    WinWaitActive, ahk_id %WinId% 
   
    SendInput !{f4}
} ; end loop
} ; eofun

; ------------------------------------------------------------

Chrome_GetProfile(hwnd:=""){
; sProfile := Chrome_GetProfile(hWnd)
; returns Profile Name
; hWnd: Window handle e.g. output of WinActive or WinExist
; If no argument is passed, will take current active window
If !hwnd
    hwnd := WinActive("A")

title := Chrome_Acc_ObjectFromWindow(hwnd).accName
RegExMatch(title, "^.+Google Chrome . .*?([^(]+[^)]).?$", match)
return match1
}

; https://stackoverflow.com/a/62954549/2043349

/*
Chrome_Acc_Init() {
    static h := DllCall("LoadLibrary", Str,"oleacc", Ptr)
}
*/

Chrome_Acc_ObjectFromWindow(hwnd, objectId := 0) {
    static OBJID_NATIVEOM := 0xFFFFFFF0

    objectId &= 0xFFFFFFFF
    If (objectId == OBJID_NATIVEOM)
        riid := -VarSetCapacity(IID, 16) + NumPut(0x46000000000000C0, NumPut(0x0000000000020400, IID, "Int64"), "Int64")
    Else
        riid := -VarSetCapacity(IID, 16) + NumPut(0x719B3800AA000C81, NumPut(0x11CF3C3D618736E0, IID, "Int64"), "Int64")

    If (DllCall("oleacc\AccessibleObjectFromWindow", Ptr,hwnd, UInt,objectId, Ptr,riid, PtrP,pacc:=0) == 0)
        Return ComObject(9, pacc, 1), ObjAddRef(pacc)
}

; ------------------------------------------------------------
RefreshProfiles() {
; Syntax: profiles := RefreshProfiles()
; profiles: ProfileName:ProfileDir comma separated
sParProfileDir := % SubStr(A_AppData, 1, -8) "\Local\Google\Chrome\User Data\"
sDir := % sParProfileDir "Default"
profileName := Chrome_GetProfileName(sDir)
profiles := profileName ":Default" 
Loop, Files, %sParProfileDir%Profile *, D 
{
    profileName := Chrome_GetProfileName(A_LoopFileFullPath)
    profiles .= "," profileName ":" A_LoopFileName
}
;MsgBox % profiles ; DBG
PowerTools_RegWrite("ChromeProfiles",profiles)
} ; eofun
; ------------------------------------------------------------

Chrome_ProfileName2Dir(sProfileName){
; Given Profile name, returns Profile Directory e.g. Default or Profile 1 based on ChromeProfiles Registry

profiles := PowerTools_RegRead("ChromeProfiles")

sPat := sProfileName ":([^,]*)"
If RegExMatch(profiles, sPat, sMatch)
    return sMatch1
;MsgBox %profiles%`n%sProfileName% -> %sMatch1% ; DBG
profiles := RefreshProfiles()

RegExMatch(profiles, sPat, sMatch)
;MsgBox %profiles% `n %sMatch1% ; DBG 

return sMatch1

} ; eofun

; ------------------------------------------------------------

Chrome_ProfileDir2Name(sDir){
; Given Profile directory, returns Profile Name based on ChromeProfiles Registry

profiles := PowerTools_RegRead("ChromeProfiles")
sPat :=  "([^,]*):" sDir
If RegExMatch(profiles, sPat, sMatch)
    return sMatch1
profiles := RefreshProfiles()
RegExMatch(profiles, sPat, sMatch)
return sMatch1

} ; eofun
; ------------------------------------------------------------

Chrome_GetProfileName(sDir){
; Get Profile Name from input Profile Directory.
; Directory can be full path or only folder name
; Loop for Default and Profile n Directory in % SubStr(A_AppData, 1, -8) "\Local\Google\Chrome\User Data\"
If !InStr(sDir,"\")
    sDir := % SubStr(A_AppData, 1, -8) "\Local\Google\Chrome\User Data\" sDir
FileRead prefs, % sDir "\Preferences"
sPat = "managed_users":[^,]*,"name":"([^"]*)"
RegExMatch(prefs,sPat,sMatch)  
return sMatch1  
}
; ------------------------------------------------------------

Chrome_ActivateProfile(sProfileName) {
; Activate Chrome Profile by Name
; If no windows exists, new Chrome window is opened
 

    WinGet hwndList, List, ahk_exe chrome.exe
    Loop % hwndList {
        hwnd := hwndList%A_Index%
       sProfile := Chrome_GetProfile(hwnd)

        If (sProfile == sProfileName) {
            WinActivate ahk_id %hwnd%
            Return
        }
    }
    Chrome_NewWindow(sProfileName)
}
; ------------------------------------------------------------

Chrome_WinGet(sProfileName) {
; Get Chrome Window matching Profile Name
; Returns hWnd of first window match
WinGet hwndList, List, ahk_exe chrome.exe
Loop % hwndList {
    hwnd := hwndList%A_Index%
    sProfile := Chrome_GetProfile(hwnd)
    If (sProfile == sProfileName) 
        return hwnd
}
} ; eofun
; ------------------------------------------------------------



Chrome_NewWindow(sProfileName) {
sProfileDir := Chrome_ProfileName2Dir(sProfileName)
;MsgBox %sProfileName% -> %sProfileDir%
; Run chrome.exe --profile-directory="%sProfileDir%"
sCmd = %comspec% /c start chrome.exe --profile-directory="%sProfileDir%"
Run %sCmd% ,,Hide     
} ; eofun
; ------------------------------------------------------------

Chrome_NewTab(sProfileName) {
sProfileDir := Chrome_ProfileName2Dir(sProfileName)
;MsgBox %sProfileName% -> %sProfileDir%
; Run chrome.exe --profile-directory="%sProfileDir%" /new-tab
sCmd = %comspec% /c start chrome.exe --profile-directory="%sProfileDir%"
Run %sCmd% ,,Hide     
}