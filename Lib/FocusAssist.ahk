; See documentation https://tdalon.blogspot.com/2023/09/autohotkey-focus-assist.html
LastCompiled = 
#SingleInstance force ; for running from editor
#Include <UIA_Interface>

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

    sTooltip = FocusAssist %LastMod%`nRight-Click on icon to access help/support.
    Menu, Tray, Tip, %sTooltip%
    return
} ; end icon tray

If (A_Args.Length() > 0) 
    FocusAssist(A_Args[1])

ExitApp


/* 
#f::
; overwrite feedback hub
FocusAssist("+")
return

#o::
FocusAssist("-")
return 
*/


;  ##################  FUNCTIONS ###########################



; ------------------------------------------------------------------------------

FocusAssist(sInput){


; Open Focus Assist Settings
; Win+R 'Focus Assist'
; Alternative open url ms-settings:quiethours  
; https://support.microsoft.com/en-us/windows/make-it-easier-to-focus-on-tasks-0d259fd9-e9d0-702c-c027-007f0e78ea
Run, ms-settings:quiethours

; Wait for WinTitle=Settings
WinWaitActive,Settings

UIA := UIA_Interface()

WinId := WinActive("A")
UIAEl := UIA.ElementFromHandle(WinId) 

Switch sInput
{
Case "f","-": ; Off
    Filter := "AutomationId=Microsoft.QuietHoursProfile.Unrestricted_Button"
Case "p","+": ; Priority Only
    Filter := "AutomationId=Microsoft.QuietHoursProfile.PriorityOnly_Button"
Case "a": ; Alarm Only
    Filter := "AutomationId=Microsoft.QuietHoursProfile.AlarmsOnly_Button"
}

Btn := UIAEl.WaitElementExist(Filter)
Btn.Click()

Send !{f4} ; Close settings window


} ; eofun



; ------------------------------------------------------------

