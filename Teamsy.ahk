; Homepage: https://tdalon.github.io/ahk/Teamsy
; You need to compile it to run from the command line
; You can compile it via running the Ahk2Exe command e.g. D:\Programs\AutoHotkey\Compiler\Ahk2Exe.exe /in "Teamsy.ahk" /icon "icons\Teams.ico"
; You can also compile it from the PowerTools Bundler
LastCompiled = 20220425160053
#Include <PowerTools>
#Include <Teamsy>

#SingleInstance force ; for running from editor

If (A_Args.Length() = 0)  { 
    PowerTools_MenuTray()
    PowerTools_Help("Teamsy") ; open help page
    TrayTip, Teamsy, See help. Script shall not be run standalone! 

    ; Tooltip
    If !a_iscompiled 
        FileGetTime, LastMod , %A_ScriptFullPath%
    Else 
        LastMod := LastCompiled
    FormatTime LastMod, %LastMod% D1 R

    sTooltip = Teamsy %LastMod%`nRight-Click on icon to access help/support.
    Menu, Tray, Tip, %sTooltip%
} ; end icon tray

If (A_Args.Length() = 1) 
    Teamsy(A_Args[1])
Else { ; because of Launchy Runner Plugin not handling "$$" properly
    If (A_Args.Length() = 2) 
        Teamsy(A_Args[1] . " " . A_Args[2]) 
    Else If (A_Args.Length() = 3) 
        Teamsy(A_Args[1] . " " . A_Args[2] . " " . A_Args[3]) 
    Else
        MsgBox Error: Wrong argument number
}