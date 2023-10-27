; Homepage: https://tdalon.github.io/ahk/Teamsy
; You need to compile it to run from the command line
; You can compile it via running the Ahk2Exe command e.g. D:\Programs\AutoHotkey\Compiler\Ahk2Exe.exe /in "Teamsy.ahk" /icon "icons\Teams.ico"
; You can also compile it from the PowerTools Bundler
LastCompiled = 20231026154346
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
    return
} ; end icon tray

; Loop, because of Launchy Runner Plugin not handling "..." properly as one input argument but splitting arguments at space
Loop % A_Args.Length() { 
    Arg := Arg . " " . A_Args[A_Index]
}
Teamsy(Trim(Arg))
