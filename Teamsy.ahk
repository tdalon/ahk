; Homepage: 
; You can compile it via running the Ahk2Exe command e.g. D:\Programs\AutoHotkey\Compiler\Ahk2Exe.exe /in "Teamsy.ahk" /icon "icons\Teams.ico"
LastCompiled = 20210416063521
#Include <Teams>
#Include <Teamsy>
#Include <Monitor>

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

If (A_Args.Length() > 0)
    Teamsy(A_Args[1])
return

