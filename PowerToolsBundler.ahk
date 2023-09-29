#SingleInstance force

#Include <PowerTools>
#Include <AHK>
#Include <Login>
; Calls: Lib/ToStartup

LastCompiled = 20220314141503

;IcoFile := RegExReplace(A_ScriptFullPath,"\..*",".ico")
IcoFile  := PathX(A_ScriptFullPath, "Ext:.ico").Full
If (FileExist(IcoFile)) 
	Menu,Tray,Icon, %IcoFile%

AppList = ConnectionsEnhancer,NWS,OutlookShortcuts,PeopleConnector,TeamsShortcuts,TeamsyLauncher,Teamsy,Atlasy,Mute,Cursor Highlighter
Config := PowerTools_GetConfig()
If (Config = "Conti") 
    AppList = MO,%AppList%


sGuiTitle = PowerTools Bundler
Gui, PTBundler:New,,%sGuiTitle%

; ListView
Gui, Add, ListView, h200 w180 Checked, %sGuiTitle%  ; Create a ListView.
ImageListID := IL_Create(9)  ; Create an ImageList to hold 10 small icons.
LV_SetImageList(ImageListID)  ; Assign the above ImageList to the current ListView.

IconCount := 0
Loop, Parse, AppList, `,
{  ; Load the ImageList with a series of icons from the DLL.
    If a_iscompiled 
        IcoFile = %A_ScriptDir%\%A_LoopField%.exe
    Else 
        IcoFile = %A_ScriptDir%\%A_LoopField%.ico
        
    If FileExist(IcoFile) {
        IL_Add(ImageListID,IcoFile)
        IconCount := IconCount +1
        LV_Add("Icon" . IconCount , A_LoopField)
    } Else
        LV_Add("Icon10" , A_LoopField)
} ; End Loop     


LV_ModifyCol()  ; Auto-adjust the column widths.

; https://jacksautohotkeyblog.wordpress.com/2019/12/30/use-autohotkey-gui-menu-bar-for-instant-hotkeys/

; MenuBar
Menu, ItemsMenu, Add, Check All`tCtrl+A, PTBSelectAll
Menu, ItemsMenu, Add, Uncheck all`tCtrl+Shift+A, PTBUncheckAll

If !a_iscompiled {
    Menu, ActionsMenu, Add, &Compile`tCtrl+C, Compile
    Menu, ActionsMenu, Add, Compile And &Push`tCtrl+P, CompileAndPush
    Menu, ActionsMenu, Add, &Tweet`tCtrl+T, Tweet
    Menu, ActionsMenu, Add, Compile Bundler, CompileSelf
    Menu, ActionsMenu, Add, &Developper Mode`tCtrl+D, DevMode
    Menu, ActionsMenu, Add, Exe Mode, ExeMode
} Else {
    Menu, ActionsMenu, Add, Check for &Update/Download`tCtrl+U, CheckForUpdate
}
Menu, ActionsMenu, Add, Add to &Startup`tCtrl+S, AddToStartup
Menu, ActionsMenu, Add, &Run`tCtrl+R, Run
Menu, ActionsMenu, Add, E&xit`tCtrl+X, Exit
Menu, ActionsMenu, Add, Open Help`tCtrl+H, OpenHelp
Menu, ActionsMenu, Add, Open Change&log`tCtrl+L, OpenChangelog
Menu, ActionsMenu, Add, Open &News`tCtrl+N, OpenNews

Menu, SettingsMenu, Add, Load Config, LoadConfig
Menu, SettingsMenu, Add, Open ini file, OpenIni

Menu, SettingsMenu, Add, Notification at Startup, MenuCb_ToggleSettingNotificationAtStartup

RegRead, SettingNotificationAtStartup, HKEY_CURRENT_USER\Software\PowerTools, NotificationAtStartup
If (SettingNotificationAtStartup = "")
	SettingNotificationAtStartup := True ; Default value
If (SettingNotificationAtStartup) {
  Menu, SettingsMenu, Check, Notification at Startup
} Else {
  Menu, SettingsMenu, UnCheck, Notification at Startup
}

Menu, HelpMenu, Add, Open Help, OpenPTHelp
Menu, HelpMenu, Add, Check for Update, CheckForUpdateSelf
Menu, HelpMenu, Add, Open Change&log, OpenPTChangelog


Menu, MyMenuBar, Add, &Items, :ItemsMenu 
Menu, MyMenuBar, Add, &Actions, :ActionsMenu 
Menu, MyMenuBar, Add, &Settings, :SettingsMenu 
Menu, MyMenuBar, Add, &Help, :HelpMenu
Gui, Menu, MyMenuBar

Gui, Show
Return

; -------------------------------------------------------------------------------------------------------------------
OpenIni:
sIniFile = %A_ScriptDir%\PowerTools.ini 
If FileExist(sIniFile)
    Run, notepad.exe %sIniFile%
Return
; -------------------------------------------------------------------------------------------------------------------

LoadConfig:
PowerTools_LoadConfig()
return

PTBSelectAll:
LV_Modify(0, "Check")  ; Uncheck all the checkboxes.
return
PTBUncheckAll:
LV_Modify(0, "-Check")  ; Uncheck all the checkboxes.
return
; -------------------------------------------------------------------------------------------------------------------

CheckForUpdateSelf:
PowerTools_CheckForUptate()
return
; -------------------------------------------------------------------------------------------------------------------

CheckForUpdate: ; CFU
; warning if connected via VPN
If (Login_IsVPN()) {
    MsgBox, 0x1011, CheckForUpdate with VPN?,It seems you are connected with VPN.`nCheck for update might not work. Consider disconnecting VPN.`nContinue now?
    IfMsgBox Cancel
        return
}

RowNumber = 0
Loop {
    RowNumber := LV_GetNext(RowNumber, "Checked")
    if not RowNumber 
	    break
	
    LV_GetText(ItemName, RowNumber, 1)
	PowerTools_CheckForUptate(ItemName)
}

; Update PowerTools.ini - only once
IniFile =  %A_ScriptDir%\PowerTools.ini
sUrl = https://github.com/tdalon/ahk/raw/main/PowerTools.ini
If Not FileExist(IniFile) {
    UrlDownloadToFile, %sUrl%, %IniFile%
} Else {
    guExe = %A_ScriptDir%\github_updater.exe
    If Not FileExist(guExe)
        UrlDownloadToFile, https://github.com/tdalon/ahk/raw/main/PowerTools/github_updater.exe, %guExe%
    UrlDownloadToFile, %sUrl%, PowerTools.ini.github
    sCmd = %guExe% PowerTools.ini
    RunWait, %sCmd%,,Hide
}

; open directory
Run %A_ScriptDir%
return
; -------------------------------------------------------------------------------------------------------------------
OpenPTHelp:
PowerTools_Help("Bundler")
return
; -------------------------------------------------------------------------------------------------------------------
Exit: 
RowNumber = 0
Loop {
    RowNumber := LV_GetNext(RowNumber, "Checked")
    if not RowNumber 
	    break
	LV_GetText(ItemName, RowNumber, 1)
	If a_iscompiled 
	    ScriptFullPath = %A_ScriptDir%\%ItemName%.exe
    Else
	    ScriptFullPath = %A_ScriptDir%\%ItemName%.ahk

    AHK_Exit(ScriptFullPath)
}
return
; -------------------------------------------------------------------------------------------------------------------
OpenChangelog: 
RowNumber = 0
Loop {
    RowNumber := LV_GetNext(RowNumber, "Checked")
    if not RowNumber 
	    break
	LV_GetText(ItemName, RowNumber, 1)
	PowerTools_Changelog(ItemName)
}
return

; -------------------------------------------------------------------------------------------------------------------
OpenNews: 
RowNumber = 0
Loop {
    RowNumber := LV_GetNext(RowNumber, "Checked")
    if not RowNumber 
	    break
	LV_GetText(ItemName, RowNumber, 1)
	PowerTools_News(ItemName)
}
return
; -------------------------------------------------------------------------------------------------------------------
OpenHelp: 
RowNumber = 0
Loop {
    RowNumber := LV_GetNext(RowNumber, "Checked")
    if not RowNumber 
	    break
	LV_GetText(ItemName, RowNumber, 1)
	PowerTools_Help(ItemName)
}
return
; -------------------------------------------------------------------------------------------------------------------
Compile:
RowNumber = 0
Loop {
    RowNumber := LV_GetNext(RowNumber, "Checked")
    if not RowNumber 
	    break
	LV_GetText(ItemName, RowNumber, 1)
	ScriptFullPath = %A_ScriptDir%\%ItemName%.ahk
    AHK_Compile(ScriptFullPath,,"PowerTools")
	}
Run %A_ScriptDir%\PowerTools
return
; -------------------------------------------------------------------------------------------------------------------

CompileSelf:
AHK_Compile(A_ScriptFullPath,,"PowerTools")
return
; -------------------------------------------------------------------------------------------------------------------
CompileAndPush:
RowNumber = 0
FileList =
Loop {
    RowNumber := LV_GetNext(RowNumber, "Checked")
    if not RowNumber 
	    break
	LV_GetText(ItemName, RowNumber, 1)
	ScriptFullPath = %A_ScriptDir%\%ItemName%.ahk
    AHK_Compile(ScriptFullPath,,"PowerTools")
    FileList =  %FileList% %ItemName%.exe
    ; Add changelog
    cl := PowerTools_Changelog(ItemName,False)
    cl := RegExReplace(cl,".*\","")
    FileList =  %FileList% %cl%
} 
RunWait, git add %FileList%, %A_ScriptDir%\PowerTools
RunWait, git commit -m "Update compiled powertools", %A_ScriptDir%\PowerTools
RunWait, git push origin master, %A_ScriptDir%\PowerTools
return

Tweet:
RowNumber = 0
FileList =
Loop {
    RowNumber := LV_GetNext(RowNumber, "Checked")
    if not RowNumber 
	    break
	LV_GetText(ItemName, RowNumber, 1)
	PowerTools_TweetPush(ItemName)
} 
return
; -------------------------------------------------------------------------------------------------------------------
AddToStartup:
RowNumber = 0
Loop % LV_GetCount()
{
    RowNumber := A_Index
	SendMessage, 4140, RowNumber - 1, 0xF000, SysListView321  ; 4140 is LVM_GETITEMSTATE. 0xF000 is LVIS_STATEIMAGEMASK.
    IsChecked := (ErrorLevel >> 12) - 1  ; This sets IsChecked to true if RowNumber is checked or false otherwise.
    LV_GetText(ItemName, RowNumber, 1)
    If (ItemName = "Teamsy") ; Skip Teamsy -> TeamsyLauncher shall be used instead
        Continue
    
    If a_iscompiled 
	    ScriptFullPath = %A_ScriptDir%\%ItemName%.exe
    Else
	    ScriptFullPath = %A_ScriptDir%\%ItemName%.ahk
        
    ToStartup(ScriptFullPath,IsChecked)
}
Run %A_Startup%
return


; -------------------------------------------------------------------------------------------------------------------
Run:
RowNumber = 0
Loop % LV_GetCount()
{
    RowNumber := A_Index
	SendMessage, 4140, RowNumber - 1, 0xF000, SysListView321  ; 4140 is LVM_GETITEMSTATE. 0xF000 is LVIS_STATEIMAGEMASK.
    IsChecked := (ErrorLevel >> 12) - 1  ; This sets IsChecked to true if RowNumber is checked or false otherwise.
    LV_GetText(ItemName, RowNumber, 1)

    If a_iscompiled 
	    ScriptFullPath = %A_ScriptDir%\%ItemName%.exe
    Else
	    ScriptFullPath = %A_ScriptDir%\%ItemName%.ahk
    If IsChecked {
        If !FileExist(ScriptFullPath) {
            TrayTip, PowerTools Bundler: Run, % ScriptFullPath " not available!"
            break
        }
        Run, %ScriptFullPath%
    } Else
        AHK_Exit(ScriptFullPath)
}
return

; -------------------------------------------------------------------------------------------------------------------
StartMO:
Run %A_ScriptDir%\PowerTools\MO.exe
return
; -------------------------------------------------------------------------------------------------------------------
OpenPTChangelog:
PowerTools_Changelog("Bundler")
return 
; -------------------------------------------------------------------------------------------------------------------

DevMode:
;DevList = ConnectionsEnhancer,NWS,OutlookShortcuts,PeopleConnector,TeamsShortcuts
Loop, Parse, AppList, `,
{  
    ScriptFullPath = %A_ScriptDir%\%A_LoopField%.ahk
    Run, %ScriptFullPath%
    ScriptFullPath = %A_ScriptDir%\%A_LoopField%.exe
    SplitPath,ScriptFullPath, FileName
    sCmd = taskkill /f /im "%FileName%"
    Run %sCmd%,,Hide     
} ; End Loop
;Run %A_ScriptDir%\PowerTools\MO.exe
return 

ExeMode:
Loop, Parse, AppList, `,
{  
    ScriptFullPath = %A_ScriptDir%\%A_LoopField%.ahk
    AHK_Exit(ScriptFullPath)        
    ScriptFullPath = %A_ScriptDir%\PowerTools\%A_LoopField%.exe
    Run, %ScriptFullPath%
} ; End Loop
return
; -------------------------------------------------------------------------------------------------------------------
MenuCb_ToggleSettingNotificationAtStartup:
If (SettingNotificationAtStartup := !SettingNotificationAtStartup) {
  Menu, SettingsMenu, Check, Notification at Startup
}
Else {
  Menu, SettingsMenu, UnCheck, Notification at Startup
}
PowerTools_RegWrite("NotificationAtStartup",SettingNotificationAtStartup)
return
; -------------------------------------------------------------------------------------------------------------------

PTBundlerGuiClose:
ExitApp
PTBundlerGuiEscape:
ExitApp

return    
