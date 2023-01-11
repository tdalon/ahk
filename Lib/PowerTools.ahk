; PowerTools Lib

#Include <AHK>
#Include <Teams>

AppList = ConnectionsEnhancer,MO,NWS,OutlookShortcuts,PeopleConnector,TeamsShortcuts

PowerTools_CheckForUptate(ToolName :="") {
If !a_iscompiled {
	Run, https://github.com/tdalon/ahk ; no direct link because of Lib dependencies
    return
} 

; warning if connected via VPN
If (Login_IsVPN()) {
MsgBox, 0x1011, CheckForUpdate with VPN?,It seems you are connected with VPN.`nCheck for update might not work. Consider disconnecting VPN.`nContinue now?
IfMsgBox Cancel
    return
}

If !ToolName    
    ScriptName := A_ScriptName
Else
    ScriptName = %ToolName%.exe
    ; Overwrites by default
sUrl = https://github.com/tdalon/ahk/raw/main/PowerTools/%ScriptName%

ExeFile = %A_ScriptDir%\%ScriptName%
If Not FileExist(ExeFile) {
    UrlDownloadToFile, %sUrl%, %ScriptName%
    return
}

UrlDownloadToFile, %sUrl%, %ScriptName%.github
guExe = %A_ScriptDir%\github_updater.exe
If Not FileExist(guExe)
    UrlDownloadToFile, https://github.com/tdalon/ahk/raw/main/PowerTools/github_updater.exe, %guExe%
    
sCmd = %guExe% %ScriptName%
RunWait, %sCmd%,,Hide
} ; eof

; ---------------------------------------------------------------------- 
PowerTools_Help(ScriptName,doOpen := True){

Switch ScriptName 
{
Case "ConnectionsEnhancer":
    sUrl = https://tdalon.github.io/ahk/Connections-Enhancer
Case "TeamsShortcuts":
    sUrl = https://tdalon.github.io/ahk/Teams-Shortcuts
Case "PeopleConnector":
    sUrl = https://tdalon.github.io/ahk/People-Connector
Case "OutlookShortcuts":
    sUrl = https://tdalon.github.io/ahk/Outlook-Shortcuts
Case "Teamsy":
    sUrl = https://tdalon.github.io/ahk/Teamsy
Case "TeamsyLauncher":
    sUrl = https://tdalon.github.io/ahk/Teamsy-Launcher
Case "NWS":
    sUrl := "https://tdalon.github.io/ahk/NWS-PowerTool"
Case "Mute":
    sUrl := "https://tdalon.github.io/ahk/Mute-PowerTool"
Case "Bundler":
    sUrl :="https://tdalon.github.io/ahk/PowerTools-Bundler"
Case "Cursor Highlighter":
    sUrl = https://tdalon.github.io/ahk/Cursor-Highlighter
Case "Chromy":
    sUrl = https://tdalon.github.io/ahk/Chromy
Case "Edgy":
    sUrl = https://tdalon.github.io/ahk/Edgy
Case "all":
Default:
    sUrl := "https://tdalon.github.io/ahk/PowerTools"	
}

If doOpen
    Run, %sUrl%
return sUrl
} ; eofun

; ---------------------------------------------------------------------- 

PowerTools_Changelog(ScriptName,doOpen := True){
Switch ScriptName 
{
Case "ConnectionsEnhancer":
    sFileName = Connections-Enhancer-Changelog
Case "TeamsShortcuts":
    sFileName = Teams-Shortcuts-Changelog
Case "PeopleConnector":
    sFileName = People-Connector-Changelog
Case "NWS":
    sFileName = NWS-PowerTool-Changelog
Case "Mute":
    sFileName = Mute-PowerTool-Changelog
Case "Bundler":
    sFileName = PowerTools-Bundler-Changelog
Case "OutlookShortcuts":
    sFileName = Outlook-Shortcuts-Changelog
Case "Teamsy":
    sFileName = Teamsy-Changelog
Case "Cursor Highlighter":
    Run, https://sites.google.com/site/boisvertlab/computer-stuff/online-teaching/cursor-highlighter-changelog
    return
Case "all":
Default:
    sFileName =  PowerTools-Changelogs
}

If Not doOpen {
    sUrl = https://tdalon.github.io/ahk/%sFileName%
    return sUrl
}
    

If !A_IsCompiled {
    sFile = %A_ScriptDir%\docs\_pages\%sFileName%.md
    If FileExist(sFile) {
        ;Run, Open "%sFile%" ; does not open Atom
        Run notepad++.exe "%sFile%"
        Return
    }
} Else {
    sUrl = https://tdalon.github.io/ahk/%sFileName%
    Run, %sUrl%
}
} ; eofun

; ---------------------------------------------------------------------- 

PowerTools_News(ScriptName){
If (ScriptName ="NWS")
    ScriptName = NWSPowerTool
Else If (ScriptName ="Mute")
    ScriptName = MutePowerTool
sUrl := "https://twitter.com/search?q=(from%3Atdalon)%23" . ScriptName
Switch ScriptName 
{
Case "ConnectionsEnhancer":
    sUrl := sUrl . "%20%23Connections"
Case "TeamsShortcuts":
    sUrl := sUrl . "%20%23MicrosoftTeams"
Case "OutlookShortcuts":
    sUrl := sUrl . "%20%23MicrosoftOutlook"
Case "Teamsy":
    sUrl := sUrl . "%20%23MicrosoftTeams"
}
Run, %sUrl%

} ; eofun

; ---------------------------------------------------------------------- 

PowerTools_RunBundler(){
If a_iscompiled {
  ExeFile = %A_ScriptDir%\PowerToolsBundler.exe
  If Not FileExist(ExeFile) {
    sUrl = https://raw.githubusercontent.com/tdalon/ahk/main/PowerTools/PowerToolsBundler.exe
		UrlDownloadToFile, %sUrl%, PowerToolsBundler.exe
  }
  Run %ExeFile%
} Else
  Run %A_AHKPath% "%A_ScriptDir%\PowerToolsBundler.ahk"

} ; eofun
; ---------------------------------------------------------------------- 




; ---------------------------------------------------------------------- 

PowerTools_OpenDoc(key:=""){
RegRead, PT_DocRootUrl, HKEY_CURRENT_USER\Software\PowerTools, DocRootUrl
If (key ="") 
    sUrl := "https://github.com/tdalon/ahk"
Else {
    If InStr(PT_DocRootUrl,".blogspot.") 
        key := StrReplace(key,"_","-")
    sUrl = %PT_DocRootUrl%/%key% 
}
Run,  "%sUrl%"
} ; eofun



; ----------------------------------------------------------------------
PowerTools_GetSetting(SettingName){
RegRead, Setting, HKEY_CURRENT_USER\Software\PowerTools, %SettingName%
If (Setting=""){
    Setting := PowerTools_SetSetting(SettingName)
}
return Setting
} ; eofun
; ----------------------------------------------------------------------
PowerTools_SetSetting(SettingName){
; for call from Menu with Name Set <Setting>
SettingName := RegExReplace(SettingName,"^Set ","") 
SettingProp := RegExReplace(SettingName," ","") ; Remove spaces 

RegRead, Setting, HKEY_CURRENT_USER\Software\PowerTools, %SettingProp%
InputBox, Setting, PowerTools Setting, Enter %SettingName%:,, 250, 125, , , , ,%Setting%
If ErrorLevel
    return
PowerTools_RegWrite(SettingProp,Setting)
return Setting
} ; eofun

; ----------------------------------------------------------------------
PowerTools_GetConfig(){
RegRead, Config, HKEY_CURRENT_USER\Software\PowerTools, Config
If (Config=""){
    Config := PowerTools_LoadConfig()
}
return Config
}
; ----------------------------------------------------------------------
PowerTools_SetConfig(){
RegRead, Config, HKEY_CURRENT_USER\Software\PowerTools, Config
DefListConfig := "Default|Ini"
If FileExist("Lib/ET.ahk")
    DefListConfig := DefListConfig . "|ET"
If FileExist("Lib/Conti.ahk")
    DefListConfig := DefListConfig . "|Conti|Vitesco"
Select := 0
Loop, parse, DefListConfig, | 
{
    If (A_LoopField = Config) {
        Select := A_Index
        break
    }
}
Config := ListBox("PowerTools Config","Select your configuration:",DefListConfig,Select)
If (Config="")
    return
PowerTools_RegWrite("Config",Config)
return Config
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
PowerTools_RegRead(Prop){
RegRead, OutputVar, HKEY_CURRENT_USER\Software\PowerTools, %Prop%
return OutputVar
}

; -------------------------------------------------------------------------------------------------------------------
PowerTools_RegWrite(Prop, Value){
RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, %Prop%, %Value%    
}

; -------------------------------------------------------------------------------------------------------------------
PowerTools_RegGet(Prop,sPrompt :=""){
Prop := StrReplace(Prop," ","") ; remove blanks
RegRead, Value, HKEY_CURRENT_USER\Software\PowerTools, %Prop% 
If (sPrompt = "")
    sPrompt = Enter value for %Prop%:
InputBox, Value, %Prop%, %sPrompt%, , 200, 150, , , , , Value 
If ErrorLevel
    return
PowerTools_RegWrite(Prop,Value)
return Value
} ;eofun

; -------------------------------------------------------------------------------------------------------------------
PowerTools_RegSet(Prop,sPrompt :=""){

If (sPrompt = "")
    sPrompt = Enter value for %Prop%:

Prop := StrReplace(Prop," ","")
RegRead, Value, HKEY_CURRENT_USER\Software\PowerTools, %Prop%
InputBox, Value, %Prop%, %sPrompt%, , 200, 150, , , , , %Value%
If ErrorLevel
    return
PowerTools_RegWrite(Prop,Value)

return Value
} ;eofun

; ----------------------------------------------------------------------
PowerTools_LoadConfig(Config :=""){
If (Config=""){
    Config := PowerTools_SetConfig()
}

IniFile = %A_ScriptDir%\PowerTools.ini

Switch Config
{
    Case "Default": ; Global default settings
        sEmpty=
        PowerTools_RegWrite("Domain","")
        IniWrite,%sEmpty% , %IniFile%, Main, Domain

        PowerTools_RegWrite("TenantName","")
        IniWrite, %sEmpty%, %IniFile%, Main, TenantName

        PowerTools_RegWrite("ProxyServer","n/a")
        IniWrite, n/a, %IniFile%, Main, ProxyServer

        PowerTools_RegWrite("ConnectionsRootUrl","")
        IniWrite, %sEmpty%, %IniFile%, Connections, ConnectionsRootUrl

        PowerTools_RegWrite("TeamsOnly",1)
        IniWrite, 1, %IniFile%, Teams, TeamsOnly

        PowerTools_RegWrite("DocRootUrl","https://tdalon.blogspot.com/")
        IniWrite, https://tdalon.blogspot.com/, %IniFile%, Main, DocRootUrl


        ; Load Parameters
        ParamList = TeamsMentionDelay,TeamsCommandDelay,TeamsClickDelay
        ; FindText
        TeamsFindTextList = Mute,Muted,Leave,Resume,MeetingActions,MeetingReactions,MeetingReactionHeart,MeetingReactionLaugh,MeetingReactionApplause,MeetingReactionLike,MeetingActionFullScreen,MeetingActionTogetherMode,MeetingActionBackgrounds,MeetingActionShare,MeetingActionUnShare
        Loop, Parse,TeamsFindTextList, `,
        {
            ParamList = %ParamList%,TeamsFindText%A_LoopField%
        }

        Loop, Parse, ParamList, `,
        {
            If RegExMatch(A_LoopField,"[A-Z][a-z]*",sMatch)
            {
                IniVal := PowerTools_GetParam(A_LoopField)
                IniWrite, %IniVal%, %IniFile%, %sMatch%, %A_LoopField%
            } 
        }

        ; Teams Global Hotkeys
        HotkeyIDList = Launcher,Mute,Video,Mute App,Share,Raise Hand,Push To Talk 
        Loop, Parse, HotkeyIDList, `,
        {
            HKid := A_LoopField
            HKid := StrReplace(HKid," ","")
            Param = TeamsHotkey%HKid%
            RegRead, HK, HKEY_CURRENT_USER\Software\PowerTools, %Param%
            IniWrite, %HK%, %IniFile%, Teams, %Param%
        }


    Case "Ini":
        If !FileExist(IniFile) {
            MSgBox 0x10, PowerTools: Error, PowerTools.ini can not be found!
            return
        }

        IniRead, IniVal, %IniFile%, Main, Domain
        PowerTools_RegWrite("Domain",IniVal)
        IniRead, IniVal, %IniFile%, Teams, TeamsOnly
        PowerTools_RegWrite("TeamsOnly",IniVal)
        IniRead, IniVal, %IniFile%, Main, DocRootUrl
        PowerTools_RegWrite("DocRootUrl",IniVal)
        IniRead, IniVal, %IniFile%, Connections, ConnectionsRootUrl
        If (IniVal != "ERROR")
            PowerTools_RegWrite("ConnectionsRootUrl",IniVal)
        Else
            PowerTools_RegWrite("ConnectionsRootUrl","")
        
        IniRead, IniVal, %IniFile%, Jira, JiraUserName
        If (IniVal != "ERROR")
            PowerTools_RegWrite("JiraUserName",IniVal)
        IniRead, IniVal, %IniFile%, Confluence, ConfluenceUserName
        If (IniVal != "ERROR")
            PowerTools_RegWrite("ConfluenceUserName",IniVal)
        

        ; Load Parameters
        ParamList = TeamsMentionDelay,TeamsCommandDelay,TeamsClickDelay
        ; FindText
        TeamsFindTextList = MeetingActions,MeetingReactions,MeetingReactionHeart,MeetingReactionLaugh,MeetingReactionApplause,MeetingReactionLike,MeetingActionFullScreen,MeetingActionTogetherMode,MeetingActionBackgrounds,MeetingActionShare,MeetingActionUnShare
        Loop, Parse,TeamsFindTextList, `,
        {
            ParamList = %ParamList%,TeamsFindText%A_LoopField%
        }

        Loop, Parse, ParamList, `,
        {
            If RegExMatch(A_LoopField,"[A-Z][a-z]*",sMatch)
            {
                IniRead, IniVal, %IniFile%, %sMatch%, %A_LoopField%
                If (IniVal != "ERROR")
                    PowerTools_RegWrite(A_LoopField,IniVal)
            } 
        }

        ; Teams Global Hotkeys
        HotkeyIDList = Launcher,Mute,Video,Mute App,Share,Raise Hand,Push To Talk 
        Loop, Parse, HotkeyIDList, `,
        {
            HKid := A_LoopField
            HKid := StrReplace(HKid," ","")
            Param = TeamsHotkey%HKid%

            IniRead, IniVal, %IniFile%, Teams, %Param%
            If (IniVal != "ERROR")
                PowerTools_RegWrite(Param,IniVal)
        }
    Case "ET":
	    If FileExist("Lib/ET.ahk") 
            FunStr := "ET_LoadConfig"
			%FunStr%()
    Case "Conti","Vitesco":        
	    If FileExist("Lib/Conti.ahk")
            FunStr := "Conti_LoadConfig"
			%FunStr%(Config)
} ; end switch

} ; eofun

; ----------------------------------------------------------------------

; ----------------------------------------------------------------------
PowerTools_CursorHighlighter(){
CHFile= %A_ScriptDir%\Cursor Highlighter
If a_iscompiled
	CHFile = %CHFile%.exe
Else
	CHFile = %CHFile%.ahk

If !FileExist(CHFile) { ; download if it doesn't exist
    return
}
Run, %CHFile%
}
; -------------------------------------------------------------------------------------------------------------------

PowerTools_MenuTray(){
; SubMenuSettings := PowerTools_MenuTray()
Menu, Tray, NoStandard
Menu, Tray, Add, &Help, MenuCb_PTHelp
Menu, Tray, Add, Tweet for support, MenuCb_PowerTools_Tweet
Menu, Tray, Add, Check for update, MenuCb_PTCheckForUpdate
Menu, Tray, Add, Changelog, MenuCb_PTChangelog
Menu, Tray, Add, News, MenuCb_PTNews


If !a_iscompiled {
    IcoFile  := PathX(A_ScriptFullPath, "Ext:.ico").Full
	If (FileExist(IcoFile)) 
		Menu,Tray,Icon, %IcoFile%
}

If (A_ScriptName = "Teamsy.exe") or (A_ScriptName = "Teamsy.ahk")
    return

; -------------------------------------------------------------------------------------------------------------------
; SETTINGS
Menu, SubMenuSettings, Add, Launch on Startup, MenuCb_ToggleSettingLaunchOnStartup
SettingLaunchOnStartup := ToStartup(A_ScriptFullPath)
If (SettingLaunchOnStartup) 
  Menu,SubMenuSettings,Check, Launch on Startup
Else 
  Menu,SubMenuSettings,UnCheck, Launch on Startup

Menu, Tray, Add, Settings, :SubMenuSettings

Menu,Tray,Add
Menu,Tray,Standard
Menu,Tray,Default,&Help

return SubMenuSettings
}

; ---------------------------------------------------------------------- STARTUP -------------------------------------------------
MenuCb_ToggleSettingLaunchOnStartup(ItemName, ItemPos, MenuName){
SettingLaunchOnStartup := !ToStartup(A_ScriptFullPath)
If (SettingLaunchOnStartup) {
 	Menu,%MenuName%,Check, %ItemName%	 
	ToStartup(A_ScriptFullPath,True)
}
Else {
    Menu,%MenuName%,UnCheck, %ItemName%	 
	ToStartup(A_ScriptFullPath,False)
}
}

MenuCb_PTHelp(ItemName, ItemPos, MenuName){
ScriptName := RegExReplace(A_ScriptName,"\..*","")
PowerTools_Help(ScriptName)
}

MenuCb_PTChangelog(ItemName, ItemPos, MenuName){
ScriptName := RegExReplace(A_ScriptName,"\..*","") 
PowerTools_Changelog(ScriptName)   
}

MenuCb_PTNews(ItemName, ItemPos, MenuName){
ScriptName := RegExReplace(A_ScriptName,"\..*","") 
PowerTools_News(ScriptName)   
}

MenuCb_PowerTools_Tweet(ItemName, ItemPos, MenuName){
ScriptName := RegExReplace(A_ScriptName,"\..*","")
PowerTools_TweetMe(ScriptName)    
}

MenuCb_PTCheckForUpdate(ItemName, ItemPos, MenuName){
PowerTools_CheckForUptate()    
}

; -------------------------------------------------------------------------------------------------------------------
PowerTools_TweetPush(ScriptName){

sLogUrl := PowerTools_Changelog(ScriptName,False)
;sToolUrl := PowerTools_Help(ScriptName,False)

If (ScriptName ="NWS")
    ScriptName = NWSPowerTool

sText = New version of #%ScriptName%. See changelog %sLogUrl%

sUrl:= uriEncode(sUrl)
sText := uriEncode(sText)
sTweetUrl = https://twitter.com/intent/tweet?text=%sText%  ;&hashtags=%ScriptName%&url=%sToolUrl%
Run, %sTweetUrl%

} ;eofun

; -------------------------------------------------------------------------------------------------------------------

PowerTools_TweetMe(ScriptName){

sLogUrl := PowerTools_Changelog(ScriptName,False)
;sToolUrl := PowerTools_Help(ScriptName,False)

If (ScriptName ="NWS")
    ScriptName = NWSPowerTool

sText = @tdalon

sText := uriEncode(sText)
sTweetUrl = https://twitter.com/intent/tweet?text=%sText%&hashtags=%ScriptName%
Run, %sTweetUrl%

} ;eofun


; -------------------------------------------------------------------------------------------------------------------
PowerTools_GetParam(Param) {
ParamVal := PowerTools_RegRead(Param)

If !(ParamVal="") 
	return ParamVal
If RegExMatch(Param,"^TeamsFindText(.*)",sMatch) 
    return Teams_GetText(sMatch1,True) ; Default value

Switch Param
{
    Case "TeamsMentionDelay":
        return 1300
    Case "TeamsCommandDelay":
        return 800
    Case "TeamsClickDelay":
        return 500
    Case "TeamsShareDelay":
        return 1500
}
} ;eofun

; -------------------------------------------------------------------------------------------------------------------
PowerTools_SetParam(Param) {
sPrompt = Enter value for %Param%:
Param := StrReplace(Param," ","")
ParamVal := PowerTools_GetParam(Param)
InputBox, Value, %Param%, %sPrompt%, , 200, 150, , , , ,%ParamVal%
If ErrorLevel
    return
PowerTools_RegWrite(Param,ParamVal)
return ParamVal

} ;eofun


PowerTools_ErrDlg(Text,sUrl:=""){

If !sUrl {
    MsgBox 0x10, %A_ScriptName%: Error, %Text% ; OK|Cancel
    Return
}
SetTimer, ChangeButtonNames, 50 

MsgBox 0x11, %A_ScriptName%: Error, %Text% ; OK|Cancel

IfMsgBox, Cancel ; Help
    Run, %sUrl%
return

ChangeButtonNames: 
SetTitleMatchMode, RegEx
IfWinNotExist, .*: Error$
	return  ; Keep waiting.
SetTimer, ChangeButtonNames, Off 
WinActivate 
ControlSetText, Button1, &OK 
ControlSetText, Button2, &Help 
return

}

