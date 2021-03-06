Teamsy(sInput){
    
If (!sInput) { ; empty
    WinId := Teams_GetMainWindow()
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
Case "-g": ; gui/ launcher    
	sCmd := TeamsyInputBox()
    if ErrorLevel
		return
	sCmd := Trim(sCmd) 

    Teamsy(sCmd)
    return
Case "w": ; Web App
    Switch sInput
    {
    Case "c","cal","ca":
        Teams_OpenWebCal()
        return
    Default:
        Teams_OpenWebApp()
    }
    return
Case "h","-h","help":
    Run, https://tdalon.github.io/ahk/Teamsy
    return
Case "bg","background":
    Teams_OpenBackgroundFolder()
    return
Case "news","-n":
    PowerTools_News(A_ScriptName)
    return
Case "wn":
    sKeyword = whatsnew
Case "u","ur":
    sKeyword = unread
Case "p":
    sKeyword = pop
Case "c":
    sKeyword = call
Case "f","fi":
    sKeyword = find
Case "free","a","av":
    sKeyword = available
Case "sa","save":
    sKeyword = saved
Case "d":
    sKeyword = dnd
Case "ca","cal","calendar":
    WinId := Teams_GetMainWindow()
    WinActivate, ahk_id %WinId%
    SendInput ^4; open calendar
    return
Case "m","me","meet": ; get meeting window
    WinId := Teams_GetMeetingWindow()
    If !WinId ; empty
        return
    WinActivate, ahk_id %WinId%
    ;Teams_NewMeeting()
    return
Case "l","le","leave": ; leave meeting
    WinId := Teams_GetMeetingWindow()
    If !WinId ; empty
        return
    WinActivate, ahk_id %WinId%
    SendInput ^+b ; ctrl+shift+b
    return
Case "raise","hand","ha","rh","ra":  
    Teams_RaiseHand()
    return
Case "li","like":
    Teams_React("Like")
    return
Case "ap","clap":
    Teams_React("Applause")
    return
Case "la","lol":
    Teams_React("Laugh")
    return
Case "he","heart":
    Teams_React("Heart")
    return
Case "sh","share":  
    Teams_Share()
    return
Case "mu","mute":  
    Switch sInput
    {
    Case "a","all","app":
        Teams_MuteApp()
        return
    Default:
    }
    Teams_Mute()
    return
Case "de":  ; decline call
    WinId := Teams_GetMainWindow()
    If !WinId ; empty
        return
    WinActivate, ahk_id %WinId%
    SendInput ^+d ;  ctrl+shift+d 
    return
Case "q","quit": ; quit
    sCmd = taskkill /f /im "Teams.exe"
    Run %sCmd%,,Hide 
    return
Case "re","restart": ; restart
    Teams_Restart()
    return
Case "clean": ; clean restart
    Teams_CleanRestart()
    return
Case "clear","cache","cl": ; clear cache
    Teams_ClearCache()
    return
Case "nm": ; new meeting
    Teams_NewMeeting()
    return
Case "n","new","x","nc": ; new expanded conversation 
    Switch sInput
    {
    Case "m","me","meeting":
        Teams_NewMeeting()
        return
    Default:
    }
    WinId := Teams_GetMainWindow()
    WinActivate, ahk_id %WinId%
    Teams_NewConversation()
    return
Case "v","vi": ; Toggle video with background
    Teams_Video()
    return
Case "fav","of": ; open favorites folder
    Teams_FavsOpenDir()
    return
Case "2fav","2f": ; link 2 favorite
    Teams_Link2Fav(Clipboard)
    return
Case "e2f","p2f": ; email|people 2 favorite
    Teams_Emails2Favs()
    return
} ; End Switch

WinId := Teams_GetMainWindow()
WinActivate, ahk_id %WinId%

Send ^e ; Select Search bar

If (SubStr(sKeyword,1,1) = "@") {
    SendInput @
    sleep, 300
    sInput := SubStr(sKeyword,2)
} Else {
    SendInput /
    sleep, 300
    SendInput %sKeyword%
    sleep, 500
    SendInput {enter}
}

If (!sInput) ; empty
    return
sleep, 500

;sLastChar := SubStr(sInput,StrLen(sInput)) 
doBreak := (SubStr(sInput,StrLen(sInput)) == "-")
If (doBreak) {
    sInput := SubStr(sInput,1,StrLen(sInput)-1) ; remove last -
}

SendInput %sInput%
sleep, 800
If !doBreak
    SendInput {enter}
} ; End function     


TeamsyInputBox(){
    static
    ButtonOK:=ButtonCancel:= false
	Gui GuiTeamsy:New,, Teamsy
    ;Gui, add, Text, ;w600, % Text
    Gui, add, Edit, w190 vTeamsyEdit
    Gui, add, Button, w60 gTeamsyOK Default, &OK
    Gui, add, Button, w60 x+10 gTeamsyHelp, &Help

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
    return TeamsyEdit
    ;----------------------
    TeamsyOK:
    ButtonOK:= true
    return
    ;---------------------- 
    TeamsyHelp:
    Run, "https://tdalon.github.io/ahk/Teamsy"

    GuiTeamsyGuiEscape:
	GuiTeamsyGuiClose:
    
    ButtonCancel:= true
    
    Gui, Cancel
    return
}