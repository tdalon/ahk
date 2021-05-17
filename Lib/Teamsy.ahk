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
Case "bgf","obg","backgrounds":
    Teams_OpenBackgroundFolder()
    return
Case "bg","bgs","background":
    Teams_MeetingAction("Backgrounds")
    return
Case "together","tm","to":
    Teams_MeetingAction("TogetherMode")
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
    Teams_MeetingReaction("Like")
    return
Case "ap","clap":
    Teams_MeetingReaction("Applause")
    return
Case "la","lol","laugh":
    Teams_MeetingReaction("Laugh")
    return
Case "he","heart":
    Teams_MeetingReaction("Heart")
    return
Case "fs":  
    Teams_MeetingAction("FullScreen")
    return
Case "sh","share":  
    Teams_Share()
    return
Case "sh+":  
    Teams_Share(1)
    return
Case "sh-":  
    Teams_Share(0)
    return
Case "mu","mute":  
    Switch sInput
    {
    Case "a","all","app":
        Teams_MuteApp()
        return
    Case "on":
        Teams_Mute(1)
        return
    Case "off":
        Teams_Mute(0)
        return
    Default:
    }
    Teams_Mute()
    return
Case "mu+":
    Teams_Mute(1)
    return
Case "mu-":
    Teams_Mute(0)
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
Case "v","vi": ; Toggle video 
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
    sleep, 500
    SendInput %sKeyword%
    Delay := PowerTools_GetParam("TeamsCommandDelay")
    Sleep %Delay% 
    SendInput +{enter}
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
If (!doBreak){
    sleep, 800
    SendInput +{enter}
}
    
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