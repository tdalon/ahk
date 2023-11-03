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
    sInput := ""
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
    Teamsy_Help(sInput)
    return
Case "bgf","bgo","backgrounds","cbg":
    Teams_BackgroundGetFolder()
    return

Case "bg":
    Teams_MeetingBackgroundSet(sInput)
    return
Case "bgs": ; open background settings via Ctrl+Shift+P depreciated
    Teams_MeetingShortcuts("bg")
    return
Case "bgi": ; import backgrounds
    Teams_BackgroundImport()
    return
Case "lob","lobby":
    Teams_MeetingShortcuts("lobby")
    return
Case "together","tm","to","tg":
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
Case "fi","find":
    sKeyword = find
Case "free","a","av":
    sKeyword = available
Case "s":
    sKeyword = saved
Case "d":
    sKeyword = dnd
Case "ca","cal","calendar":
    WinId := Teams_GetMainWindow()
    WinActivate, ahk_id %WinId%
    SendInput ^4; open calendar
    return
Case "m","me","meet": ; activate meeting window
    WinId := Teams_GetMeetingWindow()
    WinActivate, ahk_id %WinId%
    return
Case "le","leave": ; leave meeting
    Teams_Leave()
    return
Case "raise","hand","ha","rh","ra":  
    Teams_RaiseHand()
    return
Case "rh+":  
    Teams_RaiseHand(1)
    return
Case "rh-","lh":  
    Teams_RaiseHand(0)
    return
Case "li","like":
    If Teams_IsMainWindowActive()
        Teams_ConversationReaction("Like")
    Else
        Teams_MeetingReaction("Like")
    return
Case "ap","clap":
    Teams_MeetingReaction("Applause")
    return
Case "clap2":
    Teams_MeetingReaction("Applause")
    SoundPlay, C:\Users\thierry.dalon\Broadcast\Audio\Fast_Clapping.mp3
    return
Case "la","lol","laugh":
    If Teams_IsMainWindowActive()
        Teams_ConversationReaction("Laugh")
    Else
        Teams_MeetingReaction("Laugh")
    return
Case "he","heart","lo","love":
    If Teams_IsMainWindowActive()
        Teams_ConversationReaction("Heart")
    Else
        Teams_MeetingReaction("Love")
    return
Case "su","surprised","wow":
    If Teams_IsMainWindowActive()
        Teams_ConversationReaction("Surprised")
    Else
        Teams_MeetingReaction("Surprised")
    return
Case "sa","sad":
    If Teams_IsMainWindowActive()
        Teams_ConversationReaction("Sad")
    return
Case "an","angry":
    If Teams_IsMainWindowActive()
        Teams_ConversationReaction("Angry")
    return
Case "fs":  
    Teams_MeetingToggleFullscreen()
    return
Case "sh","share":  
    Teams_MeetingShare()
    return
Case "sh+":  
    Teams_MeetingShare(1)
    return
Case "sh-","unsh":  
    Teams_MeetingShare(0)
    return
Case "rec":
    Teams_MeetingRecord()
    return
Case "rec+":
    Teams_MeetingRecord(1)
    return
Case "rec-":
    Teams_MeetingRecord(1)
    return
Case "fa":  
    FocusAssist(sInput)
    return
Case "fa+":  
    FocusAssist("+")
    return
Case "fa-":  
    FocusAssist("-")
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
    Teams_Quit()
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
Case "vi+":
    Teams_Video(1)
    return
Case "vi-":
    Teams_Video(0)
    return
Case "f","fav":
    Teams_FavsOpen(sInput)
    return
Case "sw": ; switch tenant
    Teams_SwitchTenant(sInput)
    return
Case "of": ; open favorites folder
    Teams_FavsOpenDir()
    return
Case "2fav","2f": ; link 2 favorite
    Teams_Link2Fav(Clipboard)
    return
Case "e2f","p2f": ; email|people 2 favorite
    Teams_Emails2Favs()
    return
Case "f+": ; add to favorite (either link or emails)
    Teams_FavsAdd()
    return
Case "s2t": ; Share To teams
    Teams_ShareToTeams()
    return
Case "2c","oc": ; Selection To Chat, Open Chat
    Teams_OpenChat()
    return
Case "st","status": ; Set Status Message
    Teams_SetStatusMessage()()
    return
Case "t2xl":
    Teams_Members2Excel()
    return
Case "e2m": ; emails 2 mentions
    Teams_Selection2Mentions()
    return
Case "jo","join": ; join current meeting
    Outlook_JoinTeamsMeeting()
    return
Case "jo+": ; join meeting with autoJoin and openChat
    Outlook_JoinTeamsMeeting("",true,true)
    return
Case "sb": ; toggle show/hide sharing control bar
    Teams_SharingControlBar()
    return
Case "sb+": ; show sharing control bar
    Teams_SharingControlBar("+")
    return
Case "sb-": ; hide sharing control bar
    Teams_SharingControlBar("-")
    return
} ; End Switch

Teams_SendCommand(sKeyword,sInput,true)
    
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
    Gui, Submit
    Teamsy_Help(TeamsyEdit)

    GuiTeamsyGuiEscape:
	GuiTeamsyGuiClose:
    
    ButtonCancel:= true
    
    Gui, Cancel
    return
}

Teamsy_Help(sKeyword:="",doOpen:=true){
; sUrl := Teamsy_Help(sKeyword:="",doOpen:=true)
Switch sKeyword 
{
Case "":
    sUrl := "https://tdalon.github.io/ahk/Teamsy"
Case "lang":
    sUrl := "https://tdalon.blogspot.com/2023/10/teams-powertools-lang.html"
Case "getlang":
    sUrl := "https://tdalon.blogspot.com/2023/08/teams-get-client-lang.html"
Case "2c","oc":
    sUrl := "https://tdalon.blogspot.com/2023/01/teams-open-chat.html"
Case "f","fav","f+","of": ; favorites
    sUrl := "https://tdalon.blogspot.com/2023/01/teams-favorites.html"
Case "s2t": ; Share To teams
    sUrl := "https://tdalon.blogspot.com/2023/01/share-to-teams.html"
Case "bgi": ; Import Backgrounds
    sUrl := "https://tdalon.blogspot.com/2023/10/teams-import-backgrounds.html"
Case "cl": ; clear cache
    sUrl := "https://tdalon.blogspot.com/2021/01/teams-clear-cache.html" 
Case "2": ; second instance
    sUrl := "https://tdalon.blogspot.com/2020/12/open-multiple-microsoft-teams-instances.html"
Case "e2t":
    sUrl := "https://tdalon.blogspot.com/2020/03/emails-to-team.html"
Case "bg":
    sUrl := "https://tdalon.blogspot.com/2023/11/teams-set-background.html"
Case "obg","cbg":
    sUrl := "https://tdalon.blogspot.com/2021/01/teams-custom-backgrounds.html"
Case "t2xl": ; Export Team Members to Excel 
    sUrl := "https://tdalon.blogspot.com/2020/08/teams-users2excel.html"
Case "e2m": ; Emails to Mentions
    sUrl:= "https://tdalon.blogspot.com/2020/11/teams-shortcuts-send-mentions.html"
Case "jo": ; Emails to Mentions
    sUrl:= "https://tdalon.blogspot.com/2023/02/teams-quick-join-meeting.html"
Case "sw": ; switch tenant
    sUrl := "https://tdalon.blogspot.com/2023/10/teams-switch-tenant.html"
Case "st": ; set status message
    sUrl := "https://tdalon.blogspot.com/2023/10/teams-set-status-message.html"
Case "sh","sb":
    sUrl:="https://tdalon.blogspot.com/2023/10/teams-quick-share-screen.html"
Case "sh": ; share screen
    sUrl := "https://tdalon.blogspot.com/2023/10/teams-share-screen.html"
Case "ptt":
    sUrl := "https://tdalon.blogspot.com/2023/10/teams-shortcuts-ptt.html"
Case "rh":
    sUrl := "https://tdalon.blogspot.com/2021/02/teams-raise-hand.html"
Case "mu":
    sUrl := "https://tdalon.blogspot.com/2021/04/teams-shortcuts-mute-on-off.html"
Case "rec":
    sUrl :="https://tdalon.blogspot.com/2023/11/teams-shortcuts-record.html"
/* 
Case "ts2xl": ; Export Teams to Excel ; TODO
    sUrl := "https://tdalon.blogspot.com/2020/08/teams-users2excel.html" 
*/
Default:
    sUrl := "https://tdalon.github.io/ahk/Teamsy"
} ; end switch

If (doOpen)
    Run, "%sUrl%"
return sUrl
}