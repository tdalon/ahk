Outlook_GetApp() {
; olApp := Outlook_GetApp()
; check if !olApp to abort
try {
    olApp := ComObjActive("Outlook.Application") 
    return olApp  
} catch {
    MsgBox 0x14, Start Outlook,Outlook is not running.`nDo you want to start it?
		IfMsgBox, No
			return
    Run, outlook.exe
    return 
}
} ; eofun
; ------------------------------------------------------------------

Outlook_PersonalizeMentions(sName:=""){
 If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2020/11/teams-shortcuts-personalize-mentions.html"
	return
}
If (sName="") {
    SendInput +{Left}
    sLastLetter := Clip_GetSelection()
    SendInput {Right}
} Else
    sLastLetter := SubStr(sName,0)

If (sLastLetter = ")") {
    SendInput +{Backspace}+{Backspace}+{Backspace}^{Left}^{Backspace}^{Backspace}^{Right}
} Else {
    SendInput ^{Left}^{Backspace}^{Backspace}^{Right}
}
  
; Remove numbers from mention
SendInput +{Left}
sLastLetter := Clip_GetSelection()
If RegExMatch(sLastLetter,"\d")
    SendInput {Delete}
SendInput +{Left}
sLastLetter := Clip_GetSelection()
If RegExMatch(sLastLetter,"\d")
    SendInput {Delete}
SendInput {Right}
} ; eofun 


; ------------------------------------------------------------------

Outlook_GetCurrentItem(olApp:="") {
If (olApp = "")
    olApp := Outlook_GetApp()
    If !olApp
        return
try
    olItem := olApp.ActiveWindow.CurrentItem
catch
    olItem := olApp.ActiveExplorer.Selection.Item(1)

return olItem
} ; eofun

; ------------------------------------------------------------------
Outlook_Recipients2Emails(oItem) {
cnt := oItem.Recipients.Count

Loop, %cnt%
{
    sEmail := oItem.Recipients.Item(A_Index).Address 
    sEmailList := sEmailList . ";" . sEmail
}

sEmailList :=	SubStr(sEmailList,2) ; remove starting ;
return sEmailList
} ; eofun
; ------------------------------------------------------------------



Outlook_Recipient2Email(oRecip) {
;  https://stackoverflow.com/a/51939384/2043349
; takes a Display Name (i.e. "James Smith") and turns it into an email address (james.smith@myco.com)
; necessary because the Outlook address is a long, convoluted string when the email is going to someone in the organization.
; source:  https://stackoverflow.com/questions/31161726/creating-a-check-names-button-in-excel

Address := oRecip.Address
If (SubStr(Address, 1,1) == "/") { ; Inside Organization -> Resolve to SMTP email
    oRecip.Resolve    
    If oRecip.Resolved {
        Switch oRecip.AddressEntry.AddressEntryUserType
        {
        Case 0,5:                                ;olExchangeUserAddressEntry & olExchangeRemoteUserAddressEntry
            oEU := oRecip.AddressEntry.GetExchangeUser
            return oEU.PrimarySmtpAddress
        Case 10, 30:                              ;olOutlookContactAddressEntry & 'olSmtpAddressEntry
            return oRecip.AddressEntry.Address
        Default:
        } ; end switch
    } ; End If
} Else
    return Address
} ; eofun

; ------------------------------------------------------------------
Outlook_Item2Emails(oItem:="",sMode :="",validate:=False) {
; sEmailList := Outlook_Item2Emails(oItem:="",sMode:="",validate:= False)
; If oItem is empty, GetCurrentItem is used by default
; sMode : string containing noFrom, noTo, noCc
; If sMode is empty, user will be prompted for selection noFrom|noTo|noCc
; validate: Default False. if True user will have the possibility to edit the email List in an InputBox

olApp := Outlook_GetApp()
If !olApp
    return
If (oItem == "")
    oItem := Outlook_GetCurrentItem(olApp)

If (sMode == "") { ; Prompt user for selection
    sMode := ListView("To Chat","Select Fields to exclude","noCc|noFrom|noTo","Fields",False)
    If (sMode == "ListViewCancel")
        return
}

; Loop on Recipients
cnt := oItem.Recipients.Count
Loop, %cnt%
{
    oRecip := oItem.Recipients.Item(A_Index)
    ; https://learn.microsoft.com/en-us/office/vba/api/outlook.olmailrecipienttype
    
    If ((oRecip.Type == 1) and !InStr(sMode,"noTo")) 
        or ((oRecip.Type == 2) and !InStr(sMode,"noCc")) 
        or ((oRecip.Type == 0) and !InStr(sMode,"noFrom"))       
     {
        sEmail := Outlook_Recipient2Email(oRecip) 
        sEmailList := sEmailList . ";" . sEmail
    } 
}

RegExReplace(sEmailList, ";" , , nCount)
sEmailList :=	SubStr(sEmailList,2) ; remove starting ;

If (validate) {
    sEMailList:=MultiLineInputBox("Edit Email List (" . nCount . "):", sEmailList, "To Chat")
    if (ErrorLevel)
        return
}

return sEmailList
} ; eofun
; ------------------------------------------------------------------

Outlook_Meeting2Excel(oItem :="") {

If (oItem = "") {
    olApp := ComObjActive("Outlook.Application")
    oItem := Outlook_GetCurrentItem(olApp)
}

oExcel := ComObjCreate("Excel.Application") ;handle
oExcel.Workbooks.Add ;add a new workbook
oSheet := oExcel.ActiveSheet
; First Row Header
oSheet.Range("A1").Value := "Name"
oSheet.Range("B1").Value := "LastName"
oSheet.Range("C1").Value := "FirstName"
oSheet.Range("D1").Value := "email"
oSheet.Range("E1").Value := "Response"
oSheet.Range("F1").Value := "Attendance"

oExcel.Visible := True ;by default excel sheets are invisible
oExcel.StatusBar := "Copy to Excel..."

Pos = 1 
RowCount = 2
cnt := oItem.Recipients.Count
Loop, %cnt%
{
    Email := oItem.Recipients.Item(A_Index).Address 
    Name := oItem.Recipients.Item(A_Index).Name 
    LastName := People_ADGetUserField("mail=" . Email, "sn")
    FirstName := People_ADGetUserField("mail=" . Email, "GivenName")
    oSheet.Range("A" . RowCount).Value := Name
    oSheet.Range("B" . RowCount).Value := LastName
    oSheet.Range("C" . RowCount).Value := FirstName
    oSheet.Range("D" . RowCount).Value := Email

    MeetingResponseStatus := oItem.Recipients.Item(A_Index).MeetingResponseStatus
    
    Switch MeetingResponseStatus
    {
        Case 5:
            MeetingResponseStatus := "None"
        Case 4:
            MeetingResponseStatus := "Declined"
        Case 3:
            MeetingResponseStatus := "Accepted"
        Case 2:
            MeetingResponseStatus := "Tentative"
        Case 1:
            MeetingResponseStatus := "Organizer" ; not used Outlook Bug - Organizer returns 0
        Case 0:
            MeetingResponseStatus := "N/A"
    } 
    oSheet.Range("E" . RowCount).Value := MeetingResponseStatus
    
    Type := oItem.Recipients.Item(A_Index).Type
    ; MsgBox % Type DBG
    Switch Type
    {
        Case 3:
            Type := "Resource"
        Case 2:
            Type := "Optional"
        Case 1:
            Type := "Required"
        Case 0:
            Type := "Organizer" ; not used Outlook Bug. Organizer = Required
    }
    oSheet.Range("F" . RowCount).Value := Type
    RowCount +=1
}

; expression.Add (SourceType, Source, LinkSource, XlListObjectHasHeaders, Destination, TableStyleName)
oTable := oSheet.ListObjects.Add(1, oSheet.UsedRange,,1)
oTable.Name := "OutlookRecipientsExport"

oTable.Range.Columns.AutoFit
oExcel.StatusBar := "READY"
} ; eofun

; ------------------------------------------------------------------
Outlook_CopyLink(oItem:="") {
If (oItem = "") {
    olApp := ComObjActive("Outlook.Application")
    oItem := Outlook_GetCurrentItem(olApp)
}
sEntryID := oItem.EntryID
sLink = outlook:%sEntryID%
sSubject := oItem.Subject
Switch oItem.Class ; sClass
{
Case "43":
    sPost := " (Email from " . oItem.SenderName . ")"
Case "26": ; appointment
    sPost :=  " (Appointment from " . oItem.Organizer . ")"
Default:
    sPost =
} ; end switch

sText = %sSubject%%sPost%
sHtml = <a href="%sLink%">%sText%</a>
Clip_SetHtml(sHtml,sText)
TrayTipAutoHide("Outlook Shortcuts: Copy Link","Outlook link was copied to the clipboard.")
} ; eofun
; ------------------------------------------------------------------


Outlook_Item2Type(oItem){
; Returns Mail, Appointment, Meeting
; https://learn.microsoft.com/en-us/office/vba/api/Outlook.OlObjectClass
sClass := oItem.Class
Switch sClass
{
Case "43": ; olMail
    return "Mail"
Case "26": ; olAppointment
    return "Appointment"
Case "53": ; olMeetingRequest
    return "Meeting"
Default:
    return
} ; end switch
}



GetSenderEmail(oItem) {
; Works also for meetings https://stackoverflow.com/a/52150247/2043349
; 
oPA = oItem.PropertyAccessor
return oPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001E")

} ; eofun

; ------------------------------------------------------------------


Outlook_JoinTeamsMeeting(oItem:="",autoJoin := false, openChat := false) { ; @fun_outlook_jointeamsmeeting@
; Outlook_JoinTeamsMeeting(oItem:="",autoJoin := false, openChat := false)
; If oItem is empty, call Outlook_GetTeamsMeeting: user will be prompted to select Today's Teams meeting to join (Meetings are extracted from Outlook main calendar)
If (oItem == "") 
    oItem := Outlook_GetTeamsMeeting()
    
; Join meeting
If RegExMatch(oItem.Body,"U)<(https://teams\.microsoft\.com/l/meetup-join/.*)>",sMeetingLink)
    sMeetingLink := sMeetingLink1
; for issue with plain Meeting invitation https://techcommunity.microsoft.com/t5/microsoft-teams/meeting-invitation-in-outlook-not-rendering-correctly/m-p/3268856
Else If RegExMatch(oItem.Body,"Um)^https://teams\.microsoft\.com/l/meetup-join/.*$",sMeetingLink) { 
    ; do nothing
}
    
If (sMeetingLink = "") {
    ; MsgBox % oItem.Body
    return
}

; Use microsoft edge because better integrated. Teams Links can be whitelisted (Application Links) to be always opened in Teams Client
Run, msedge.exe "%sMeetingLink%" " --new-window"
WinWaitActive, ahk_exe msedge.exe
NewEdgeWinId := WinExist("ahk_exe msedge.exe") 

TeamsExe := Teams_GetExeName()
WinWaitActive, ahk_exe %TeamsExe%,,3
If ErrorLevel {
    TrayTipAutoHide("Error!","Joined Teams Meeting Window not found!",2000,3)
    return
}

JoinWinId := WinActive("A")

If WinExist("ahk_id " . NewEdgeWinId) {
    WinActivate
    Send ^w ; Close leftover browser window
}


; Click on join button
If (autoJoin) {
    UIA := UIA_Interface()
    TeamsEl := UIA.ElementFromHandle(JoinWinId)
    
    TeamsJoinBtn := TeamsEl.FindFirstBy("AutomationId=prejoin-join-button")

    If !TeamsJoinBtn {
        TrayTip, Join Teams Meeting! , "Join button not found!",,0x2
        return
    }
    TeamsJoinBtn.Click()   
    ; Wait for meeting window to open
    MuteEl:=TeamsEl.WaitElementExist("AutomationId=microphone-button",,,,2000)
    ; Unmute
    Name:=Teams_GetLangName("Unmute","Unmute")
    If (MuteEl.Name = %Name%)
        MuteEl.Click()
    ; Maximize
    WinMaximize, ahk_exe %TeamsExe%

}

; Open Meeting Chat in Web browser
If (openChat) 
    Teams_MeetingOpenChat(sMeetingLink)

; Dismiss Meeting Reminder
RemWinId := WinExist("Reminder(s) ahk_class #32770")  ; Reminder Windows
If RemWinId {
    UIA := UIA_Interface()  
    OlEl := UIA.ElementFromHandle(RemWinId)
    OlEl.FindFirstByName("Meeting. " . oItem.Subject).Click()
    SendInput {tab}{2}
    SendInput d ; Dismiss
}
    

} ; eofun


;-------------------------------------------------------------------------
Outlook_GetTeamsMeeting() {
; oItem := Outlook_GetTeamsMeeting()
; Get Teams Meeting Appointment Item from today meetings in Outlook main calendar
; User is prompted to pick up the meeting via ListBox sorted by Start date with Meeting Title, Start and End Time information
olApp := Outlook_GetApp()
If !olApp {
    msg := "Retry after starting Outlook!" 
    TrayTip, Outlook not running! , %msg%,,0x2
    return
}
   
; DEMO: adjust today time for Demo
;EnvAdd, today, -1 , days

FormatTime, today, %today%, ShortDate

myStart := today . " 0:00am"
myEnd := today . " 11:59pm"


; Construct filter 
strRestriction := "[Start] >= """ . myStart . """ AND [End] < """ . myEnd . """"
;MsgBox % strRestriction

oCalendar := olApp.GetNameSpace("MAPI").GetDefaultFolder(olFolderCalendar := 9)  ; olFolderCalendar = 9
oItems := oCalendar.items()
oItems.IncludeRecurrences := True
oItems.Sort("[Start]")
; Restrict the Items collection for the date range
oItemsInDateRange := oItems.Restrict(strRestriction)

appts := Array()
cnt := 0
for appt in oItemsInDateRange  {
    If !InStr(appt.Body,"https://teams.microsoft.com/l/meetup-join/")  
        Continue
    cnt += 1
    ti := DateParse(appt.Start)
    FormatTime, stTime, %ti%, Time
    
    ti := DateParse(appt.End)
    FormatTime, endTime, %ti%, Time
    
    Title := appt.Subject . " (" . stTime " - " . endTime . ")"
    ApptList .= ( (ApptList<>"") ? "|" : "" ) . Title
    appts.Push(appt)       
} ; end for

If (cnt= 0) {
    msg := "No Teams meeting found in Calendar for " . strRestriction . "!"
    TrayTip, No Online Meeting found! , %msg%,,0x2
    return
}

;-------------------------------------------------------------------------
; Select Meeting closest to current time
EnvAdd, myStartSel, -30 , Minutes
EnvAdd, myEndSel, +30 , Minutes

FormatTime, stTime, %myStartSel%, Time ;h:mm tt
myStartSel := today . " " . stTime
FormatTime, endTime, %myEndSel%, Time
myEndSel := today . " " . endTime

; Construct filter 
strRestriction := "[Start] >= """ . myStartSel . """ AND [Start] < """ . myEndSel . """ AND [End] < """ . myEnd . """"

; Restrict the Items collection for the date range

; https://learn.microsoft.com/en-us/answers/questions/930189/restrict-function-not-working-properly-with-recurr

oItemsInDateRange := oItems.Restrict(strRestriction)


cnt := 0
for apptsel in oItemsInDateRange  {
    If !InStr(appt.Body,"https://teams.microsoft.com/l/meetup-join/")    
        Continue   
    cnt := 1
    break   
} ; end for

sel = 1
If (cnt=1) {
    For index, appt in appts {
        If (appt.Subject == apptsel.Subject) and (appt.Start == apptsel.Start) {
            sel := A_Index
            break
        }    
    }
}
; end find preselection closest to current time
;-------------------------------------------------------------------------


LB := ListBox("Join Meeting","Select Teams Meeting",ApptList,sel,,True)
If (LB ="")
    return
oItem := appts[LB]

return oItem

} ; eofun
;-------------------------------------------------------------------------
