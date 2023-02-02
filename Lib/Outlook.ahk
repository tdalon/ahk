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
    olApp := ComObjActive("Outlook.Application")
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
; If oItem is empty, GetCurrentItem
; If sMode is empty, user will be prompted for selection noFrom|noTo|noCc
; sMode string containing noFrom, noTo, noCc
; validate: if True user will have the possibility to edit the email List in an InputBox

olApp := ComObjActive("Outlook.Application")
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

If (oItem ="") {
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
If (oItem ="") {
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