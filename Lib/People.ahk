; Library File for People Social Functions
; Used by PeopleConnector and ConnectionsEnhancer

#Include <Teams>
#Include <Clip>
;#Include <Connections> ; remove Lib declaration because it prompts for ConnectionsRootUrl
; Super global-variables
global PowerTools_ConnectionsRootUrl
global PowerTools_ADCommand
global PowerTools_ADConnection
global PowerTools_ADPath

; ----------------------------------------------------------------------
People_GetSelection(){
If WinActive("ahk_exe Excel.exe") {
    sSelection := Clip_GetSelection()
} Else {
    SelArr := Clip_GetSel2()
    sSelection := SelArr[1]
    ; Limit to body part (because of Connections profile url as source)
    If RegExMatch(sSelection,"s)<body>(.*)</body>",sMatch)
        sSelection := sMatch1
}
If !(sSelection) { ; empty
    sSelection := SelArr[2]
    ;TrayTipAutoHide("People Connector warning!","You need first to select something!")   
    ;return
    If !(sSelection)  ; empty - no selection -> take value from clipboard
        sSelection := clipboard
} 
SelArr[1] := sSelection
return SelArr
} ; eofun
; ----------------------------------------------------------------------
People_GetEmailList(sInput){
; Get EmailList from input string
; Extract Emails from String e.g. copied to clipboard Outlook addresses or Html source
; List is separated with a ;
; Syntax: sEmailList := People_GetEmailList(sInput)

; For call from PeopleConnector
If IsObject(sInput)
    sInput := sInput[1]
sInput := StrReplace(sInput,"%40","@") ; for connext profile links - decode @

sPat = ([0-9a-zA-Z\.\-]+@[0-9a-zA-Z\-\.]*\.[a-z]{2,4})[^a-z\.]{0,1}
; TODO bug if Connections mention
Pos = 1 
While Pos := RegExMatch(sInput,sPat,sMatch,Pos+StrLen(sMatch)){
    ; Teams visit card in html ends with @unq.gbl.spaces @thread.skype
    If InStr(sMatch1,"@thread.sky") or InStr(sMatch1,"@unq.gbl") or InStr(sMatch1,"@thread.tacv") 
        continue 
    If InStr(sEmailList,sMatch1 . ";")
        continue
    sEmailList := sEmailList . sMatch1 . ";"
}

sPat = https?://%PowerTools_ConnectionsRootUrl%/profiles/html/profileView.do\?(userid|key)=[0-9A-Za-z\-]*
sPat := StrReplace(sPat,".","\.")
Pos = 1 
While Pos := RegExMatch(sInput,sPat,sMatch,Pos+StrLen(sMatch)){
    sEmail := Connections_Profile2Email(sMatch)
    If InStr(sEmailList,sEmail . ";")
        continue
    sEmailList := sEmailList . sEmail . ";"
}
return SubStr(sEmailList,1,-1) ; remove ending ;
} ; eof
; ----------------------------------------------------------------------

Email2Uid(sEmail,FieldName:="sAMAccountName"){
    sUid := People_ADGetUserField("mail=" . sEmail, FieldName) ; mailNickname - office uid 
    ;sWinUid := People_ADGetUserField("mail=" . sEmail, "sAMAccountName")  ;- login uid
    ;sOfficeUid := People_ADGetUserField("mail=" . sEmail, "mailNickname")  ;- Office uid
    return sUid
}
; ----------------------------------------------------------------------
People_Emails2Uids(sEmailList,FieldName:="mailNickname"){
;Emails2Uids(sEmailList,FieldName:="mailNickname"|"sAMAccountName")
Loop, parse, sEmailList, ";"
{
    sUid := Email2Uid(A_LoopField,FieldName)
    sUidList = %sUidList%, %sUid%
}	
return SubStr(sUidList,2) ; remove starting ,
}

; ----------------------------------------------------------------------
People_Emails2DUids(sEmailList){
;return list of domain/uids from Email
Loop, parse, sEmailList, ";"
{
    sDN := People_ADGetUserField("mail=" . A_LoopField, "distinguishedName")
    RegExMatch(sDN,",DC=([^,]*)",sDC)
    sUid := Email2Uid(A_LoopField,"sAMAccountName") ; Windows Id
    sUidList = %sUidList%, %sDC1%\%sUid%
}	

return SubStr(sUidList,2) ; remove starting ,
}

; ----------------------------------------------------------------------

winUid2Email(sUid){
sEmail := People_ADGetUserField("sAMAccountName=" . sUid, "mail")
return sEmail
}
; ----------------------------------------------------------------------

winUids2Emails(sUidList){
Loop, parse, sUidList, `;%A_Tab%`,
{
    sEmail := winUid2Email(Trim(A_LoopField))
    sEmailList := sEmailList . ";" . sEmail
}	
return SubStr(sEmailList,2) ; remove starting ;
}

; ----------------------------------------------------------------------
People_oUid2Email(sUid) {
If Instr(sUid,"@") ; no domain in sUid)
    sUid := RegExReplace(sUid,"@.*","")

mail := People_ADGetUserField("mailNickname=" . sUid, "mail") ; mailNickname = office uid 
return mail
}

; ----------------------------------------------------------------------
AD_Init(force:= False){

If (!force And Not (PowerTools_ADPath =""))
    return

sDomain := People_GetDomain()
If (sDomain ="") 
    return "ERROR: Domain not provided"           
    
PowerTools_ADPath := "GC://dc=" . StrReplace(sDomain,".",",dc=") 
; ADODB Connection to AD
PowerTools_ADConnection := ComObjCreate("ADODB.Connection")
PowerTools_ADConnection.Open("Provider=ADsDSOObject")

; Connection
PowerTools_ADCommand := ComObjCreate("ADODB.Command")
PowerTools_ADCommand.ActiveConnection := PowerTools_ADConnection    
} ; eofun
; ----------------------------------------------------------------------

AD_Close(){
If (PowerTools_ADPath ="")
    return  
; Close connection
PowerTools_ADConnection.Close()
ObjRelease(PowerTools_ADCommand)
ObjRelease(PowerTools_ADConnection)
}

; ----------------------------------------------------------------------
People_ADGetUserField(sFilter, sField){
; https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2012-r2-and-2012/cc754232(v=ws.11)
    
AD_Init()

; Search the AD recursively, starting at root of the domain
PowerTools_ADCommand.CommandText := "<" . PowerTools_ADPath . ">" . ";(&(objectCategory=User)(" . sFilter . "));" . sField . ";subtree"

Try {
    objRecordSet := PowerTools_ADCommand.Execute
} Catch e {
    MsgBox 0x10, AD Command Error, AD Command Error.`nCheck that you are connected to your company network or the input field '%sField%'!
    strTxt = ERROR: AD Command failed.
    return strTxt
}

If (objRecordSet.RecordCount == 0){
    strTxt :=  "No Data"  ; no records returned
} Else {
    strTxt := objRecordSet.Fields(sField).Value  ; return value
}

; Cleanup
ObjRelease(objRecordSet)
return strTxt
}

; ----------------------------------------------------------------------
People_GetDomain(){
RegRead, Domain, HKEY_CURRENT_USER\Software\PowerTools, Domain
If (Domain=""){
    Domain := People_SetDomain()
}
return Domain
}
; ----------------------------------------------------------------------
People_SetDomain(){
RegRead, Domain, HKEY_CURRENT_USER\Software\PowerTools, Domain
InputBox, Domain, Domain, Enter your Domain, , 200, 125,,,,, %Domain%
If ErrorLevel
    return
PowerTools_RegWrite("Domain",Domain)
return Domain
} ; eofun

; ----------------------------------------------------------------------
People_Emails2Excel(sSelection){
; Calls: People_GetEmailList, People_Email2Name

If GetKeyState("Ctrl") {
	Run, "https://connectionsroot/blogs/tdalon/entry/people_connector_emails2excel" ; TODO
	return
}

sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") { 
    TrayTipAutoHide("People Connector warning!","No email could be found!")   
    return
}

oExcel := ComObjCreate("Excel.Application") ;handle
oExcel.Workbooks.Add ;add a new workbook
oSheet := oExcel.ActiveSheet
; First Row Header
oSheet.Range("A1").Value := "Name"
oSheet.Range("B1").Value := "LastName"
oSheet.Range("C1").Value := "FirstName"
oSheet.Range("D1").Value := "email"

oExcel.Visible := True ;by default excel sheets are invisible
oExcel.StatusBar := "Copy to Excel..."

Pos = 1 
RowCount = 2

Loop, parse, sEmailList, ";"
{
    Email := A_LoopField
    Name := People_ADGetUserField("mail=" . Email, "DisplayName")
    LastName := People_ADGetUserField("mail=" . Email, "sn")
    FirstName := People_ADGetUserField("mail=" . Email, "GivenName")
    oSheet.Range("A" . RowCount).Value := Name
    oSheet.Range("B" . RowCount).Value := LastName
    oSheet.Range("C" . RowCount).Value := FirstName
    oSheet.Range("D" . RowCount).Value := Email
    RowCount +=1
}	

; expression.Add (SourceType, Source, LinkSource, XlListObjectHasHeaders, Destination, TableStyleName)
oTable := oSheet.ListObjects.Add(1, oSheet.UsedRange,,1)
oTable.Name := "OutlookExport"

oTable.Range.Columns.AutoFit

oExcel.StatusBar := "READY"
}

; ----------------------------------------------------------------------
People_GetName(sSelection) {
; Get Name from selection e.g. Outlook Person or email adress
; sName := GetName(sSelection)

sEmailPat = [^\s@]+@[^\s\.]*\.[a-z]{2,3}

;sPat = [0-9a-zA-Z\.\-]+@[0-9a-zA-Z\-\.]*\.[a-z]{2,3}

; From Outlook field
sPat = U)(.*,.*) \<(%sEmailPat%)\>
If RegExMatch(sSelection,sPat,sMatch){
    sName := SwitchName(sMatch1)
    return sName
} 
; From email
sEmailPat = ([0-9a-zA-Z\.\-]+@[0-9a-zA-Z\-\.]*\.[a-z]{2,3})[^a-z\.]
If RegExMatch(sSelection,sEmailPat,sMatch){
    sName := People_Email2Name(sMatch1)
} Else {
    sName := SwitchName(sSelection)
}
return sName
}


; ----------------------------------------------------------------------
People_Email2Name(Email){

If InStr(Email,"@conti") { ; internal Email
    sName := People_ADGetUserField("mail=" . Email, "DisplayName")
    sName := SwitchName(sName)
    return sName
}
; Office Name Lastname, Firstname
Email := RegexReplace(Email,"@.*","")

If RegExMatch(Email,"([a-z]*)\.(\d*)\.([a-z]*)",sMatch) {
    StringUpper, sMatch1, sMatch1 , T
    StringUpper, sMatch3, sMatch3 , T
    If (StrLen(sMatch2) = 1)
        FirstName := sMatch1 . "0" . sMatch2
    Else
        FirstName := sMatch1 . sMatch2
    LastName := sMatch3
} Else If RegExMatch(Email,"([a-z]*)\.([a-z]*)\.([a-z]*)",sMatch) {
    StringUpper, sMatch1, sMatch1 , T
    StringUpper, sMatch2, sMatch2 , T
    StringUpper, sMatch3, sMatch3 , T
    FirstName := sMatch1 
    LastName := sMatch2 . " " . sMatch3
} Else {
    LastName := RegexReplace(Email,".*\.","")
    FirstName := StrReplace(Email,LastName,"")
    FirstName := SubStr(FirstName,1,-1) ; remove ending .
    StringUpper, FirstName, FirstName , T
    FirstName := RegExReplace(FirstName,"\.(\d)","0$1") ; replace .2 by 02
    ;LastName := StrReplace(Email,FirstName,"")
    StringUpper, LastName, LastName , T
}

sName = %LastName%, %FirstName%

return sName
}

; ----------------------------------------------------------------------

SwitchName(sName){
If InStr(sName,"<!--StartFragment-->") { ; HTML Clipboard
    If RegexMatch(sName,"<span[^>]*>(.*)</span>",sMatch)
        sName := sMatch1
    Else If RegexMatch(sName,"<a[^>]*>([^>]*)</a>",sMatch) ; Connections copy Name from Profile
        sName := sMatch1
}
; keep only first line
If InStr(sName,"`n") {
    sName := SubStr(sName,1,InStr(sName,"`r")-1)
}

If (InStr(sName,",")) {
    LastName := RegexReplace(sName,",.*","")
    FirstName := RegexReplace(sName,".*, ","")
    FirstName := Trim(FirstName) ; for name selection by triple click in teams, remove ending spaces
    FirstName := RegExReplace(FirstName," \(.*\)","") ; Remove (uid) in firstname
    sName = %FirstName% %LastName%
}

return sName
}

; ----------------------------------------------------------------------
People_GetMe(){
; suc := People_GetMe()
MyEmail := People_ADGetUserField("sAMAccountName=" . A_UserName, "mail")
If InStr(MyEmail,"ERROR")
    return False
PowerTools_RegWrite("MyEmail",MyEmail)
MyName := People_ADGetUserField("sAMAccountName=" . A_UserName, "DisplayName")
PowerTools_RegWrite("MyName",MyName)
MyOUid:=People_GetMyOUid()
return True
} ; eofun

; ----------------------------------------------------------------------
People_IsMe(sInput){
; returns true if input is my email, my uid, my displayName or my office Uid
; based on People_ADGetUserField. store values in registery for offline call (e.g. Teams_SmartReply)
 
MyEmail := PowerTools_RegRead("MyEmail")
If (MyEmail="") {
    suc := People_GetMe()
    If (Not suc)
        return
}
If InStr(sInput,"@") {
    MyEmail := People_ADGetUserField("sAMAccountName=" . A_UserName, "mail")
    return (MyEmail = sInput)
}
If (sInput = A_UserName)
    return True

MyName := PowerTools_RegRead("MyName")
If (sInput = MyName)
    return True

MyOUid:=People_GetMyOUid()
If (sInput = MyOUid)
    return True

return False
} ; end of fun   
; ----------------------------------------------------------------------
People_GetMyEmail(){
; MyEmail := People_GetMyEmail()
MyEmail := PowerTools_RegRead("MyEmail")
If (MyEmail="") {
    suc := People_GetMe()
    If (Not suc)
        return
}
MyEmail := PowerTools_RegRead("MyEmail")
return MyEmail
}
; ----------------------------------------------------------------------
People_GetMyOUid(){
; OfficeUid := People_GetMyOUid()
RegRead, OfficeUid, HKEY_CURRENT_USER\Software\PowerTools, OfficeUid
If (OfficeUid="") {
    OfficeUid := People_ADGetUserField("sAMAccountName=" . A_UserName, "mailNickname") ; mailNickname - office uid 
    If InStr(OfficeUid,"ERROR")
        return
    PowerTools_RegWrite("OfficeUid",OfficeUid)
}
return OfficeUid
}

; ----------------------------------------------------------------------
People_ConnectionsOpenProfile(sSelection){
global PowerTools_ConnectionsRootUrl
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") {
    If InStr(sSelection,"<html>") { ; convert html code to plain text
        If RegExMatch(sSelection,"s)<html>(.*)</html>",sMatch) ; remove StartFragment
            sSelection := sMatch1
        sSelection := UnHtml(sSelection)
    }
    sName := SwitchName(sSelection)
    Run, https://%PowerTools_ConnectionsRootUrl%/profiles/html/simpleSearch.do?searchBy=name&searchFor=%sName%
} Else {
    Loop, parse, sEmailList, ";"
    {
         Run,  https://%PowerTools_ConnectionsRootUrl%/profiles/html/profileView.do?email=%A_LoopField%
    }	; End Loop 
}
} ; eofun

; ----------------------------------------------------------------------

People_ConnectionsOpenNetwork(sSelection){
global PowerTools_ConnectionsRootUrl
sEmailList := People_GetEmailList(sSelection)
If (sEmailList = "") {
    ;TODO Warning
} Else {
    Loop, parse, sEmailList, ";"
    {
         sKey:= Connections_Email2Key(A_LoopField)
         Run,  https://%PowerTools_ConnectionsRootUrl%/profiles/html/networkView.do?widgetId=friends&key=%sKey%
    }	; End Loop 
}
} ; eofun

; ----------------------------------------------------------------------
People_PeopleView(sSelection){
If GetKeyState("Ctrl") {
	Run, "https://connectionsroot/blogs/tdalon/entry/people_connector_peopleview" ; TODO
	return
}
CompanyId := PowerTools_RegRead("PeopleViewCompanyId")
sEmailList := People_GetEmailList(sSelection)
Loop, parse, sEmailList, ";"
{
        Uid := People_ADGetUserField("mail=" . A_LoopField, "employeeNumber") 
        Run,  https://performancemanager5.successfactors.eu/sf/orgchart?&company=%CompanyId%&selected_user=%Uid%
}	; End Loop Parse Clipboard
} ; eofun


; ----------------------------------------------------------------------
People_DownloadProfilePicture(sEmail,sFolder){

If GetKeyState("Ctrl") {
	Run, "https://connectionsroot/blogs/tdalon/entry/emails2profilepic" ; TODO
	return
}

; Create .ps1 file
PsFile = %A_Temp%\o365_GetProfilePic.ps1
; Fill the file with commands
If FileExist(PsFile)
    FileDelete, %PsFile%

sUserName := People_ADGetUserField("mail=" . sEmail, "mailNickname")

sText := "$photo=Get-Userphoto -identity %sUserName% -ErrorAction SilentlyContinue`n If($photo.PictureData -ne $null)`n{[io.file]::WriteAllBytes($path,$photo.PictureData)}"

FileAppend, %sText%,%PsFile%

; Run it
RunWait, PowerShell.exe -NoExit -ExecutionPolicy Bypass -Command %PsFile% ;,, Hide
;RunWait, PowerShell.exe -ExecutionPolicy Bypass -Command %PsFile% ,, Hide


} ; eofun



UnHtml(html) {	
   ; original name: ComUnHTML() by 'Guest' from
   ; https://autohotkey.com/board/topic/47356-unhtm-remove-html-formatting-from-a-string-updated/page-2 
   ;html := RegExReplace(html, "\r?\n|\r", "<br>") ; added this because original strips line breaks
   oHTML := ComObjCreate("HtmlFile") 
   oHTML.write(html)
   return % oHTML.documentElement.innerText 
}

