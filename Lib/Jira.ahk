; Jira Lib
; Includes Jira_Search, Jira_IsUrl, Jira_CleanLink
; for GetPassword
#Include <Login> 
#Include <IntelliPaste>
#Include <Clip>

; ----------------------------------------------------------------------
Jira_BasicAuth(sUrl:="",sToken:="") {
; Get Auth String for basic authentification
; sAuth := Jira_BasicAuth(sUrl:="",sToken:="")
; Default Url is Setting JiraRootUrl

; Calls: b64Encode

If !RegExMatch(sUrl,"^http")  ; missing root url or default url
	sUrl := Jira_GetRootUrl() . sUrl

If (sToken = "") {

	; If Cloud
	If InStr(sUrl,".atlassian.net") {
		; Read Jira PowerTools.ini setting
		If !FileExist("PowerTools.ini") {
			PowerTools_ErrDlg("No PowerTools.ini file found!")
			return
		}
			
		IniRead, JiraAuth, PowerTools.ini,Jira,JiraAuth
		If (JiraAuth="ERROR") { ; Section [Jira] Key JiraAuth not found
			PowerTools_ErrDlg("JiraAuth key not found in PowerTools.ini file [Jira] section!")
			return
		}
		
		JsonObj := Jxon_Load(JiraAuth)
		For i, inst in JsonObj 
		{
			url := inst["url"]
			If InStr(sUrl,url) {
				JiraToken := inst["apitoken"]
				If (JiraToken="") { ; Section [Jira] Key JiraAuth not found
					PowerTools_ErrDlg("ApiToken is not defined in PowerTools.ini file [Jira] section, JiraAuth key for url '" . url . "'!")
					return
				}
				JiraUserName := inst["username"]
				break
			}
		}
		If (JiraToken="") { ; Section [Jira] Key JiraAuth not found
			RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
			PowerTools_ErrDlg("No instance defined in PowerTools.ini file [Jira] section, JiraAuth key for url '" . sRootUrl . "'!")
			return
		} 	
		
	} Else { ; server
		JiraToken := Jira_GetPassword()
		If (JiraToken="") ; cancel
			return	
	}
}
If (JiraUserName ="")
	JiraUserName :=  Jira_GetUserName()
If (JiraUserName ="")
	return
sAuth := b64Encode( JiraUserName . ":" . JiraToken)
return sAuth

} ; eofun
; ----------------------------------------------------------------------
Jira_Get(sUrl,sPassword:=""){
; Syntax: sResponseText := Jira_Get(sUrl,sPassword*)
; sPassword Password for server or Api Token for cloud

; Calls: Jira_BasicAuth

If !RegExMatch(sUrl,"^http")  ; missing root url or default url
	sUrl := Jira_GetRootUrl() . sUrl
sAuth := Jira_BasicAuth(sUrl,sPassword)
If (sAuth="") {
	TrayTip, Error, Jira Authentication failed!,,3
	return
}
WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("GET", sUrl, false) ; Async=false

; https://developer.atlassian.com/cloud/jira/platform/basic-auth-for-rest-apis/


sAuth := "Basic " . sAuth
WebRequest.setRequestHeader("Authorization", sAuth) 
WebRequest.setRequestHeader("Content-Type", "application/json")
WebRequest.Send()        
return WebRequest.responseText
}

; ----------------------------------------------------------------------
Jira_Post(sUrl,sBody:="",sPassword:=""){
; Syntax: sResponseText := Jira_Post(sUrl,sBody,sPassword*)
; Calls: Jira_BasicAuth
	
If !RegExMatch(sUrl,"^http")  ; missing root url or default url
	sUrl := Jira_GetRootUrl() . sUrl
sAuth := Jira_BasicAuth(sUrl,sPassword)
If (sAuth="") {
	TrayTip, Error, Jira Authentication failed!,,3
	return
}

WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("POST", sUrl, false) ; Async=false

WebRequest.setRequestHeader("Authorization", "Basic " . sAuth) 
WebRequest.setRequestHeader("Content-Type", "application/json")
If (sBody = "")
	WebRequest.Send() 
Else
	WebRequest.Send(sBody)   
WebRequest.WaitForResponse()
;MsgBox %sUrl%`n%sBody% 
;MsgBox % WebRequest.Status    ; DBG
return WebRequest.responseText
}

; ----------------------------------------------------------------------
Jira_WebRequest(sReqType,sUrl,sBody:="",sPassword:=""){
; Syntax: WebRequest := Jira_WebRequest(sReqType,sUrl,sBody:="",sPassword:="")
; Output WebRequest with fields Status and ResponseText

; Calls: Jira_BasicAuth
	
If !RegExMatch(sUrl,"^http")  ; missing root url or default url
	sUrl := Jira_GetRootUrl() . sUrl
sAuth := Jira_BasicAuth(sUrl,sPassword)
If (sAuth="") {
	TrayTip, Error, Jira Authentication failed!,,3
	return
}

WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open(sReqType, sUrl, false) ; Async=false
WebRequest.setRequestHeader("Authorization", "Basic " . sAuth) 
WebRequest.setRequestHeader("Content-Type", "application/json")
If (sBody = "")
	WebRequest.Send() 
Else
	WebRequest.Send(sBody)   
WebRequest.WaitForResponse()
return WebRequest
}


; ----------------------------------------------------------------------
Jira_IsUrl(sUrl){
return  ( InStr(sUrl,"jira.") or InStr(sUrl,".atlassian.net") or RegExMatch(sUrl,"/(servicedesk|desk)/.*/portal/") ) ; TODO: edit according your need/ add setting
} ;eofun
; ----------------------------------------------------------------------

Jira_IsWinActive(){
If Not Browser_WinActive()
    return False
sUrl := Browser_GetUrl()
return Jira_IsUrl(sUrl)
} ; eofun

; ----------------------------------------------------------------------

Jira_ServerType(sUrl:=""){
; serverType := Jira_ServerType(sUrl:="")
; read deploymentType from serverInfo API response
; serverType = "Server"|"Cloud"
If (sUrl="")  
	sUrl := Jira_GetRootUrl() 
sResponse := Jira_Get(sUrl . "/rest/api/2/serverInfo")
sResponse := Jira_Get(sUrl)
Json := Jxon_Load(sResponse)
serverType := Json["deploymentType"]
return serverType
} ; eofun

; ----------------------------------------------------------------------
Jira_Url2IssueKey(sUrl){
If R4J_IsUrl(sUrl) 
	return R4J_Url2IssueKey(sUrl)
Else {
	sUrl := RegExReplace(sUrl,"\?.*$","") ; remove optional arguments after ? in url
	If RegExMatch(sUrl,"/([A-Z]*\-\d*)$",sMatch) ; url ending with Issue Key
		return sMatch1
}

} ; eofun

; ----------------------------------------------------------------------
Jira_Url2ProjectKey(sUrl){
; sProjectKey := Jira_Url2ProjectKey(sUrl)
; /browse/issuekey
if RegExMatch(sUrl,"/browse/([A-Z]*)\-\d*$",sMatch) ; url ending with Issue Key
	return sMatch1
; /com.easesolutions.jira.plugins.requirements/project?detail=RDMO&issueKey=RDMO-14
if RegExMatch(sUrl,"/project\?detail=([A-Z]*)",sMatch) ; url ending with Issue Key
	return sMatch1
; com.easesolutions.jira.plugins.requirements/coverage?prj=RDMO
; com.easesolutions.jira.plugins.requirements/tracematrix?prj=RDMO
if RegExMatch(sUrl,"\?prj=([A-Z]*)",sMatch) ; url ending with Issue Key
	return sMatch1
} ; eofun

; ----------------------------------------------------------------------
; Jira Search - Search within current Jira Project
; Called by: NWS.ahk Quick Search (Win+F Hotkey)
Jira_Search(sUrl){
static S_JiraSearch, S_ProjectKey	

RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
ReRootUrl := StrReplace(sRootUrl,".","\.")
; issue detailed view
If RegExMatch(sUrl,ReRootUrl . "/browse/([^/]*)",sNewProjectKey) {
    sNewProjectKey := RegExReplace(sNewProjectKey1,"-.*","")
	If (sNewProjectKey = %S_ProjectKey%) and (!S_JiraSearch)
		sDefSearch := S_JiraSearch
	Else {
		S_ProjectKey := sNewProjectKey
		sDefSearch := "project=" . S_ProjectKey . " AND summary ~"
	}
; filter view	
} Else If  RegExMatch(sUrl,ReRootUrl . "/issues/\?jql=(.*)",sJql) { ; https://jira.etelligent.ai/issues/?jql=project%20%3D%20TPI%20AND%20summary%20~%20reuse
	sJql := StrReplace(sJql1,"%20"," ")
	sJql := StrReplace(sJql,"%3D","=")
	sDefSearch := sJql
} 

InputBox, sJql , Search string, Enter Jql string:,,640,125,,,,,%sDefSearch% 
if ErrorLevel
	return
S_JiraSearch := sJql
; Convert labels to Jql
sPat := "#([^#\s]*)" 
Pos=1
While Pos :=    RegExMatch(sJql, sPat, label,Pos+StrLen(label)) {
	sJqlLabels := sJqlLabels . " and labels = " . label1
} ; end while

; remove labels from search string
sJql := RegExReplace(sJql, sPat , "")

sJql := Trim(sJql) 
S_JiraSearch := sJql

; Enclose summary~ description~ with "" if using wildcards ? or * see https://tdalon.blogspot.com/2022/02/jira-partial-text-search.html
sPat1 = [^(?:\s*AND\s*|\s*OR\s*|"\*\(\))]*
;sPat1 = [^"]*
sRep = summary~"$1"
sPat = summary\s*~\s*(%sPat1%\*%sPat1%)
sJql := RegExReplace(sJql,sPat,sRep)

sRep = description~"$1"
sPat = description\s*~\s*(%sPat1%\*%sPat1%)
sJql := RegExReplace(sJql,sPat,sRep)

; Escape Html
sJql := StrReplace(sJql," ","%20")
sJql := StrReplace(sJql,"=","%3D")

sSearchUrl := sRootUrl . "/issues/?jql=" . sJql . sJqlLabels

If sJql ; not empty means update search 
	Send ^l
Else
	Send ^n ; New Window
Sleep 500
Clip_Paste(sSearchUrl)
Send {Enter}

} ; eofun
; ----------------------------------------------------------------------


; ----------------------------------------------------------------------

Jira_FormatLinks(sLinks,sStyle){
If Not InStr(sLinks,"`n") { ; single line
    sLink := IntelliPaste_CleanUrl(sLinks)	; calls also GetSharepointUrl
    sLink := StrReplace(sLink,"%20%"," ")
    linktext := Link2Text(sLink)
    sLink = [%linktext%|%sLink%]  
    return sLink
}

Loop, parse, sLinks, `n, `r
{
    link := Jira_CleanLink(A_LoopField)
    sLink := StrReplace(link[1],"%20%"," ")
    linktext := link[2]
    sLink = [%linktext%|%sLink%]

    If (sStyle = "single-line")
        sFull := sFull . sLink . " "
    Else If (sStyle = "lines-break")
        sFull := sFull . sLink . "`n"
    Else If (sStyle = "bullet-list")
        sFull := sFull . "* " . sLink . "`n"
}
return sFull
} ; eofun
; ----------------------------------------------------------------------


; -------------------------------------------------------------------------------------------------------------------

Jira_CleanLink(sUrl){
; link := Jira_CleanLink(sUrl)
; link[1]: Link url
; link[2]: Link display text
; Paste link with Clip_PasteLink(link[1],link[2])
; Works also for issue link from Confluence with ?src=confmacro

; remove url parameters e.g. ?src=confmacro
sUrl := RegExReplace(sUrl,"\?.*=.*$","")

; Jira Issue link
If RegExMatch(sUrl,"/browse/(?P<IssueKey>.*)$",OutputVar) {
	sRestUrl :=StrReplace(sUrl,"/browse/","/rest/api/latest/issue/")
	sRestUrl := sRestUrl . "?fields=summary"
	sResponse := Jira_Get(sRestUrl)
	sPat = "summary":"(.*?)"
	If RegExMatch(sResponse,sPat,sSummary)
		sLinkText = %OutputVarIssueKey%: %sSummary1%
	Else
		sLinkText = %OutputVarIssueKey%
}
Else If RegExMatch(sUrl,"/browse/(?P<ProjectKey>[A-Z]*)-(?P<IssueNb>\d*)$",OutputVar) {
	sLinkText = %OutputVarProjectKey%-%OutputVarIssueNb%

; Jira ServiceDesk e.g. https://xxx.atlassian.net/servicedesk/customer/portal/15/PPP-632
} Else If RegExMatch(sUrl,"/portal/\d*/(?P<ProjectKey>[A-Z]*)-(?P<IssueNb>\d*)$",OutputVar) {
	
	/* 
	RegExMatch(sUrl, "https://[^/]*", sRootUrl) ; Get RootUrl
	RestUrl = %sRootUrl%/rest/api/2/issue/%OutputVarIssueKey%?fields=summary
	sResponse := Jira_Get(RestUrl)

	sPat = "summary":"(.*?)"
	If RegExMatch(sResponse,sPat,sSummary)
		sLinkText = %OutputVarIssueKey%: %sSummary1%
	Else
		sLinkText = %OutputVarIssueKey% 
	*/
    sIssueKey := OutputVarProjectKey . "-" . OutputVarIssueNb
    ; Convert Ticket Link to Issue link
	MsgBox, 0x24,IntelliPaste: Question, Do you want to convert Ticket link into an Issue link?	
	IfMsgBox Yes 
	{
		RegExMatch(sUrl, "https://[^/]*", sRootUrl) ; Get RootUrl
		sUrl = %sRootUrl%/browse/%sIssueKey%
	}
	sLinkText := sIssueKey	
}

return [sUrl, sLinkText]
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Jira_CreateIssue(Json){


JiraUserName :=  Jira_GetUserName()


} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Jira_ProjectKey2Id(sProjectKey) {
; sProjectId := Jira_ProjectKey2Id(sProjectKey)
	; Get ProjectId from ProjectKey

sUrl := Jira_GetRootUrl() . "/rest/api/latest/project/" . sProjectKey
sResponse := Jira_Get(sUrl)
sPat = "id":"([^"]*)"
RegExMatch(sResponse,sPat,sMatch)
return sMatch1
}


; -------------------------------------------------------------------------------------------------------------------

Jira_InputProjectKey(sTitle:="Project Key?",sPrompt:="Enter Project Key:"){
; sProjectKey := Jira_InputProjectKey()
; Last entered ProjectKey is memorized for next call
; Input can be lower case
; Default Project is got from Active Browser Url
static S_JiraProjectKey
; Get ProjectKey from Url
If Browser_WinActive()
	sUrl := Browser_GetUrl()
sProjectKey := Jira_Url2ProjectKey(sUrl)
If (sProjectKey="")
	sProjectKey := S_JiraProjectKey ; previous value
; Confirm ProjectKey
InputBox, sProjectKey, %sTitle%, %sPrompt%,, 250, 125, , , , ,%sProjectKey%
If ErrorLevel
    return
StringUpper, sProjectKey, sProjectKey
S_JiraProjectKey := sProjectKey
return sProjectKey
} ; eofun
; -------------------------------------------------------------------------------------------------------------------


Jira_GetRootUrl() {
return PowerTools_GetSetting("JiraRootUrl")
}

Jira_GetUserName() {
return PowerTools_GetSetting("JiraUserName")
}

Jira_GetPassword() {
return Login_GetPassword()
}

; -------------------------------------------------------------------------------------------------------------------
Jira_Excel_GetIssueKey() {
oExcel := XL_Handle(1)
;oExcel := ComObjCreate("Excel.Application") 
tblSelected := oExcel.ActiveCell.ListObject
If IsObject(tblSelected)  { ; table
	Found := tblSelected.Range.Rows(1).Find("Key", , -4163,1)
	If IsObject(Found) {
		KeyCell := tblSelected.Range.Cells(oExcel.ActiveCell.Row - tblSelected.HeaderRowRange.Row + 1, Found.Column - tblSelected.HeaderRowRange.Column + 1)
		GoTo, LabelFound
	} Else 
		GoTo, LabelNotFound
} Else { ; No Table
	
	Found := oExcel.ActiveSheet.UsedRange.Find("Key", , -4163,1)  ;After: upper left, LookIn= xlValues=-4163, LookAt: xlWhole=1 https://learn.microsoft.com/en-us/office/vba/api/excel.range.find
	
	If !IsObject(Found)
		GoTo, LabelNotFound
	FoundFirst := Found
	
	LabelGIKFindNext:
	;FoundNext := oExcel.ActiveSheet.UsedRange.FindNext() ; FindNext does not work
	FoundNext := oExcel.ActiveSheet.UsedRange.Find("Key",Found , -4163,1)

	; MsgBox % Found.Address ; DBG
	; MsgBox % FoundNext.Address ; DBG
	If (FoundNext.Address <> FoundFirst.Address) {
		If (FoundNext.Column < oExcel.ActiveCell.Column) {
			Found := FoundNext
			GoTo, LabelGIKFindNext
		}
	}			
	KeyCell := oExcel.ActiveSheet.UsedRange.Cells(oExcel.ActiveCell.Row,Found.Column)
}

LabelFound:
sKey := Trim(KeyCell.Value)	
If !RegExMatch(sKey,"[A-Z]*\-\d*") {
	TrayTip, Error, Wrong format for issue Key= '%sKey%'!,,3
	return
}
return sKey

LabelNotFound:
TrayTip, Error, No Header Cell with 'Key' as Text value found!,,3
return

} ;eofun
; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------
Jira_Excel_GetIssueKeys() {
; Return an array of String for Issue Keys matching to the current Selection
; Issue Keys are searched by matching the column with a "Key" header row
oExcel := XL_Handle(1)
;oExcel := ComObjCreate("Excel.Application") 
tblSelected := oExcel.ActiveCell.ListObject
KeyArray := []

If IsObject(tblSelected)  { ; table
	Found := tblSelected.Range.Rows(1).Find("Key", , -4163,1)
	If !IsObject(Found) {
		TrayTip, Error, No Header Cell with 'Key' as Text value found!,,3
		GoTo, LabelKeyNotFound
	}
		
	For curCell In oExcel.Selection 
	{
		KeyCell := tblSelected.Range.Cells(curCell.Row - tblSelected.HeaderRowRange.Row + 1, Found.Column - tblSelected.HeaderRowRange.Column + 1)
		sKey := Trim(KeyCell.Value)	
		If RegExMatch(sKey,"[A-Z]*\-\d*") {
			KeyArray.Push(sKey)
		}		
	}		
			
} Else { ; No Table
	
	Found := oExcel.ActiveSheet.UsedRange.Find("Key", , -4163,1)  ;After: upper left, LookIn= xlValues=-4163, LookAt: xlWhole=1 https://learn.microsoft.com/en-us/office/vba/api/excel.range.find
	If !IsObject(Found)
		GoTo, LabelKeyNotFound
	FoundFirst := Found

	For curCell In oExcel.Selection  
	{
		Found := oExcel.ActiveSheet.UsedRange.Find("Key", , -4163,1)  ;After: upper left, LookIn= xlValues=-4163, LookAt: xlWhole=1 https://learn.microsoft.com/en-us/office/vba/api/excel.range.find
		FoundFirst := Found
		
		LabelGIKsFindNext:
		;FoundNext := oExcel.ActiveSheet.UsedRange.FindNext() ; FindNext does not work
		FoundNext := oExcel.ActiveSheet.UsedRange.Find("Key",Found , -4163,1)

		If (FoundNext.Address <> FoundFirst.Address) {
			If (FoundNext.Column < curCell.Column) {
				Found := FoundNext
				GoTo, LabelGIKsFindNext
			}
		}			
		KeyCell := oExcel.ActiveSheet.UsedRange.Cells(curCell.Row,Found.Column)
		sKey := Trim(KeyCell.Value)	
		If RegExMatch(sKey,"[A-Z]*\-\d*") {
			KeyArray.Push(sKey)
		}
	}
	
}
return KeyArray

LabelKeyNotFound:

For curCell In oExcel.Selection 
{
	sKey := Trim(curCell.Value)	
	If RegExMatch(sKey,"[A-Z]*\-\d*") {
		KeyArray.Push(sKey)
	}		
}		

;TrayTip, Error, No Header Cell with 'Key' as Text value found!,,3
return KeyArray

} ;eofun
; -------------------------------------------------------------------------------------------------------------------


; -------------------------------------------------------------------------------------------------------------------
Jira_GetIssueKey() {
; IssueKey := Jira_GetIssueKey()
; Get IssueKey from current url or text selection

; From Browser Url
If Browser_WinActive() {
	sUrl := Browser_GetUrl()
	If Jira_IsUrl(sUrl) {
		IssueKey := Jira_Url2IssueKey(sUrl)
		If (IssueKey <> "")
			return IssueKey
	}
}
; From selection
ClipSaved := ClipboardAll

; Try first if selection is manually set (Triple click doesn't work in Outlook #35)
sSelection := Clip_GetSel()

If (sSelection = "")
{
	Click 3 ; Click 2 won't get the word because "-" split it. Select the line
	; Does not work in Outlook
	Send, ^c			;Copy (Ctrl+C)	
	Click ; Remove word selection
	
}

sSelection := Clipboard
Clipboard := ClipSaved ; restore clipboard

If (sSelection = "")
	return

sPat := "([A-Z]{3,})-([\d]{1,})"
RegExMatch(sSelection,sPat,sMatch)
return sMatch 

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Jira_Jql2IssueKeys(sJql,RootUrl:=""){
If (RootUrl="")
	RootUrl:=Jira_GetRootUrl()
sUrl := RootUrl . "/rest/api/2/search?jql=" . sJql
sResponse := Jira_Get(sUrl)
Json := Jxon_Load(sResponse)
;MsgBox % Jxon_Dump(Json)
JsonIssues := Json["issues"]


For i, issue in JsonIssues 
	{
		MsgBox % Jxon_Dump(issue)
		issueJson := Jxon_Load(issue)
		key := issueJson["key"]
	}

} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Jira_GetIssueKeys(sIssueKeys :=""){
; KeyArray := Jira_GetIssueKeys()
; returns an Array which each element contains an IssueKey (string)

KeyArray := []

If !(sIssueKeys = "") {
	sSelection := sIssueKeys
	GoTo GetIssueKeysFromSelection
}

If Browser_WinActive() { 
	sUrl := Browser_GetUrl()
	If Jira_IsUrl(sUrl) {
		sUrl := RegExReplace(sUrl,"&.*$","") ; remove optional arguments after & in url
		; From Search Jql
		; https://jira.etelligent.ai/issues/?jql=Key%20in(%20RDMO-16%2CRDMO-5) Key or issueKey work
		sUrl := StrReplace(sUrl,"%2C",",")
		sUrl := StrReplace(sUrl,"%20"," ")
		
		If RegExMatch(sUrl,"i)/issues/\?jql=(?:issue)?Key=([^&])",sMatch) {
			sIssueKeys := StrReplace(sMatch1,"%2C",",")
			Loop sMatch1,  `,
			{
				KeyArray.Push(A_LoopField)
			}
			return KeyArray
		}
		; From Bulk Edit e.g. R4J
		; https://jira.etelligent.ai/secure/views/bulkedit/BulkEdit1!default.jspa?issueKeys=RDMO-16%2CRDMO-18&reset=true
		
		
		If RegExMatch(sUrl,"i)/bulkedit/.*\?issueKeys=([^&]*)",sMatch) {
			Loop,Parse, sMatch1,`,
			{
				KeyArray.Push(Trim(A_LoopField))
			}
			return KeyArray
		}
		If R4J_IsUrl(sUrl) {
			sIssueKey := R4J_Url2IssueKey(sUrl)
			If !(sIssueKey = "") {
				KeyArray.Push(sIssueKey)
				return KeyArray
			}
				
		} ; end if R4J_IsUrl
					

		sUrl := RegExReplace(sUrl,"\?.*$","") ; remove optional arguments after ? in url
		If RegExMatch(sUrl,"/([A-Z]*\-\d{1,})$",sMatch) { ; url ending with Issue Key
			KeyArray.Push(sMatch1)
			return KeyArray
		}

	} ; end if Jira_IsUrl
	
} Else If WinActive("ahk_exe EXCEL.EXE") {
	KeyArray := Jira_Excel_GetIssueKeys()
	return KeyArray
}

sSelection := Clip_GetSelection()

If (sSelection = "")
	return []

; Loop on issue keys
GetIssueKeysFromSelection:
sKeyPat := "([A-Z]{3,})-([\d]{1,})"
Pos = 1 
While Pos := RegExMatch(sSelection,sKeyPat,sMatch,Pos+StrLen(sMatch)){ 
	sIssueKey := sMatch1 . "-" . sMatch2
	If InStr(sIssueKeyList,sIssueKey . ";")
		continue
	sIssueKeyList := sIssueKeyList . sIssueKey . ";"
	KeyArray.Push(sIssueKey)
}

return KeyArray

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Jira_OpenIssueSelection() {
; Called by Hotkey Ctrl+Shift+I
; Open issues from selection
; Open multiple issues if multiple keys are selected
ClipSaved := ClipboardAll

; Try first if selection is manually set (Triple click doesn't work in Outlook #35)
sSelection := Clip_GetSel()

If (sSelection = "")
{
	Click 3 ; Click 2 won't get the word because "-" split it. Select the line
	; Does not work in Outlook
	Send, ^c			;Copy (Ctrl+C)	
	Click ; Remove word selection
	
}

sSelection := Clipboard
Clipboard := ClipSaved ; restore clipboard

If (sSelection = "")
	return

; Loop on issue keys

sPat := "([A-Z]{3,})-([\d]{1,})"
Pos = 1 
While Pos := RegExMatch(sSelection,sPat,sMatch,Pos+StrLen(sMatch)){ 
	sIssueKey := sMatch1 . "-" . sMatch2
	If InStr(sIssueKeyList,sIssueKey . ";")
		continue
	sIssueKeyList := sIssueKeyList . sIssueKey . ";"
	Jira_OpenIssue(sIssueKey)
}
return SubStr(sIssueKeyList,1,-1) ; remove ending ;
} ; eofun
; ---------------------------------------------------------------------------------


Jira_OpenIssues(IssueArray) {
If IsObject(IssueArray) {
	for index, element in IssueArray 
		{
			 Jira_OpenIssue(element)
		}
} Else {
	Jira_OpenIssue(IssueArray)
}

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Jira_OpenIssue(sIssueKey) {
; Open Issue given its key. Key-Url Mapping is looked in settings defined the INI file 
sIssueUrl := Jira_IssueKey2Url(sIssueKey)
Sleep 1000 ; pause required for BrowserTamer
Run, %sIssueUrl%
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Jira_IssueKey2Url(sIssueKey) {
; IssueUrl := Jira_IssueKey2Url(sIssueKey)
JiraRootUrl := Jira_IssueKey2RootUrl(sIssueKey)
return JiraRootUrl . "/browse/" . sIssueKey
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Jira_IssueKey2RootUrl(sKey) {
; RootUrl := Jira_IssueKey2RootUrl(sKey)
; Key-Url Mapping is looked in settings defined the PowerTools.ini file 
If !FileExist("PowerTools.ini")
	return Jira_GetRootUrl()

If (sKey="")
	return Jira_GetRootUrl()

ProjectKey := RegExReplace(sKey,"\-.*$")
IniRead, JiraUrlMap, PowerTools.ini,Jira,JiraUrlMap
If (JiraUrlMap="ERROR") ; Section [Jira] Key JiraUrlMap not found
	return Jira_GetRootUrl()

JsonObj := Jxon_Load(JiraUrlMap)
JiraRootUrl := JsonObj[ProjectKey]
If !(JiraRootUrl)
	return Jira_GetRootUrl()

JiraRootUrl := RegExReplace(JiraRootUrl,"/$") ; remove ending /
return JiraRootUrl		

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Jira_AddLinkByName(sLinkName, inwardIssueKey, outwardIssueKey) {
; suc := Jira_AddLinkByName(sLinkName, inwardIssueKey, outwardIssueKey)
sRootUrl := Jira_IssueKey2RootUrl(inwardIssueKey)
sUrl := sRootUrl . "/rest/api/2/issueLink"
sBody = {"type":{"name":"%sLinkName%"},"inwardIssue":{"key":"%inwardIssueKey%"},"outwardIssue":{"key":"%outwardIssueKey%"}}
;sResponse:= Jira_Post(sUrl, sBody)
; MsgBox %sUrl%`n%sBody%`n%sResponse%

WebRequest := Jira_WebRequest("POST",sUrl, sBody)
suc := (WebRequest.Status = "201")
return suc
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Jira_AddLink(sLinkName:="",srcIssueKey:="",tgtIssueKey:="") {
; sLog := Jira_AddLink(sLinkName,srcIssueKey,tgtIssueKey)
; srcIssueKey and tgtIssueKey can be a string containing multiple keys or an Array

If (srcIssueKey="") {
	InputBox, srcIssueKey , Input Keys, Input issue Key(s) for link source(s):,, 640, 125
	If ErrorLevel ; user cancelled
		return
}
If (tgtIssueKey="") {
	InputBox, tgtIssueKey , Input Keys, Input issue Key(s) for link target(s):,, 640, 125
	If ErrorLevel ; user cancelled
		return
}

; Input linkName
If (sLinkName = "") {
	InputBox, sLinkName , Input Link, Input (inward or outward) Link name:,, 640, 125
	If ErrorLevel ; user cancelled
		return
}

; Get Link definition
sRootUrl := Jira_IssueKey2RootUrl(inwardIssueKey)
sUrl := sRootUrl . "/rest/api/2/issueLinkType"	
sResponse:= Jira_Get(sUrl)

Json := Jxon_Load(sResponse)
JsonLinks := Json["issueLinkTypes"]

; Shorten LinkName matching
sPat:= StrReplace(sLinkName," ","[^\s]*\s")
sPat := RegExReplace(sPat,"\s$") ; remove trailing \s
sPat := "i)^" . sPat . "[^\s]*" ; case insensitive, start with 
; Loop on Links to find the relevant one and direction


For i, link in JsonLinks 
{
	If RegExMatch(link["inward"],sPat,sMatch) { ;(l["inward"] = sLinkName)
        If !(StrLen(sMatch) = StrLen(link["inward"])) ; full match only
			continue
		inwardIssueKey := tgtIssueKey
		outwardIssueKey := srcIssueKey
		linkName := link["name"]
        Break
    }

	If RegExMatch(link["outward"],sPat,sMatch) {
		If !(StrLen(sMatch) = StrLen(link["outward"])) ; full match only
			continue
		inwardIssueKey := srcIssueKey
		outwardIssueKey := tgtIssueKey
		linkName := link["name"]
        Break
    }
}     

If (linkName ="") {
	TrayTip, Error, No link found matching input name '%sLinkName%'!,,3
	return
}

; Convert to array if necessary

If  (inwardIssueKey.Length() ="") ; no array
	inwardKeys := Jira_GetIssueKeys(inwardIssueKey)
Else 
	inwardKeys := inwardIssueKey

If (outwardIssueKey.Length() ="") ; no array
	outwardKeys := Jira_GetIssueKeys(outwardIssueKey)
Else
	outwardKeys := outwardIssueKey


sPrompt:= "Create '" . linkName . "'' link(s) from:"
; Prompt for confirmation
for index_in, inKey in inwardKeys 
{
	sPrompt := 	sPrompt . inKey . ", "
} ; end for

sPrompt := RegExReplace(sPrompt,", $")
sPrompt:= sPrompt . " to: "
for index_out, outKey in outwardKeys 
{
	sPrompt := sPrompt . outKey . ", "
}
sPrompt := RegExReplace(sPrompt,", $","?")

MsgBox, 4, Create Links?, %sPrompt%
IfMsgBox No ; Exit
	return

; Add Links
for index_in, inKey in inwardKeys 
{
	for index_out, outKey in outwardKeys 
	{
		suc := Jira_AddLinkByName(linkName, inKey, outKey)
		If (suc)
			sLog := sLog . "`n'" . linkName . "' link between " . inKey " and " . outKey . " added."
		Else
			sLog := sLog . "`n'" . linkName . "' link between " . inKey " and " . outKey . " failed to add."
	}
} ; end for

sLog := RegExReplace(sLog,"$\n") ; remove starting \n
return sLog
} ; eofun
; -------------------------------------------------------------------------------------------------------------------


Jira_ViewLinkedIssues(IssueKey:="",sLinkName:="") {
If (IssueKey="") {
	InputBox, IssueKey , Input Key, Input issue Key:,, 300, 125
	If ErrorLevel ; user cancelled
		return
}

; Input linkName
If (sLinkName = "") {
	InputBox, sLinkName , Input Link, Input (inward or outward) Link name(s) (separate by `,):,, 640, 125
	If ErrorLevel ; user cancelled
		return
}

; Get Link definition
sRootUrl := Jira_IssueKey2RootUrl(IssueKey)
sUrl := sRootUrl . "/rest/api/2/issueLinkType"	
sResponse:= Jira_Get(sUrl)
Json := Jxon_Load(sResponse)
JsonLinks := Json["issueLinkTypes"]

Loop, Parse, sLinkName,`,
	{
		; Shorten LinkName matching
		sPat:= StrReplace(A_LoopField," ","[^\s]*\s")
		sPat := RegExReplace(sPat,"\s$") ; remove trailing \s
		sPat := "i)^" . sPat . "[^\s]*" ; case insensitive, start with 
		; Loop on Links to find the relevant one and direction	
		For i, link in JsonLinks 
		{
			li := link["inward"]
			If RegExMatch(li,sPat,sMatch) { 
				If !(StrLen(sMatch) = StrLen(li)) ; full match only
					continue
				inwardIssueKey := tgtIssueKey
				outwardIssueKey := srcIssueKey
				linkName := li
				Break
			}
	
			li := link["outward"]
			If RegExMatch(li,sPat,sMatch) {
				If !(StrLen(sMatch) = StrLen(li)) ; full match only
					continue
				inwardIssueKey := srcIssueKey
				outwardIssueKey := tgtIssueKey
				linkName := li
				Break
			}
		}  
	
		If (linkName ="") {
			TrayTip, Error, No link found matching input name '%A_LoopField%'!,,3
			return
		}
	
		linkNames :=  linkNames . ",'" . linkName . "'"
	
	} ; End Loop


linkNames := RegExReplace(linkNames, "^,","") ; remove first ,

sJql := "issueFunction in linkedIssuesOf('key =" . IssueKey . "'," . linkNames . ")"

sRootUrl := Jira_GetRootUrl()
sUrl := sRootUrl . "/issues/?jql=" . sJql
Run, %sUrl%


} ; eofun
; -------------------------------------------------------------------------------------------------------------------


Jira_GetCfId(JiraRootUrl,CfName){
; CfId := Jira_GetCfId(JiraRootUrl,CfName)
sUrl := JiraRootUrl . "/rest/api/2/field"
sResponse:= Jira_Get(sUrl)
Json := Jxon_Load(sResponse)
For i, field in Json 
{
	If (field["name"] = CfName) {
		return  field["id"] 
	}
}

} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Jira_EditEpic(IssueKey, EpicKey:="", EpicLinkCfId:="", JiraRootUrl := "") {
; suc := Jira_EditEpic( IssueKey, EpicKey, EpicLinkCfId:="")
If (JiraRootUrl="")
	JiraRootUrl := Jira_IssueKey2RootUrl(IssueKey)
If (EpicLinkCfId ="") {
	EpicLinkCfId := Jira_GetCfId(JiraRootUrl,"Epic Link")
}

If (EpicKey ="") {
	EpicKey := Jira_IssueKey2EpicKey(IssueKey,"",JiraRootUrl)
}

; https://docs.atlassian.com/software/jira/docs/api/REST/9.4.2/#api/2/issue-editIssue
sUrl := JiraRootUrl . "/rest/api/2/issue/" . IssueKey
sBody = {"fields": {"%EpicLinkCfId%": "%EpicKey%"}}
;sResponse:= Jira_Post(sUrl, sBody)
; MsgBox %sUrl%`n%sBody%`n%sResponse%

WebRequest := Jira_WebRequest("PUT",sUrl, sBody)
suc := (WebRequest.Status = "204")
return suc
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
	
Jira_IssueKey2EpicKey(IssueKey,EpicNameCfId:="",JiraRootUrl:="") {
; Select Epic belonging to same project as input IssueKey via ListView GUI
If (JiraRootUrl = "")
	JiraRootUrl := Jira_IssueKey2RootUrl(IssueKey)
sJql := "project = " . RegExReplace(IssueKey,"-.*") . " and issuetype = Epic and resolution is EMPTY"
sUrl := JiraRootUrl . "/rest/api/2/search?jql=" . sJql	

sResponse:= Jira_Get(sUrl)
;Run %sUrl%

Json := Jxon_Load(sResponse)
JsonIssues := Json["issues"]

If (EpicNameCfId ="") 
	EpicNameCfId := Jira_GetCfId(JiraRootUrl,"Epic Name")
EpicNameArray := []

EpicArray:= {}

For i, issue in JsonIssues 
{
	EpicNameArray.Push(issue["fields"][EpicNameCfId])
	EpicKeyArray.Push(issue["key"])
	EpicArray[i,1]:=issue["fields"][EpicNameCfId]
	EpicArray[i,2]:=issue["key"]
}  

;EpicIndex := ListView_Select(EpicNameArray,"Select Epic","Epic Name")
EpicIndex := ListView_Select(EpicArray,"Select Epic","Epic Name|Key")
EpicKey := JsonIssues[EpicIndex]["key"]

return EpicKey

} ; eofun
; -------------------------------------------------------------------------------------------------------------------


Jira_AddEpic(srcIssueKey:="", EpicIssueKey:="") {
; sLog := Jira_AddEpic(srcIssueKey)
; srcIssueKey can be a string containing multiple keys or an Array

; Calls: Jira_GetIssueKeys

If (srcIssueKey="") {
	InputBox, srcIssueKey , Input Keys, Input issue Key(s) to link to Epic:,, 640, 125
	If ErrorLevel ; user cancelled
		return
}

; Convert to array if necessary
If  (srcIssueKey.Length() = "") ; no array
	srcIssueKeys := Jira_GetIssueKeys(srcIssueKey)
Else 
	srcIssueKeys := srcIssueKey


sPrompt:= "Link Epic '" . EpicIssueKey . "'' to issues:"
; Prompt for confirmation
for index_in, inKey in srcIssueKeys 
{
	If (EpicIssueKey = "") {
		EpicIssueKey := Jira_IssueKey2EpicKey(inKey)
		If (EpicIssueKey="") {
			; Error Handling
			TrayTip, Error, No Epic selected!,,3
			return
		}
		JiraRootUrl := Jira_IssueKey2RootUrl(inKey)
		EpicLinkCfId := Jira_GetCfId(JiraRootUrl,"Epic Link")
	}
	sPrompt := 	sPrompt . inKey . ", "
} ; end for
sPrompt := RegExReplace(sPrompt,", $")


MsgBox, 4, Link to Epic?, %sPrompt%
IfMsgBox No ; Exit
	return

; Add Links
for index_in, inKey in srcIssueKeys 
{
	suc := Jira_EditEpic(inKey,EpicIssueKey,EpicLinkCfId, JiraRootUrl)
	If (suc)
		sLog := sLog . "`nLink Issue '" . inKey . "' to Epic '" . EpicIssueKey "': ok."
	Else
		sLog := sLog . "`nLink Issue '" . inKey . "' to Epic '" . EpicIssueKey "' -> FAILED!"
	
} ; end for

sLog := RegExReplace(sLog,"$\n") ; remove starting \n
return sLog
} ; eofun
; -------------------------------------------------------------------------------------------------------------------