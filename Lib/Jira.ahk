; Jira Lib
#Include <IntelliPaste>
#Include <Clip>

; ----------------------------------------------------------------------
Jira_IsCloud(JiraRootUrl:=""){
	If (JiraRootUrl ="")
		JiraRootUrl := Jira_GetRootUrl()
	return  InStr(JiraRootUrl,".atlassian.net") 
} ;eofun
; ----------------------------------------------------------------------
Jira_IsUrl(sUrl){
return  (InStr(sUrl,"jira.") or InStr(sUrl,".atlassian.net") or RegExMatch(sUrl,"/(servicedesk|desk)/.*/portal/")) ; TODO: edit according your need/ add setting
} ;eofun
; ----------------------------------------------------------------------
Jira_IsWinActive(){
	If Not Browser_WinActive()
		return False
	sUrl := Browser_GetUrl()
	return Jira_IsUrl(sUrl)
} ; eofun
; ----------------------------------------------------------------------

Jira_BasicAuth(sUrl:="",sToken:="") {
; Get Auth String for basic authentification
; sAuth := Jira_BasicAuth(sUrl:="",sToken:="")
; Default Url is Setting JiraRootUrl

; Calls: b64Encode

If !RegExMatch(sUrl,"^http")  ; missing root url or default url
	sUrl := Jira_GetRootUrl() . "/" . RegExReplace(sUrl,"^/")

If (sToken = "") {
	; If Cloud
	If InStr(sUrl,".atlassian.net") {
		; Read Jira PowerTools.ini setting
		If !FileExist("PowerTools.ini") {
			PowerTools_ErrDlg("No PowerTools.ini file found!")
			return
		}
		App := "Jira"
		IniRead, Auth, PowerTools.ini,%App%,%App%Auth
		If (Auth="ERROR") { ; Section [Jira] Key JiraAuth not found
			PowerTools_ErrDlg(App . "Auth key not found in PowerTools.ini file [" . App . "] section!")
			return
		}

		JsonObj := Jxon_Load(Auth)
		For i, inst in JsonObj 
		{
			url := inst["url"]
			If InStr(sUrl,url) {
				ApiToken := inst["apitoken"]
				If (ApiToken="") { ; Section [Jira] Key JiraAuth not found
					PowerTools_ErrDlg("ApiToken is not defined in PowerTools.ini file [" . App . "] section, " . App . "Auth key for url '" . url . "'!")
					return
				}
				UserName := inst["username"]
				break
			}
		}
		If (ApiToken="") { ; Section [Jira] Key JiraAuth not found
			RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
			PowerTools_ErrDlg("No instance defined in PowerTools.ini file [" . App . "] section, " . App . "Auth key for url '" . sRootUrl . "'!")
			return
		} 		
		
	} Else { ; server
		ApiToken := Jira_GetPassword()
		If (ApiToken="") ; cancel
			return	
	}
}
If (UserName ="")
	UserName := Jira_GetUserName()
If (UserName ="")
	return
sAuth := b64Encode(UserName . ":" . ApiToken)
return sAuth

} ; eofun
; ----------------------------------------------------------------------
Jira_Get(sUrl,sPassword:=""){
; Syntax: sResponseText := Jira_Get(sUrl,sPassword*)
; sPassword Password for server or Api Token for cloud

; Calls: Jira_BasicAuth

sAuth := Jira_BasicAuth(sUrl,sPassword)
If (sAuth="") {
	TrayTip, Error, Jira Authentication failed!,,3
	return
}

If !RegExMatch(sUrl,"^http")  ; missing root url or default url
	sUrl := Jira_GetRootUrl() . "/" . RegExReplace(sUrl,"^/")
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
	sUrl := Jira_GetRootUrl() . "/" . RegExReplace(sUrl,"^/")
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
	sUrl := Jira_GetRootUrl() . "/" . RegExReplace(sUrl,"^/")
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
Jira_Url2IssueKey(sUrl) {
sUrl := RegExReplace(sUrl,"\?.*$","") ; remove optional arguments after ? in url
If RegExMatch(sUrl,"i)/([A-Z\d]*\-\d{1,})$",sMatch) {  ; url ending with Issue Key
	StringUpper,issueKey, sMatch1	
	return issueKey
}
} ; eofun

; ----------------------------------------------------------------------
Jira_Url2ProjectKey(sUrl){
; sProjectKey := Jira_Url2ProjectKey(sUrl)

;/projects/projectkey/
If RegExMatch(sUrl,"/projects/([A-Z\d]*\-\d{1,})/",sMatch) ; url ending with Issue Key
	return sMatch1
; issueKey
issuekey := Jira_Url2IssueKey(sUrl)
If !(issueKey="") {
	StringUpper, issueKey, issueKey
	return RegExReplace(issueKey,"-.*")
}

} ; eofun



; ----------------------------------------------------------------------
; Jira Quick Search - Search in Jira Project
; Called by: Altasy Quick Search 
Jira_QuickSearch(searchString){

jiraRootUrl := Jira_GetRootUrl()
sJql := Query2Jql(searchString)
Jira_OpenJql(sJql,jiraRootUrl)
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
} Else If  RegExMatch(sUrl,ReRootUrl . "/issues/\?jql=(.*)",sJql) { ; <root>/issues/?jql=project%20%3D%20TPI%20AND%20summary%20~%20reuse
	sJql := StrReplace(sJql1,"%20"," ")
	sJql := StrReplace(sJql,"%3D","=")
	sDefSearch := sJql
} 

InputBox, sQuery , Search string, Enter Query string:,,640,125,,,,,%sDefSearch% 
if ErrorLevel
	return
sJql := Query2Jql(sQuery)
S_JiraSearch := sJql
; Convert labels to Jql
sPat := "#([^#\s]*)" 
Pos=1
While Pos :=    RegExMatch(sJql, sPat, label,Pos+StrLen(label)) {
	sJqlLabels := sJqlLabels . " AND labels = " . label1
} ; end while

; remove labels from search string
sJql := RegExReplace(sJql, sPat , "")

sJql := Trim(sJql) 
S_JiraSearch := sJql
sSearchUrl := sRootUrl . "/issues/?jql=" . sJql 
If sJql ; not empty means update search 
	Send ^l
Else
	Send ^n ; New Window
Sleep 500
Clip_Paste(sSearchUrl)
Send {Enter}

} ; eofun
; ----------------------------------------------------------------------
Query2Jql(searchString) {
; Convert labels to Jql
sPat := "#([^#\s]*)" 
Pos=1
While Pos :=    RegExMatch(searchString, sPat, label,Pos+StrLen(label)) {
	sJqlLabels := sJqlLabels . " AND labels = " . label1
} ; end while

; remove labels from search string
sJql := " " . Trim(RegExReplace(searchString, sPat))

; shorten s~ and d~

sJql := StrReplace(sJql," s~"," summary~")
sJql := StrReplace(sJql," d~"," description~")

; Enclose summary~ description~ with "" if using wildcards ? or * see https://tdalon.blogspot.com/2022/02/jira-partial-text-search.html
sPat1 = [^(?:\s*AND\s*|\s*OR\s*|"\*\(\))]*
;sPat1 = [^"]*

sPat = (summary|description)\s?~\s*(%sPat1%\*%sPat1%)
sRep = $1~"$2"
sJql := RegExReplace(sJql,sPat,sRep)


; Escape Html
;sJql := StrReplace(sJql," ","%20")
;sJql := StrReplace(sJql,"=","%3D")

; -ua for unassigned
needle := "\s\-ua"
If RegExMatch(sJql,needle) 
	sJql := RegExReplace(sJql, needle, " AND assignee is EMPTY")

; -w for watched
needle := "\s\-w"
If RegExMatch(sJql,needle) 
	sJql := RegExReplace(sJql, needle, " AND watcher = currentUser()")

; -a for assigned to me
needle := "\s\-a"
If RegExMatch(sJql,needle) 
	sJql := RegExReplace(sJql, needle, " AND assignee = currentUser()")

; -c for creator= me
needle := "\s\-c"
If RegExMatch(sJql,needle) 
	sJql := RegExReplace(sJql, needle, " AND creator = currentUser()")

; -r for reporter= me
needle := "\s\-r"
If RegExMatch(sJql,needle) 
	sJql := RegExReplace(sJql, needle, " AND reporter = currentUser()")

; -u for Unresolved
needle := "\s\-u"
If RegExMatch(sJql,needle) 
	sJql := RegExReplace(sJql, needle, " AND resolution = Unresolved")

; Default Project
If !RegExMatch(sJql,"project\s?=") {
	needle := "\s\-?p\s([^\s]*)"
	If RegExMatch(sJql,needle,sMatch) {
		projectKey := sMatch1
		sJql := RegExReplace(sJql, needle)
	} Else {
		defProjectKey := PowerTools_GetSetting("JiraProject")
		If !(defProjectKey="")
			projectKey := defProjectKey
	}
	If !(projectKey = "")
		sDefJql = project = %projectKey%

	sJql := RegExReplace(sJql,"^\sAND\s")
	sJql := Trim(sJql)
	If !(sDefJql = "") {
		If (sJql ="")
			sJql := sDefJql
		Else
			sJql := sDefJql . " AND " . RegExReplace(sJql,"^\sAND\s")
	}
}
; Default Filter from ini file
JiraDefJql := PowerTools_IniRead("Jira","JiraDefJql")
If !(JiraDefJql="ERROR") {
	If (sJql = "")
		sJql := JiraDefJql
	Else
		sJql := JiraDefJql . " AND " . RegExReplace(sJql,"^\sAND\s")
}

sJql := RegExReplace(sJql,"^\sAND\s")
If (sJql ="")
	sJqlLabels := RegExReplace(sJqlLabels,"^\sAND\s")

return sJql . sJqlLabels
} ; eofun

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
Else If RegExMatch(sUrl,"/browse/(?P<ProjectKey>[A-Z\d]*)-(?P<IssueNb>\d*)$",OutputVar) {
	sLinkText = %OutputVarProjectKey%-%OutputVarIssueNb%

; Jira ServiceDesk e.g. https://xxx.atlassian.net/servicedesk/customer/portal/15/PPP-632
} Else If RegExMatch(sUrl,"/portal/\d*/(?P<ProjectKey>[A-Z\d]*)-(?P<IssueNb>\d*)$",OutputVar) {
	
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
		RegExMatch(sUrl,"https://[^/]*", sRootUrl) ; Get RootUrl
		sUrl = %sRootUrl%/browse/%sIssueKey%
	}
	sLinkText := sIssueKey	
}

return [sUrl, sLinkText]
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Jira_CreateIssue(Json := "", projectKey := ""){
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	PowerTools_OpenDoc("Jira_CreateIssue") 
	return
}
If (Json = "") {
	jiraRootUrl := Jira_GetRootUrl()
	sUrl := jiraRootUrl . "/secure/CreateIssue!default.jspa"

	If !(projectKey = "") {
		pid := Jira_ProjectKey2Id(projectKey,jiraRootUrl)
		sUrl := sUrl . "?pid=" . pid
	}
	Atlasy_OpenUrl(sUrl)
}

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Jira_ProjectKey2Id(projectKey,JiraRootUrl:="") {
; projectId := Jira_ProjectKey2Id(projectKey)
; Get ProjectId from ProjectKey
; ProjectKey will be upper cased
StringUpper, projectKey, projectKey ; required to be upper case
sUrl :=  JiraRootUrl . "/rest/api/latest/project/" . projectKey
sResponse := Jira_Get(sUrl)
sPat = "id":"([^"]*)"
RegExMatch(sResponse,sPat,sMatch)
return sMatch1
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Jira_ProjectId2Key(projectId,JiraRootUrl:="") {
; Get ProjectKey from ProjectId
; projectId := Jira_ProjectKey2Id(projectKey)
	sUrl := JiraRootUrl . "/rest/api/latest/project/" . projectId
	sResponse := Jira_Get(sUrl)
	sPat = "key":"([^"]*)"
	RegExMatch(sResponse,sPat,sMatch)
	return sMatch1
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Jira_ProjectName2Key(prjName,JiraRootUrl:="") {
; Get Project Key from Project Name
; projectKey := Jira_ProjectKey2Id(projectName)
	sUrl := JiraRootUrl . "/rest/api/latest/project/search?query=" . prjName
	sResponse := Jira_Get(sUrl)
	sPat = U)"key":"(.*)","name":"%prjName%"
	RegExMatch(sResponse,sPat,sMatch)
    return sMatch1
} ; eofun
	
; -------------------------------------------------------------------------------------------------------------------
Jira_IssueKey2Id(sIssueKey) {
; sIssueId := Jira_IssueKey2Id(sIssueKey)
; Get IssueId from IssueKey
	sUrl := "/rest/api/latest/issue/" . sIssueKey "?fields=id"
	sResponse := Jira_Get(sUrl)
	sPat = "id":"([^"]*)"
	RegExMatch(sResponse,sPat,sMatch)
	return sMatch1
} ; eofun
; -------------------------------------------------------------------------------------------------------------------


Jira_IssueId2Key(issueId) {
; issueKey := Jira_IssueId2Key(issueId) 
; Get IssueKey from IssueId
	sUrl := "/rest/api/latest/issue/" . issueId "?fields=key"
	sResponse := Jira_Get(sUrl)
	sPat = "key":"([^"]*)"
	RegExMatch(sResponse,sPat,sMatch)
	return sMatch1
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Jira_GetIssueField(issueKey,fieldName) {
; fieldValue := Jira_GetIssueField(issueKey|Id, fieldName) 
; Get field Value of Issue
	sUrl := "/rest/api/latest/issue/" . issueKey "?fields=" . fieldName
	sResponse := Jira_Get(sUrl)
	sPat = "%fieldName%":"([^"]*)"
	RegExMatch(sResponse,sPat,sMatch)
	return sMatch1
} ; eofun

; -------------------------------------------------------------------------------------------------------------------


Jira_InputProjectKey(sTitle:="Project Key?",sPrompt:="Enter Project Key:",DefProjectKey:="") {
; sProjectKey := Jira_InputProjectKey()
; Last entered ProjectKey is memorized for next call
; Input can be lower case
; Default Project is got from Active Browser Url
static S_JiraProjectKey
; Get ProjectKey from Url

If Browser_WinActive() {
	sUrl := Browser_GetUrl()
	If !(sUrl="")
		sProjectKey := Jira_Url2ProjectKey(sUrl)
}
	
If !(sProjectKey="")
	DefProjectKey := sProjectKey 
If (DefProjectKey = "")	
	DefProjectKey := S_JiraProjectKey ; previous value
; Confirm ProjectKey
InputBox, sProjectKey, %sTitle%, %sPrompt%,, 250, 125, , , , ,%DefProjectKey%
If ErrorLevel
    return
StringUpper, sProjectKey, sProjectKey
S_JiraProjectKey := sProjectKey
return sProjectKey
} ; eofun
; -------------------------------------------------------------------------------------------------------------------


Jira_GetRootUrl() { ;@fun_jira_getrooturl@
; Get Jira Root Url

; From Browser Url
If Browser_WinActive() {
	sUrl := Browser_GetUrl()
	If Jira_IsUrl(sUrl) {
		RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
		return sRootUrl
	}
}
return PowerTools_GetSetting("JiraRootUrl")
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Jira_GetUserName() {
	return PowerTools_GetSetting("JiraUserName")
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Jira_GetPassword() {
	static sPassword
	If !(sPassword = "")
		return sPassword
	hWin := WinExist("A")
	InputBox, sPassword, Password, Enter Password for Login, Hide, 200, 125
	WinActivate, ahk_id %hWin% ; required because active window looses focus after the Gui closes
	If ErrorLevel
		return
	return sPassword
} ; eofun

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

return Jira_Selection2IssueKey()

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Jira_Selection2IssueKey() {
; Get IssueKey from Selection

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

sPat := "([A-Z\d]{3,})-([\d]{1,})"
RegExMatch(sSelection,sPat,sMatch)
return sMatch 

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Jira_Jql2IssueKeys(sJql,RootUrl:=""){
If (RootUrl="")
	RootUrl:= Jira_GetRootUrl()
sUrl := RootUrl . "/rest/api/2/search?jql=" . sJql
sResponse := Jira_Get(sUrl)
Json := Jxon_Load(sResponse)
;MsgBox % Jxon_Dump(Json)
JsonIssues := Json["issues"]


For i, issue in JsonIssues 
{
	;MsgBox % Jxon_Dump(issue)
	issueJson := Jxon_Load(issue)
	key := issueJson["key"]
}

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Jira_GetIssueKeys(sIssueKeys :=""){ ; @fun_Jira_GetIssueKeys@
; KeyArray := Jira_GetIssueKeys()
; returns a String Array which each element containing an IssueKey (string)

If !(sIssueKeys = "") {
	sSelection := sIssueKeys
	GoTo GetIssueKeysFromSelection
}

KeyArray := []
If Browser_WinActive() { 
	sUrl := Browser_GetUrl()
	If R4J_IsUrl(sUrl) {
		sIssueKey := R4J_Url2IssueKey(sUrl)
		If !(sIssueKey = "") {
			StringUpper, sIssueKey, sIssueKey
			KeyArray.Push(sIssueKey)
			return KeyArray
		}
	} ; end if R4J_IsUrl
				
	If Jira_IsUrl(sUrl) {
		KeyArray := Jira_Url2IssueKeys(sUrl)
		If !(KeyArray="")
			return KeyArray
	} ; end if Jira_IsUrl
	
} Else If WinActive("ahk_exe EXCEL.EXE") {
	KeyArray := Jira_Excel_GetIssueKeys()
	return KeyArray
}

GetIssueKeysFromSelection:
return Jira_Selection2IssueKeys(sSelection)
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Jira_Url2IssueKeys(sUrl) {
	
	KeyArray := []
			
	; From Search Jql
	; /issues/?jql=Key%20in(%20RDMO-16%2CRDMO-5) Key or issueKey work
	sUrl := StrReplace(sUrl,"%2C",",")
	sUrl := StrReplace(sUrl,"%20"," ")
	
	If RegExMatch(sUrl,"iU)/issues/\?jql=(?:issue)?Key\sin\s\((.*)\)",sMatch) {
		Loop sMatch1,  `,
			KeyArray.Push(A_LoopField)
		return KeyArray
	}
	; From Bulk Edit e.g. R4J
	; server
	; $root/secure/views/bulkedit/BulkEdit1!default.jspa?issueKeys=RDMO-16%2CRDMO-18&reset=true
	If RegExMatch(sUrl,"i)/bulkedit/.*\?issueKeys=([^&]*)",sMatch) {
		Loop,Parse, sMatch1,`,
		{
			KeyArray.Push(Trim(A_LoopField))
		}
		return KeyArray
	}
	; cloud 
	; $root/secure/views/bulkedit/BulkEdit1!default.jspa?reset=true&jql=key%20in%20(RDMO-24%2CRDMO-23%2CRDMO-25)
	If RegExMatch(sUrl,"iU)/bulkedit/.*jql=key\sin\s\((.*)\)",sMatch) {
		Loop,Parse, sMatch1,`,
		{
			;MsgBox % A_LoopField ; DBG
			KeyArray.Push(Trim(A_LoopField))
		}
		return KeyArray
	}
	Key := Jira_Url2IssueKey(sUrl)
	If (Key = "")
		return
	KeyArray.Push(Key)
	return KeyArray
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Jira_Selection2IssueKeys(sSelection :="") {
; Parse selection or input string for issue keys
; return a String Array of Keys
	If (sSelection="")
		sSelection := Clip_GetSelection()
	If (sSelection = "")
		return []
	KeyArray := []
	; Loop on issue keys
	sKeyPat := "i)([A-Z\d]{3,}\-[\d]{1,})"
	Pos = 1 
	While Pos := RegExMatch(sSelection,sKeyPat,sMatch,Pos+StrLen(sMatch))
	{ 
		StringUpper,sIssueKey, sMatch1	
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
; return empty if no issue found; else returns list of issue keys

; Calls: Jira_OpenIssue
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
sPat := "([A-Z\d]{3,})-([\d]{1,})"
Pos = 1 
While Pos := RegExMatch(sSelection,sPat,sMatch,Pos+StrLen(sMatch))
{ 
	sIssueKey := sMatch1 . "-" . sMatch2
	If InStr(sIssueKeyList,sIssueKey . ";")
		continue
	sIssueKeyList := sIssueKeyList . sIssueKey . ";"
	Jira_OpenIssue(sIssueKey)
}
return SubStr(sIssueKeyList,1,-1) ; remove ending ;
} ; eofun
; ---------------------------------------------------------------------------------

Jira_OpenIssues(IssueArray:="") {
; Open Issues one-by-one in issue details view
If (IssueArray ="")
	IssueArray := Jira_GetIssueKeys()
If IsObject(IssueArray) {
	For index, element in IssueArray 
	{
		Jira_OpenIssue(element)
	}
} Else 
	Jira_OpenIssue(IssueArray)
} ; eofun
; ---------------------------------------------------------------------------------

Jira_OpenIssuesNav(IssueArray:="") { ; @fun_Jira_OpenIssuesNav@
; Open Issues in issue navigator
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
    PowerTools_OpenDoc("Jira_OpenIssuesNav") 
    return
}

If (IssueArray ="") {
	IssueArray := Jira_GetIssueKeys()
}

If IsObject(IssueArray) {
	For index, key in IssueArray 
	{
		If (index=1)
			jql := "key in (" . key
		Else
			jql := jql . "," . key
	}
	jql := jql . ")"
} Else 
	jql := "key in (" . key . ")"

Jira_OpenJql(Jql)
} ; eofun
; ---------------------------------------------------------------------------------

Jira_OpenJql(Jql,rootUrl:="") {
If (rootUrl = "")
	rootUrl := Jira_GetRootUrl()
Url := rootUrl . "/issues/?jql=" . Jql
Atlasy_OpenUrl(Url)
} ; eofun
; ---------------------------------------------------------------------------------

Jira_BulkEdit(IssueArray:="") {
; Open Issues in Bulk Edit
; $root/secure/views/bulkedit/BulkEdit1!default.jspa?issueKeys=RDMO-16%2CRDMO-18&reset=true

If GetKeyState("Ctrl") and !GetKeyState("Shift") {
    PowerTools_OpenDoc("Jira_BulkEdit") 
    return
}
If (IssueArray ="")
	IssueArray := Jira_GetIssueKeys()

If IsObject(IssueArray) {
	for index, key in IssueArray 
	{
		If (index=1)
			ikl := key
		Else
			ikl := ikl . "%2C" . key
	}
} Else 
	ikl := key
	
Url := Jira_GetRootUrl() . "/secure/views/bulkedit/BulkEdit1!default.jspa?issueKeys=" . ikl
Atlasy_OpenUrl(Url)
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Jira_OpenIssue(sIssueKey) {
; Open Issue given its key. Key-Url Mapping is looked in settings defined the INI file 
sIssueUrl := Jira_IssueKey2Url(sIssueKey)
Atlasy_OpenUrl(sIssueUrl)
;Run, %sIssueUrl%
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
; srcIssueKey and tgtIssueKey can be a string containing multiple keys or a String Array

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
If  (inwardIssueKey.Length() ="") { ; no array
	inwardKeys := Jira_GetIssueKeys(inwardIssueKey)
} Else 
	inwardKeys := inwardIssueKey

If (outwardIssueKey.Length() ="") { ; no array
	outwardKeys := Jira_GetIssueKeys(outwardIssueKey)
} Else
	outwardKeys := outwardIssueKey


sPrompt:= "Create '" . linkName . "' link(s) from:"
; Prompt for confirmation
sPrompt:= sPrompt . Jira_Keys2String(inwardKeys) . " to: " . Jira_Keys2String(outwardKeys)

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
If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	PowerTools_OpenDoc("Jira_ViewLinkedIssues") 
	return
}
If (IssueKey="") {
	InputBox, IssueKey , Input Key, Input issue Key:,, 300, 125
	If ErrorLevel ; user cancelled
		return
}

; Input linkName
If (sLinkName = "") {
	Prompt := "
	(
		Input (inward or outward) shortened* Link name(s) (separate multiple links by `,)
		Leave empty or Cancel for all links.
	)"
	InputBox, sLinkName , Input Link(s), %Prompt%:,, 640, 150
	If ErrorLevel ; user cancelled
		sLinkName := ""
}

If ! (sLinkName = "") {
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
		}  ; end For JsonLiks

		If (linkName = "") {
			TrayTip, Error, No link found matching input name '%A_LoopField%'!,,3
			return
		}

		linkNames :=  linkNames . ",'" . linkName . "'"

	} ; End For sLinkName


	linkNames := RegExReplace(linkNames, "^,","") ; remove first ,
	
} ; end if sLinkName not empty

If Jira_IsCloud(sRootUrl) {
	If (sLinkName ="")
		sJql := "issue in linkedIssues(" . IssueKey . ")"
	Else
		sJql := "issue in linkedIssues(" . IssueKey . "'," . linkNames . ")"

} Else {
	If (sLinkName ="")
		sJql := "issueFunction in linkedIssuesOf('key =" . IssueKey  . ")"
	Else
		sJql := "issueFunction in linkedIssuesOf('key =" . IssueKey . "'," . linkNames . ")"
}
Jira_OpenJql(sJql)
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

Jira_Keys2String(Keys,sep:=", ") {
	for i, Key in Keys 
	{
		s := 	s . Key . sep
	}	
	s := RegExReplace(s,sep . "$")
	return s
} ; eofun
; -------------------------------------------------------------------------------------------------------------------


Jira_Redirect(sUrl,tgtRootUrl:="") {
If (tgtRootUrl = "") { ; read from Ini file Cloud JiraRootUrls
	If !FileExist("PowerTools.ini") {
		PowerTools_ErrDlg("No PowerTools.ini file found and not tgtRootUrl passed!")
		return
	}
		
	IniRead, JiraRootUrls, PowerTools.ini,Jira,JiraRootUrls
	If (JiraRootUrls="ERROR") { ; Section [Jira] Key JiraRootUrls not found
		PowerTools_ErrDlg("JiraRootUrls key not found in PowerTools.ini file [Jira] section!")
		return
	}

	If InStr(sUrl,".atlassian.net") { ; from Cloud to Server
		Loop, Parse, JiraRootUrls,`,
		{
			If !InStr(A_LoopField,".atlassian.net")
				tgtRootUrl := A_LoopField
		}
		If (tgtRootUrl ="") {
			PowerTools_ErrDlg("JiraRootUrls does not contain a Server url!")
			return
		}
		
	} Else { ; from server to Cloud
		Loop, Parse, JiraRootUrls,`,
		{
			If InStr(A_LoopField,".atlassian.net")
				tgtRootUrl := A_LoopField
		}
		If (tgtRootUrl ="") {
			PowerTools_ErrDlg("JiraRootUrls does not contain a Cloud url!")
			return
		}
	}
} 
tgtRootUrl := RegExReplace(tgtRootUrl,"/$") ; remove trailing sep
RegExMatch(sUrl,"https?://[^/]*",srcRootUrl)
tgtUrl := StrReplace(sUrl,srcRootUrl,tgtRootUrl)

; project links
; $rootc/jira/software/c/projects/DDMO/issues -> $roots/projects/DDMO/issues
If (!InStr(tgtUrl,".atlassian.net") & InStr(sUrl,".atlassian.net")) { ; cloud to server
	tgtUrl := StrReplace(tgtUrl,"/jira/software/c/projects/","/projects/")
	If RegexMatch(sUrl,"/jira/dashboards/(\d*)",sMatch)
		tgtUrl := tgtRootUrl . "/secure/Dashboard.jspa?selectPageId=" . sMatch1
; dashboard links
; Example of dashboard  $rootc/jira/dashboards/19902 original $roots/secure/Dashboard.jspa?selectPageId=19902)	
} Else If (InStr(tgtUrl,".atlassian.net") & !InStr(sUrl,".atlassian.net")) { ; server to cloud
	tgtUrl := StrReplace(tgtUrl,"/projects/","/jira/software/c/projects/")
	If RegexMatch(sUrl,"Dashboard\.jspa\?selectPageId=(\d*)",sMatch)
		tgtUrl := tgtRootUrl . "/jira/dashboards/" . sMatch1
}
; board links
If (!InStr(tgtUrl,".atlassian.net") & InStr(sUrl,".atlassian.net")) { ; cloud to server
	If RegexMatch(sUrl,"/jira/.*/boards/(\d*)",sMatch)
		tgtUrl := tgtRootUrl . "/secure/RapidBoard.jspa?rapidView=" . sMatch1
} Else If (InStr(tgtUrl,".atlassian.net") & !InStr(sUrl,".atlassian.net")) { ; server to cloud
	If RegexMatch(sUrl,"/secure/RapidBoard\.jspa\?rapidView=(\d*)",sMatch) {
		boardId := sMatch1
	; get board name
	restUrl := tgtRootUrl . "/rest/agile/1.0/board/" . boardId
	sResponse := Jira_Get(restUrl)
	sPat = "name": "([^"]*)"
	If RegExMatch(sResponse,sPat,sMatch)
		tgtUrl := tgtRootUrl . "/jira/boards?contains=" . sMatch1
	}
		
}

return tgtUrl

} ; eofun
; -------------------------------------------------------------------------------------------------------------------