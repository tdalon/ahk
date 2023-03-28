; Jira Lib
; Includes Jira_Search, Jira_IsUrl, Jira_CleanLink
; for GetPassword
#Include <Login> 
#Include <IntelliPaste>
#Include <Clip>

; ----------------------------------------------------------------------
Jira_Get(sUrl,sPassword:=""){
; Syntax: sResponseText := Jira_Get(sUrl,sPassword*)
; Calls: b64Encode

If (sPassword = "") {
	sPassword := Jira_GetPassword()
	If (sPassword="") ; cancel
		return	
}
If !RegExMatch(sUrl,"^http") { ; missing root url
	sUrl:=Jira_GetRootUrl() . sUrl
}
WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("GET", sUrl, false) ; Async=false

; https://developer.atlassian.com/cloud/jira/platform/basic-auth-for-rest-apis/
JiraUserName :=  Jira_GetUserName()
sAuth := b64Encode( JiraUserName . ":" . sPassword) ; user:password in base64 
WebRequest.setRequestHeader("Authorization", "Basic " . sAuth) 
WebRequest.setRequestHeader("Content-Type", "application/json")
WebRequest.Send()        
return WebRequest.responseText
}

; ----------------------------------------------------------------------
Jira_Post(sUrl,sBody:="",sPassword:=""){
; Syntax: sResponseText .= Jira_Get(sUrl,sPassword*)
; Calls: b64Encode
	
If (sPassword = "") {
	sPassword := Jira_GetPassword()
	If (sPassword="") ; cancel
		return	
}
If !RegExMatch(sUrl,"^http") { ; missing root url
	sUrl:=Jira_GetRootUrl() . sUrl
}
WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("POST", sUrl, false) ; Async=false

; https://developer.atlassian.com/cloud/jira/platform/basic-auth-for-rest-apis/
JiraUserName :=  Jira_GetUserName()
sAuth := b64Encode( JiraUserName . ":" . sPassword) ; user:password in base64 
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
; Uses JiraUserName in PowerTools Settings
; Calls: b64Encode
	
If (sPassword = "") {
	sPassword := Jira_GetPassword()
	If (sPassword="") ; cancel
		return	
}
If !RegExMatch(sUrl,"^http") { ; missing root url
	sUrl:=Jira_GetRootUrl() . sUrl
}
WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open(sReqType, sUrl, false) ; Async=false

; https://developer.atlassian.com/cloud/jira/platform/basic-auth-for-rest-apis/
JiraUserName :=  Jira_GetUserName()
sAuth := b64Encode( JiraUserName . ":" . sPassword) ; user:password in base64 
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
return  ( InStr(sUrl,"jira.") or RegExMatch(sUrl,"/(servicedesk|desk)/.*/portal/") ) ; TODO: edit according your need/ add setting
} ;eofun
; ----------------------------------------------------------------------

Jira_IsWinActive(){
If Not Browser_WinActive()
    return False
sUrl := Browser_GetUrl()
return Jira_IsUrl(sUrl)
} ; eofun

; ----------------------------------------------------------------------
Jira_Url2IssueKey(sUrl){
sUrl := RegExReplace(sUrl,"\?.*$","") ; remove optional arguments after ? in url
if RegExMatch(sUrl,"/([A-Z]*\-\d*)$",sMatch) ; url ending with Issue Key
	return sMatch1
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
; Convert labels to CQL
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
	
sUrl := Jira_GetRootUrl() . "/rest/api/2/project/" . sProjectKey
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
	sUrl := Browser_GetActiveUrl()
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
JiraUserName :=  PowerTools_RegRead("JiraUserName")
If !JiraUserName
	JiraUserName := A_UserName
return JiraUserName
}

Jira_GetPassword() {
return Login_GetPassword()
}