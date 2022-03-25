; Jira Lib
; Includes Jira_Search, Jira_IsUrl, Jira_CleanLink
; for GetPassword
#Include <Login> 
#Include <IntelliPaste>
#Include <Clip>

; ----------------------------------------------------------------------
Jira_Get(sUrl){
; Syntax: sResponse .= Jira_Get(sUrl)
; Calls: b64Encode

sPassword := Login_GetPassword()
If (sPassword="") ; cancel
    return	

WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("GET", sUrl, false) ; Async=false

 ; https://developer.atlassian.com/cloud/jira/platform/basic-auth-for-rest-apis/
JiraUserName :=  PowerTools_RegRead("JiraUserName")
If !JiraUserName
	JiraUserName := A_UserName
sAuth := b64Encode( JiraUserName . ":" . sPassword) ; user:password in base64 
WebRequest.setRequestHeader("Authorization", "Basic " . sAuth) 
WebRequest.setRequestHeader("Content-Type", "application/json")

WebRequest.Send()        
sResponse := WebRequest.responseText
return sResponse
}

; ----------------------------------------------------------------------
Jira_IsUrl(sUrl){
return  (InStr(sUrl,"jira.")) ; TODO: edit according your need/ add setting
} ;eofun
; ----------------------------------------------------------------------

Jira_IsWinActive(){
If Not Browser_WinActive()
    return False
sUrl := Browser_GetUrl()
return Jira_IsUrl(sUrl)
} ; eofun

; ----------------------------------------------------------------------
; Jira Search - Search within current Jira Project TODO
; Called by: NWS.ahk Quick Search (Win+F Hotkey)
Jira_Search(sUrl){
static sJiraSearch, sProjectKey

RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
ReRootUrl := StrReplace(sRootUrl,".","\.")
; issue detailed view
If RegExMatch(sUrl,ReRootUrl . "/browse/([^/]*)",sNewProjectKey) {
    sNewProjectKey := RegExReplace(sNewProjectKey1,"-.*","")
	If sNewProjectKey = %sProjectKey%
		sDefSearch := sJiraSearch
	Else {
		sProjectKey := sNewProjectKey
		sDefSearch := "project=" . sProjectKey . " AND summary ~"
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
sJql := Trim(sJql) 
sJiraSearch := sJql
sJql := StrReplace(sJql," ","%20")
sJql := StrReplace(sJql,"=","%3D")
sSearchUrl = %sRootUrl%/issues/?jql=%sJql%

; TODO
; Enclose ~summary ~description "" if using wildcards ? or *

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
sUrl := RegExReplace(sUrl,"\?.*","")

RegExMatch(sUrl, "https://[^/]*", sRootUrl)


; Jira Issue link
If RegExMatch(sUrl,"/browse/(?P<IssueKey>.*)$",OutputVar) {
	sUrl :=StrReplace(sUrl,"/browse/","/rest/api/latest/issue/")
	sUrl := sUrl . "?fields=summary"
	sResponse := Jira_Get(sUrl)
	sPat = "summary":"(.*?)"
	If RegExMatch(sResponse,sPat,sSummary)
		sLinkText = %OutputVarIssueKey%: %sSummary1%
	Else
		sLinkText = %OutputVarIssueKey%
}
Else If RegExMatch(sUrl,"/browse/(?P<ProjectKey>[A-Z]*)-(?P<IssueNb>\d*)$",OutputVar) {
	sLinkText = %OutputVarProjectKey%-%OutputVarIssueNb%
; Jira ServiceDesk
} Else If RegExMatch(sUrl,"/desk/portal/1/(?P<IssueKey>.*)$",OutputVar) {
	RestUrl = https://%sRootUrl%/rest/api/2/issue/%OutputVarIssueKey%?fields=summary
	sResponse := Jira_Get(RestUrl)
	sPat = "summary":"(.*?)"
	If RegExMatch(sResponse,sPat,sSummary)
		sLinkText = %OutputVarIssueKey%: %sSummary1%
	Else
		sLinkText = %OutputVarIssueKey%
    
    ; Convert Ticket Link to Issue link
	MsgBox, 0x24,IntelliPaste: Question, Do you want to convert Ticket link into an Issue link?	
	IfMsgBox Yes 
		sUrl = https://%JiraRoot%/browse/%OutputVarIssueKey%
}

return [sUrl, sLinkText]
}