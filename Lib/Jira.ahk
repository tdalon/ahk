; Jira Lib
; Includes JiraSearch, IsJiraUrl, JiraAuth
; for GetPassword
#Include <Login> 

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
sAuth := b64Encode(A_UserName . ":" . sPassword) ; user:password in base64 
WebRequest.setRequestHeader("Authorization", "Basic " . sAuth) 
WebRequest.setRequestHeader("Content-Type", "application/json")

WebRequest.Send()        
sResponse := WebRequest.responseText
return sResponse
}

; ----------------------------------------------------------------------
Jira_IsUrl(sUrl){
return  (InStr(sUrl,"jira")) ; TODO: edit according your need/ add setting
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
static sJiraSearch, sKey

If RegExMatch(sUrl,sRootUrl . "/display/([^/]*)",sKey)
    sKey := sKey1
Else If  RegExMatch(sUrl,sRootUrl . "/spaces/viewspace.action?key=([^&?/]*)",sKey)
    sKey := sKey1
Else
    Return
sOldKey := sKey
If sKey = %sOldKey%
	sDefSearch := sConfluenceSearch
Else
	sDefSearch = 

InputBox, sSearch , Search string, Enter search string (use # for tags):,,640,125,,,,,%sDefSearch% 
if ErrorLevel
	return
sSearch := Trim(sSearch) 
sConfluenceSearch := StrReplace(sSearch," ","+")
sSearchUrl = /dosearchsite.action?cql=siteSearch+~+"%sConfluenceSearch%"+and+space+=+"%sKey%"+and+type+=+"page""
Run, %sSearchUrl%

}
; ----------------------------------------------------------------------


b64Encode(string)
; ref: https://github.com/jNizM/AHK_Scripts/blob/master/src/encoding_decoding/base64.ahk
{
    VarSetCapacity(bin, StrPut(string, "UTF-8")) && len := StrPut(string, &bin, "UTF-8") - 1 
    if !(DllCall("crypt32\CryptBinaryToString", "ptr", &bin, "uint", len, "uint", 0x1, "ptr", 0, "uint*", size))
        throw Exception("CryptBinaryToString failed", -1)
    VarSetCapacity(buf, size << 1, 0)
    if !(DllCall("crypt32\CryptBinaryToString", "ptr", &bin, "uint", len, "uint", 0x1, "ptr", &buf, "uint*", size))
        throw Exception("CryptBinaryToString failed", -1)
    return StrGet(&buf)
}


; ----------------------------------------------------------------------

Jira_FormatLinks(sLinks,sStyle){
If Not InStr(sLinks,"`n") { ; single line
    sLink := CleanUrl(sLinks)	; calls also GetSharepointUrl
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