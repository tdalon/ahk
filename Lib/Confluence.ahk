#Include <IntelliPaste>
#Include <Clip>
; for CleanUrl

Confluence_IsUrl(sUrl){
If RegExMatch(sUrl,"confluence[\.-]") or RegExMatch(sUrl,"//wiki[\.]") or RegExMatch(sUrl,"\.atlassian\.net/wiki") 
	return True
}
; -------------------------------------------------------------------------------------------------------------------

Confluence_IsWinActive(){
If Not Browser_WinActive()
    return False
sUrl := Browser_GetUrl()
return Confluence_IsUrl(sUrl)
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Confluence_ExpandLinks(sLinks){
; Called by IntelliPaste
	Loop, parse, sLinks, `n, `r
	{
		sLink_i := IntelliPaste_CleanUrl(A_LoopField)	; calls also GetSharepointUrl
		ExpandLink(sLink_i)
	}
}
; -------------------------------------------------------------------------------------------------------------------

Confluence_CleanLink(sUrl){
; link := Confluence_CleanLink(sUrl)
; link[1]: Link url
; link[2]: Link display text
; Paste link with Clip_PasteLink(link[1],link[2])
; Works for page link by title or pageId or tinyUrl


; Extract meta data tag from page source
/* WebRequest := ComObjCreate("Msxml2.XMLHTTP") 
WebRequest.Open("GET", sUrl, false) ; Async=false
WebRequest.Send()
sResponse := WebRequest.ResponseText 
*/

RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
ReRootUrl := StrReplace(sRootUrl,".","\.")

If RegExMatch(sUrl,ReRootUrl . "/dosearchsite\.action\?cql=(.*)",sCQL) { 	
	sCQL := sCQL1
	sCQL := StrReplace(sCQL,"=","%3D")
	sQuote = "
	sCQL := StrReplace(sCQL,sQuote,"%22")
	sCQL := StrReplace(sCQL,"%2B","+")
	sCQL := StrReplace(sCQL,"%28","(")
	sCQL := StrReplace(sCQL,"%29",")")

	; Extract Labels (only AND label= not label in)	
	sPat := "label\+%3D\+%22([^%]*)%22" 
	Pos=1
	While Pos :=    RegExMatch(sCQL, sPat, label,Pos+StrLen(label)) 
		sLabels := sLabels . " #" . label1 
	sLabels := Trim(sLabels) ; remove starting space			
	sDefSearch := sLabels	
	
	; Extract Space Key
	If RegExMatch(sCQL,"space\+?%3D\+?%22([^%]*)%22",sSpace) 
		sSpace := sSpace1

	; Extract Search String
	If RegExMatch(sCQL,"\&queryString%3D(.*)",sSearchString) {
		sDefSearch := sDefSearch . " " . sSearchString1
	}
	If sSpace
		sLinkText := "Confluence Search in Space: " . sSpace
	Else
		sLinkText := "Confluence global Search"
	
	If sLabels
		sLinkText :=  sLinkText . " with Labels: " . sLabels
	
	If sSearchString1
		sLinkText :=  sLinkText . " and Search String: " . sSearchString1
	return [sUrl, sLinkText]
}

sResponse := Confluence_Get(sUrl)
sPat = s)<meta name="ajs-page-title" content="([^"]*)">.*<meta name="ajs-space-name" content="([^"]*)">.*<meta name="ajs-page-id" content="([^"]*)">
RegExMatch(sResponse, sPat, sMatch)
sLinkText := sMatch1 " | " sMatch2 " - Confluence"
sLinkText := StrReplace(sLinkText,"&amp;","&")
RegExMatch(sUrl, "https://[^/]*", sRootUrl)
sUrl := sRootUrl "/pages/viewpage.action?pageId=" sMatch3
return [sUrl, sLinkText]
}

; ----------------------------------------------------------------------
Confluence_CleanUrl(sUrl){
If InStr(sUrl,"/display/") ; pretty link
sResponse := Confluence_Get(sUrl)
sPat = s)<meta name="ajs-page-id" content="([^"]*)">
RegExMatch(sResponse, sPat, sMatch)
RegExMatch(sUrl, "https://[^/]*", sRootUrl)
sUrl := sRootUrl "/pages/viewpage.action?pageId=" sMatch1
return sUrl
} ; eofun

; ----------------------------------------------------------------------
Confluence_Get(sUrl){
; Requires JiraUserName or ConfluenceUserName to be set in the Registry if different from Windows Username
; Syntax: sResponse .= Jira_Get(sUrl)
; Calls: b64Encode

sPassword := Login_GetPassword()
If (sPassword="") ; cancel
    return	

WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("GET", sUrl, false) ; Async=false

 ; https://developer.atlassian.com/cloud/jira/platform/basic-auth-for-rest-apis/
JiraUserName :=  PowerTools_RegRead("ConfluenceUserName")
If !JiraUserName
    JiraUserName :=  PowerTools_RegRead("JiraUserName")
If !JiraUserName
	JiraUserName := A_UserName
sAuth := b64Encode(JiraUserName . ":" . sPassword) ; user:password in base64 
WebRequest.setRequestHeader("Authorization", "Basic " . sAuth) 
WebRequest.setRequestHeader("Content-Type", "application/json")

WebRequest.Send()        
sResponse := WebRequest.responseText
return sResponse

} ; eofun
; ----------------------------------------------------------------------



; -------------------------------------------------------------------------------------------------------------------
; Confluence Expand Link
ExpandLink(sLink){
	sLinkText := Link2Text(sLink)
	MoveBack := StrLen(sLinkText)
	
	SendInput {Raw} %sLinkText%
	SendInput {Shift down}{Left %MoveBack%}{Shift up}
	SendInput, ^k
	sleep, 2000 
	SendInput {Raw} %sLink% 
	SendInput {Enter}
	return
}

; Confluence Search - Search within current Confluence Space
; Called by: NWS.ahk (Win+F Hotkey)
; 
Confluence_Search(sUrl){
static sConfluenceSearch, sSpace

;https://confluenceroot/dosearchsite.action?cql=+space+%3D+%22PMPD%22+and+type+%3D+%22page%22+and+label+in+(%22best_practice%22%2C%22bitbucket%22)&queryString=code
;https://confluenceroot/dosearchsite.action?cql=+space+%3D+%22PMPD%22+and+type+%3D+%22page%22+and+label+in+(%22best_practice%22%2C%22bitbucket%22)&queryString=code
;https://confluenceroot/dosearchsite.action%3Fcql=%2Bspace%2B=%2B%22PMPD%22%2Band%2Btype%2B=%2B%22page%22%2Band%2Blabel%2Bin%2B%28%22best_practice%22%2C%22bitbucket%22%29&queryString=code
;https://confluenceroot/dosearchsite.action?cql=+space+=+"PMPD"+and+type+=+"page"+and+label+in+("best_practice","bitbucket")&queryString=code

; , %2C
; + %2B
; = %3D
; " %22


; extract def search and key from url
RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
ReRootUrl := StrReplace(sRootUrl,".","\.")
sQuote = "
If RegExMatch(sUrl,ReRootUrl . "/dosearchsite\.action\?cql=(.*)",sCQL) { 	
	sCQL := sCQL1
	sCQL := StrReplace(sCQL,"=","%3D")
	sCQL := StrReplace(sCQL,sQuote,"%22")
	sCQL := StrReplace(sCQL,"%2B","+")
	sCQL := StrReplace(sCQL,"%28","(")
	sCQL := StrReplace(sCQL,"%29",")")

	; Extract Labels (only AND label= not label in)	
	sPat := "label\+%3D\+%22([^%]*)%22" 
	Pos=1
	While Pos :=    RegExMatch(sCQL, sPat, label,Pos+StrLen(label)) 
		sLabels := sLabels . " #" . label1 
	sLabels := Trim(sLabels) ; remove starting space			
	sDefSearch := sLabels

	; Extract Space Key
	If RegExMatch(sCQL,"space\+?%3D\+?%22([^%]*)%22",sSpace) 
		sSpace := sSpace1

	; Extract Search String
	If RegExMatch(sCQL,"\&queryString%3D(.*)",sSearchString) {
		sDefSearch := sDefSearch . " " . sSearchString1
	}
	
; Not from advanced search - page view -> Extract Space 
} Else {
	sOldSpace := sSpace
	If RegExMatch(sUrl,ReRootUrl . "/display/([^/]*)",sSpace)
		sSpace := sSpace1
	Else If InStr(sUrl, "/pages/viewpage.action?pageId=") {
		sResponse := Confluence_Get(sUrl)
		sPat = s)<meta name="ajs-space-name" content="([^"]*)">
		RegExMatch(sResponse, sPat, sMatch)
		sSpace := sMatch1
	} Else If  RegExMatch(sUrl,ReRootUrl . "/spaces/viewspace\.action\?key=([^&\?/]*)",sSpace)
		sSpace := sSpace1
	
	If sSpace = %sOldSpace%
		sDefSearch := sConfluenceSearch
	Else
		sDefSearch = 
}

InputBox, sSearch , Confluence Search, Enter search string (use # for labels):,,640,125,,,,,%sDefSearch% 
if ErrorLevel
	return
sSearch := Trim(sSearch) 

; ----
; Convert labels to CQL
sPat := "#([^#\s]*)" 
Pos=1
While Pos :=    RegExMatch(sSearch, sPat, label,Pos+StrLen(label)) {
	sCQLLabels := sCQLLabels . "+and+label+%3D+" . sQuote . label1 . sQuote
} ; end while

; remove labels from search string
sSearch := RegExReplace(sSearch, sPat , "")
sSearch := Trim(sSearch)

sSearchUrl = %sRootUrl%/dosearchsite.action?cql=type+=+"page"
If sSpace
	sSearchUrl := sSearchUrl . "+and+space=%22" . sSpace "%22"

If sCQLLabels ; not empty
	sSearchUrl := sSearchUrl . sCQLLabels 

If sSearch ; not empty
	sSearchUrl = %sSearchUrl%&queryString=%sSearch%
sConfluenceSearch := sDefSearch

If sCQL ; not empty means update search 
	Send ^l
Else
	Send ^n ; New Window
Sleep 500
Clip_Paste(sSearchUrl)
Send {Enter}

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Confluence_PersonalizeMention() {
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2020/11/confluence-personalize-mentions.html"
	return
}
SendInput +{Left}
sLastLetter := Clip_GetSelectionHtml()
IsRealMention := InStr(sLastLetter,"http://schema.skype.com/Mention") 
sLastLetter := Clip_GetSelection()    
SendInput {Right}

If (IsRealMention) {
    If (sLastLetter = ")")
        SendInput {Backspace}^{Left}
    Else 
        SendInput ^{Left}

    SendInput +{Left} 
    sLastLetter := Clip_GetSelection()   
    If (sLastLetter = "-")
        SendInput ^{Left}{Backspace}
    Else 
        SendInput {Backspace}
} Else {
    ; do not personalize if no match to avoid ambiguity which name was mentioned if shorten to firstname and reprompt for matching
    return
    If (sLastLetter = ")") 
        SendInput {Backspace}^{Backspace}{Backspace} ; remove ()

    SendInput ^{Left}^{Backspace}^{Backspace}^{Right}
}
} ; eofun


; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------

; NOT USED

; -------------------------------------------------------------------------------------------------------------------
Confluence_PageTitleLink2Link(sUrl){

RegExMatch(sUrl,"(.*)/display/([^/]*)/([^/]*)",sMatch) 
sRootUrl := sMatch1
spaceKey := sMatch2
pageTitle := sMatch3

restUrl := sRootUrl "/rest/api/content?spaceKey=" spaceKey "&title=" pageTitle

WebRequest := ComObjCreate("Msxml2.XMLHTTP") 
WebRequest.Open("GET", restUrl, false) ; Async=false
WebRequest.Send()


sResponse := WebRequest.ResponseText
;If (WebRequest.Status=200)

sPat = s)"id":"([^"]*)".*"title":"([^"]*)"
RegExMatch(sResponse, sPat, sMatch)

sUrl := sRootUrl "/pages/viewpage.action?pageId=" sMatch1


; Get space name from Space key
restUrl := sRootUrl "/rest/api/space/" spaceKey "/content"
WebRequest.Send()
sResponse := WebRequest.ResponseText


sPat = s)"name":"([^"]*)"
RegExMatch(sResponse, sPat, sMatch)

sLinkText := sMatch2 " (" sMatch1 " Confluence)"
return [sUrl, sLinkText]
}