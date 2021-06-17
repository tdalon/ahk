#Include <IntelliPaste>
; for CleanUrl

Confluence_IsUrl(sUrl){
If RegExMatch(sUrl,"confluence[\.-]")
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
		sLink_i := CleanUrl(A_LoopField)	; calls also GetSharepointUrl
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
WebRequest := ComObjCreate("Msxml2.XMLHTTP") 
WebRequest.Open("GET", sUrl, false) ; Async=false
WebRequest.Send()
sResponse := WebRequest.ResponseText

sPat = s)<meta name="ajs-page-title" content="([^"]*)">.*<meta name="ajs-space-name" content="([^"]*)">.*<meta name="ajs-page-id" content="([^"]*)">
RegExMatch(sResponse, sPat, sMatch)
sLinkText := sMatch1 " | " sMatch2 " - Confluence"
sLinkText := StrReplace(sLinkText,"&amp;","&")
RegExMatch(sUrl, "https://[^/]*", sRootUrl)
sUrl := sRootUrl "/pages/viewpage.action?pageId=" sMatch3
;MsgBox %sLinkText% %sUrl%
return [sUrl, sLinkText]
}



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
Confluence_Search(sUrl){
static sConfluenceSearch, sKey
; http://confluenceroot/dosearchsite.action?cql=siteSearch+~+"project+status+report"+and+space+=+"projectCFTPT"+and+type+=+"page"

; Extract project key from Url and confluence root url
; http://confluenceroot/spaces/viewspace.action?key=projectCFTI
; or http://confluenceroot/display/projectCFTI/Newsletters

RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
If RegExMatch(sUrl,sRootUrl . "/display/([^/]*)",sKey)
    sKey := sKey1
Else If  RegExMatch(sUrl,sRootUrl . "/spaces/viewspace.action\?key=([^&\?/]*)",sKey)
    sKey := sKey1
Else
    Return
sOldKey := sKey
If sKey = %sOldKey%
	sDefSearch := sConfluenceSearch
Else
	sDefSearch = 

InputBox, sSearch , Confluence Search, Enter search string:,,640,125,,,,,%sDefSearch% 
if ErrorLevel
	return
sSearch := Trim(sSearch) 
sConfluenceSearch := StrReplace(sSearch," ","+")
sSearchUrl = /dosearchsite.action?cql=siteSearch+~+"%sConfluenceSearch%"+and+space+=+"%sKey%"+and+type+=+"page"
Run, %sRootUrl%%sSearchUrl%

}


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