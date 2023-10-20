#Include <IntelliPaste>
#Include <Clip>
; for CleanUrl



; ----------------------------------------------------------------------
Confluence_BasicAuth(sUrl:="",sToken:="") {
	; Get Auth String for basic authentification
	; sAuth := Confluence_BasicAuth(sUrl:="",sToken:="")
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
	
		
	
	If (JiraUserName ="") {
		JiraUserName :=  PowerTools_RegRead("ConfluenceUserName")
		If !JiraUserName
			JiraUserName :=  JiraUserName :=  Jira_GetUserName()
	}
		
	If (JiraUserName ="")
		return
	sAuth := b64Encode( JiraUserName . ":" . JiraToken)
	return sAuth
	
	} ; eofun


Confluence_IsUrl(sUrl){
If RegExMatch(sUrl,"confluence[\.-]") or RegExMatch(sUrl,"\.atlassian\.net/wiki") or RegExMatch(sUrl,"^https://wiki\.")	
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
	sPat = label\+=\+"([^"]*)" 
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
If !RegExMatch(sResponse, sPat, sMatch) {
	TrayTipAutoHide("Confluence: CleanLink!","Getting PageTitle, SpaceName and PageId by HttpGet failed!")  	
	return
}
sLinkText := sMatch1 " - " sMatch2  ; | will break link in Jira RTF Field
sLinkText := StrReplace(sLinkText,"&amp;","&")
; extract section
If RegExMatch(sUrl,"#([^?&]*)",sSection) {
	sSection1 := RegExReplace(sSection1,".*-","")
	sLinkText := sLinkText . ": " . sSection1
}
sLinkText := sLinkText . " - Confluence"
sUrl := sRootUrl . "/pages/viewpage.action?pageId=" . sMatch3 . sSection
return [sUrl, sLinkText]
}

; ----------------------------------------------------------------------
Confluence_CleanUrl(sUrl){
; Convert Confluence url to link to page by name to link to page by pageId
If InStr(sUrl,"/display/") { ; pretty link
	sResponse := Confluence_Get(sUrl)
	sPat = s)<meta name="ajs-page-id" content="([^"]*)">
	If !RegExMatch(sResponse, sPat, sMatch) { ; error on Confluence_Get
	 	TrayTipAutoHide("Confluence: CleanUrl!","Getting PageId by API failed!")   
		return sUrl
	}
	RegExMatch(sUrl, "https://[^/]*", sRootUrl)

	; extract section
	sNewUrl := sRootUrl . "/pages/viewpage.action?pageId=" . sMatch1
	If RegExMatch(sUrl,"#([^?&]*)",sSection) 
		sNewUrl := sNewUrl . sSection
	return sNewUrl
}
return sUrl
} ; eofun

; ----------------------------------------------------------------------
Confluence_Get(sUrl){
; Requires JiraUserName or ConfluenceUserName to be set in the Registry if different from Windows Username
; Syntax: sResponse := Confluence_Get(sUrl)
; Calls: b64Encode

sPassword := Login_GetPassword()
If (sPassword="") ; cancel
    return	
sUrl := RegExReplace(sUrl,"#.*","") ; remove link to section
WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("GET", sUrl, false) ; Async=false

sAuth := Confluence_BasicAuth(sUrl)
If (sAuth="") {
	TrayTip, Error, Confluence Authentication failed!,,3
	return
} 
WebRequest.setRequestHeader("Authorization", "Basic " . sAuth) 
WebRequest.setRequestHeader("Content-Type", "application/json")

WebRequest.Send()        
sResponse := WebRequest.responseText
return sResponse

} ; eofun
; ----------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------

Confluence_ViewInHierachy(sUrl :=""){
; Confluence_ViewInHierachy(sUrl*)
	
If (sUrl="")
	sUrl:= Browser_GetUrl()

sResponse := Confluence_Get(sUrl)
sPat = s)<meta name="ajs-page-id" content="([^"]*)">.*<meta name="ajs-space-key" content="([^"]*)">
If !RegExMatch(sResponse, sPat, sMatch) {
	TrayTipAutoHide("Confluence Error!","Getting PageId by API failed!")  	
	return
}
; extract section
If RegExMatch(sUrl,"#([^?&]*)",sSection) {
	sSection := RegExReplace(sSection1,".*-","")
	sLinkText := sLinkText . ": " . sSection
}
sLinkText := sLinkText . " - Confluence"

RegExMatch(sUrl, "https://[^/]*", sRootUrl)
sUrl := sRootUrl . "/pages/reorderpages.action?key=" . sMatch2 . "&openId=" . sMatch1 . "#selectedPageInHierarchy"
Run, %sUrl%
} ; eofun


Confluence_ViewPageInfo(sUrl :=""){
; Confluence_ViewPageIno(sUrl*)
		
If (sUrl="")
	sUrl:= Browser_GetUrl()


pageId := Confluence_GetPageId(sUrl)
If (pageId="")
	return
RegExMatch(sUrl, "https://[^/]*", sRootUrl)
sUrl := sRootUrl . "/pages/viewinfo.action?pageId=" . pageId
Run, %sUrl%
} ; eofun
	
; -------------------------------------------------------------------------------------------------------------------

Confluence_GetPageId(sUrl) {
; PageId := Confluence_GetPageId(sUrl)

If RegExMatch(sUrl,"pageId=(\d*)",sMatch)
	return sMatch1

sResponse := Confluence_Get(sUrl)
sPat = s)<meta name="ajs-page-id" content="([^"]*)">
If !RegExMatch(sResponse, sPat, sMatch) {
	TrayTipAutoHide("Confluence Error!","Getting PageId by API failed!")  	
	return
}
return sMatch1
} ; eofun


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
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Confluence_GetRootUrl(){
	If FileExist("PowerTools.ini") {
		IniRead, ConfluenceRootUrl, PowerTools.ini,Confluence,ConfluenceRootUrl
		If !(ConfluenceRootUrl="ERROR") ; Key found
			return ConfluenceRootUrl
	}
	sJiraUrl := Jira_GetRootUrl()
	If InStr(sJiraUrl,".atlassian.net") ; cloud version
		sUrl := sJiraUrl . "/wiki"
	return sUrl
}


; -------------------------------------------------------------------------------------------------------------------
Confluence_SearchSpace(sSpace,sQuery) {
	; Called by Attlasy Launcher c keyword
	sRootUrl := Confluence_GetRootUrl()

	If (sRootUrl = "") {
		PowerTools_ErrDlg("ConfluenceRootUrl not set!`nEdit the PowerTools.ini file!")
		return
	}

	If (sQuery="") { ; Open Space Root
		If InStr(sRootUrl,".atlassian.net")
			sUrl := sRootUrl . "/spaces/" . sSpace
		Else
			sUrl := sRootUrl . "/display/" . sSpace
	} Else {
		sCQL := Query2CQL(sQuery,sSpace)
		sUrl := sRootUrl . "/dosearchsite.action?cql=" . sCQL
	}

	Atlasy_OpenUrl(sUrl)

} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Confluence_Search(sUrl){
; Confluence Search - Search within current Confluence Space
; Called by: NWS.ahk-> QuickSearch (Win+F Hotkey)
; 
	static sConfluenceSearch, sSpace

	;https://confluenceroot/dosearchsite.action?cql=+space+%3D+%22PMPD%22+and+type+%3D+%22page%22+and+label+in+(%22best_practice%22%2C%22bitbucket%22)&queryString=code
	;https://confluenceroot/dosearchsite.action?cql=+space+%3D+%22PMPD%22+and+type+%3D+%22page%22+and+label+in+(%22best_practice%22%2C%22bitbucket%22)&queryString=code
	;https://confluenceroot/dosearchsite.action%3Fcql=%2Bspace%2B=%2B%22PMPD%22%2Band%2Btype%2B=%2B%22page%22%2Band%2Blabel%2Bin%2B%28%22best_practice%22%2C%22bitbucket%22%29&queryString=code
	;https://confluenceroot/dosearchsite.action?cql=+space+=+"PMPD"+and+type+=+"page"+and+label+in+("best_practice","bitbucket")&queryString=code
	; https://cr/dosearchsite.action?cql=siteSearch+~+%22move+page%22+and+space+%3D+%22PMT%22+and+type+%3D+%22page%22+and+label+%3D+%22confluence%22+and+label+%3D+%22microlearning%22&queryString=move+page

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
		sCQL := StrReplace(sCQL,"%3D","=")
		sCQL := StrReplace(sCQL,"%22",sQuote)
		sCQL := StrReplace(sCQL,"%2B","+")
		sCQL := StrReplace(sCQL,"%28","(")
		sCQL := StrReplace(sCQL,"%29",")")

		; Extract Labels (only AND label= not label in)	
		sPat = label\+=\+"([^"]*)"
		Pos=1
		While Pos :=    RegExMatch(sCQL, sPat, label,Pos+StrLen(label)) 
			sLabels := sLabels . " #" . label1 
		sLabels := Trim(sLabels) ; remove starting space			
		sDefSearch := sLabels

		; Extract Space Key
		sPat = space\+?=\+?"([^"]*)"
		If RegExMatch(sCQL,sPat,sSpace) 
			sSpace := sSpace1

		; Extract Search String
		
		sPat = siteSearch\+~\+"([^"]*)"
		If RegExMatch(sCQL,sPat,sSearchString) {
			; reverse regexp to wildcards
			sSearchString := StrReplace(sSearchString1,"%2F","/")
			sSearchString := StrReplace(sSearchString,"%20"," ")
			sSearchString := StrReplace(sSearchString,"+"," ")
			;If RegExMatch(sSearchString,"^/(.*)/$",sSearchString)
			;	sSearchString := StrReplace(sSearchString1,".*","*")
			sDefSearch := sDefSearch . " " . sSearchString
		}
		/* Comment out, because else search string is doubled
		If RegExMatch(sCQL,"\&queryString=(.*)",sSearchString) {
			sDefSearch := sDefSearch . " " . sSearchString1
		} 
		*/
		
	; Not from advanced search -> Extract Space 
	} Else {
		If RegExMatch(sUrl,ReRootUrl . "/label/([^/]*)/([^/]*)",sSpace) {
			sSpace := sSpace1
			sDefSearch := "#" . sSpace2
		} Else {
			sOldSpace := sSpace
			If RegExMatch(sUrl,ReRootUrl . "/display/([^/]*)",sSpace)
				sSpace := sSpace1
			Else If InStr(sUrl, "/pages/viewpage.action?pageId=") {
				sResponse := Confluence_Get(sUrl)
				sPat = s)<meta name="ajs-space-key" content="([^"]*)">
				RegExMatch(sResponse, sPat, sMatch)
				sSpace := sMatch1
			; } Else If  RegExMatch(sUrl,ReRootUrl . "/spaces/viewspace\.action\?key=([^&\?/]*)",sSpace)
			; 	sSpace := sSpace1
			} Else If  RegExMatch(sUrl,"\?key=([^&\?/]*)",sSpace)
				sSpace := sSpace1
			}
			
			If sSpace = %sOldSpace%
				sDefSearch := sConfluenceSearch
			Else
				sDefSearch = 
			
	}
	sDefSearch := Trim(sDefSearch)
	InputBox, sSearch , Confluence Search, Enter search string (use # for labels):,,640,125,,,,,%sDefSearch% 
	if ErrorLevel
		return
	sSearch := Trim(sSearch) 

	sCQL2 := Query2CQL(sSearch,sSpace)
	
	sSearchUrl := sRootUrl . "/dosearchsite.action?cql=" . sCQL2

	sConfluenceSearch := sDefSearch

	If sCQL ; not empty means update search 
		Send ^l
	Else
		Send ^n ; New Window
	Sleep 500
	Clip_Paste(sSearchUrl)
	Send {Enter}


	; /dosearchsite.action?cql=type+=+%22page%22+and+space=%22EMIK%22+and+label+%3D+%22r4j%22
	; /dosearchsite.action?cql=siteSearch+~+%22reuse%22+and+space+%3D+%22EMIK%22+and+type+%3D+%22page%22+and+label+%3D+%22r4j%22&queryString=reuse
} ; eofun



; -------------------------------------------------------------------------------------------------------------------

;.atlassian.net/wiki/search?spaces=ES&type=page&labels=meeting-notes%2Ctoto&text=test

Query2CQL(sSearch,sSpace) {

	sQuote = "

	; Convert labels to CQL
	sPat := "#([^#\s]*)" 
	Pos=1
	While Pos :=    RegExMatch(sSearch, sPat, label,Pos+StrLen(label)) {
		sCQLLabels := sCQLLabels . "+and+label+=+" . sQuote . label1 . sQuote
	} ; end while

	; remove labels from search string
	sSearch := RegExReplace(sSearch, sPat , "")
	sSearch := Trim(sSearch)

	; Check for leading wildcard -> convert to regexp
	If RegExMatch(sSearch,"^\*") {
		sSearch := "%2F" . sSearch "%2F" ; %2F is / encoded
		sSearch := StrReplace(sSearch,"*",".*")
	}
	
	If sSearch { ; not empty
		sCQL = siteSearch+~+"%sSearch%"+and+type+=+"page"
	} Else
		sCQL = type+=+"page"
	If sSpace
		sCQL = %sCQL%+and+space=%sQuote%%sSpace%%sQuote%

	If sCQLLabels ; not empty
		sCQL := sCQL . sCQLLabels 

	return sCQL
} ; eofun
	

; -------------------------------------------------------------------------------------------------------------------
Confluence_PersonalizeMention() {
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2023/07/confluence-personalize-mentions.html"
	return
}
ClipboardBackup := Clipboard
SendInput {Ctrl down}{Shift down}{Left}
SendInput {Ctrl up}{Shift up}
sClip := Clip_GetSelection(False)   

; Remove company in ()
If InStr(sClip,")") {
	While !RegExMatch(sClip,"\(.*\)") {
		SendInput {Ctrl down}{Shift down}
		SendInput {Left}
		SendInput {Ctrl up}{Shift up}
		sClip := Clip_GetSelection(False) 
	}
	SendInput {Backspace}
}

; Remove Lastname  in Lastname, Firstname format
SendInput {Ctrl down}{Left}
Sleep 500
SendInput {Backspace 2}{Right}
SendInput {Ctrl up}{Space}

Clipboard := ClipboardBackup

} ; eofun

Confluence_ShortenMention() {
; Remove company name in ()
If GetKeyState("Ctrl") {
	Run, "https://tdalon.blogspot.com/2020/11/confluence-personalize-mentions.html"
	return
}
ClipboardBackup := Clipboard
SendInput {Ctrl down}{Shift down}{Left}
SendInput {Ctrl up}{Shift up}
sClip := Clip_GetSelection(False)   

; Remove company in ()
If InStr(sClip,")") {
	While !RegExMatch(sClip,"\(.*\)") {
		SendInput {Ctrl down}{Shift down}
		SendInput {Left}
		SendInput {Ctrl up}{Shift up}
		sClip := Clip_GetSelection(False) 
	}
	SendInput {Backspace 2}
}
SendInput {Space}
Clipboard := ClipboardBackup
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