;#Include <IntelliPaste>
;#Include <Clip>

; ----------------------------------------------------------------------
Confluence_BasicAuth(sUrl:="",sToken:="") {
	; Get Auth String for basic authentification
	; sAuth := Confluence_BasicAuth(sUrl:="",sToken:="")
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
			JiraUserName :=  Jira_GetUserName()
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
} ; eofun

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
; Calls: Confluence_BasicAuth

If !RegExMatch(sUrl,"^http")  ; missing root url or default url
	sUrl := Confluence_GetRootUrl() . "/" . RegExReplace(sUrl,"^/")

sAuth := Confluence_BasicAuth(sUrl,sPassword)
If (sAuth="") {
	TrayTip, Error, Confluence Authentication failed!,,3
	return
}
sUrl := RegExReplace(sUrl,"#.*","") ; remove link to section
WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("GET", sUrl, false) ; Async=false

; https://developer.atlassian.com/cloud/jira/platform/basic-auth-for-rest-apis/

sAuth := "Basic " . sAuth
WebRequest.setRequestHeader("Authorization", sAuth) 
WebRequest.setRequestHeader("Content-Type", "application/json")
WebRequest.Send()        
return WebRequest.responseText
} ; eofun
; ----------------------------------------------------------------------
Confluence_Post(sUrl,sBody:="",sPassword:=""){
; Syntax: sResponseText := Confluence_Post(sUrl,sBody,sPassword*)
; Calls: Confluence_BasicAuth
If !RegExMatch(sUrl,"^http")  ; missing root url or default url
	sUrl := Confluence_GetRootUrl() . "/" . RegExReplace(sUrl,"^/")

sAuth := Confluence_BasicAuth(sUrl,sPassword)
If (sAuth="") {
	TrayTip, Error, Confluence Authentication failed!,,3
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
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Confluence_WebRequest(sReqType,sUrl,sBody:="",sPassword:=""){
; Syntax: WebRequest := Confluence_WebRequest(sReqType,sUrl,sBody:="",sPassword:="")
; Output WebRequest with fields Status and ResponseText
	
; Calls: Confluence_BasicAuth
		
If !RegExMatch(sUrl,"^http")  ; missing root url or default url
	sUrl := Confluence_GetRootUrl() . "/" . RegExReplace(sUrl,"^/")
sAuth := Confluence_BasicAuth(sUrl,sPassword)
If (sAuth="") {
	TrayTip, Error, Confluence Authentication failed!,,3
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
; -------------------------------------------------------------------------------------------------------------------


Confluence_ViewInHierachy(sUrl :=""){
; Confluence_ViewInHierachy(sUrl*)
	
If (sUrl="")
	sUrl:= Browser_GetUrl()

RegExMatch(sUrl, "https://[^/]*", sRootUrl)
If InStr(sRootUrl,".atlassian.net") { ; Cloud version
	sRootUrl := sRootUrl . "/wiki"
	If RegExMatch(sUrl, "\.atlassian\.net/wiki/spaces/([^/]*)/pages/([^/]*)",sMatch) {
		spaceKey := sMatch1
		pageId := sMatch2
	} Else {
		TrayTipAutoHide("Confluence Error!","Getting PageId from Url failed!")  	
		return
	}

} Else { ; server
	sResponse := Confluence_Get(sUrl)
	sPat = s)<meta name="ajs-page-id" content="([^"]*)">.*<meta name="ajs-space-key" content="([^"]*)">
	If !RegExMatch(sResponse, sPat, sMatch) {
		TrayTipAutoHide("Confluence Error!","Getting PageId and SpaceKey by API failed!")  	
		return
	}
	spaceKey := sMatch2
	pageId := sMatch1
}

sUrl := sRootUrl . "/pages/reorderpages.action?key=" . spaceKey . "&openId=" . pageId . "#selectedPageInHierarchy"
Atlasy_OpenUrl(sUrl)
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Confluence_ViewPageInfo(sUrl :=""){
; Confluence_ViewPageInfo(sUrl*)
		
If (sUrl="")
	sUrl:= Browser_GetUrl()

pageId := Confluence_GetPageId(sUrl)
If (pageId="")
	return
RegExMatch(sUrl, "https://[^/]*", sRootUrl)
If InStr(sRootUrl,".atlassian.net") ; Cloud version
	sRootUrl := sRootUrl . "/wiki"
sUrl := sRootUrl . "/pages/viewinfo.action?pageId=" . pageId

Atlasy_OpenUrl(sUrl)
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Confluence_ViewPageHistory(sUrl :=""){
	; Confluence_ViewPageHistory(sUrl*)
			
	If (sUrl="")
		sUrl:= Browser_GetUrl()

	RegExMatch(sUrl, "https://[^/]*", sRootUrl)
	If InStr(sRootUrl,".atlassian.net") { ; Cloud version
		sRootUrl := sRootUrl . "/wiki"
		If RegExMatch(sUrl, "\.atlassian\.net/wiki/spaces/([^/]*)/pages/([^/]*)",sMatch) {
			spaceKey := sMatch1
			pageId := sMatch2
		} Else {
			TrayTipAutoHide("Confluence Error!","Getting PageId from Url failed!")  	
			return
		}

	} Else { ; server
		sResponse := Confluence_Get(sUrl)
		sPat = s)<meta name="ajs-page-id" content="([^"]*)">.*<meta name="ajs-space-key" content="([^"]*)">
		If !RegExMatch(sResponse, sPat, sMatch) {
			TrayTipAutoHide("Confluence Error!","Getting PageId and SpaceKey by API failed!")  	
			return
		}
		spaceKey := sMatch2
		pageId := sMatch1
	}
	
	sUrl := sRootUrl . "/spaces/" . spaceKey . "/history/" . pageId
	Atlasy_OpenUrl(sUrl)

} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Confluence_ViewAttachments(sUrl :=""){
	; Confluence_ViewAttachments(sUrl*)
			
	If (sUrl="")
		sUrl:= Browser_GetUrl()
	pageId := Confluence_GetPageId(sUrl)
	If (pageId="")
		return
	RegExMatch(sUrl, "https://[^/]*", sRootUrl)
	If InStr(sRootUrl,".atlassian.net") ; Cloud version
		sRootUrl := sRootUrl . "/wiki"
	sUrl := sRootUrl . "/pages/viewpageattachments.action?pageId=" . pageId
	Run, %sUrl%
} ; eofun
	
; -------------------------------------------------------------------------------------------------------------------

Confluence_GetSpace(sUrl) {
If InStr(sUrl,".atlassian.net")  ; cloud
	If RegExMatch(sUrl, "\.atlassian\.net/wiki/spaces/([^/]*)/pages/",sMatch)
		return sMatch1
Else
	TrayTipAutoHide("Confluence Error!","GetSpace not implemented for server/datacenter!")  	
return
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Confluence_GetPageId(sUrl) {
; PageId := Confluence_GetPageId(sUrl)
If InStr(sUrl,".atlassian.net") { ; cloud
	If RegExMatch(sUrl, "\.atlassian\.net/wiki/spaces/([^/]*)/pages/(?:edit/|edit\-v2/|)([^/]*)",sMatch)
		return sMatch2
	Else {
		TrayTipAutoHide("Confluence Error!","Getting PageId from Url failed!")  	
		return
	}
}
If RegExMatch(sUrl,"pageId=(\d*)",sMatch)
	return sMatch1

; for server only
sResponse := Confluence_Get(sUrl)
sPat = sU)<meta name="ajs-page-id" content="(.*)">
If !RegExMatch(sResponse, sPat, sMatch) {
	TrayTipAutoHide("Confluence Error!","Getting PageId by API failed!")  	
	return
}
return sMatch1
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Confluence_GetPageTitle(sUrl) {
	; PageId := Confluence_GetPageId(sUrl)
	
	If InStr(sUrl,".atlassian.net") { ; cloud
		pageId := Confluence_GetPageId(sUrl)
		RegExMatch(sUrl, "https://[^/]*", rootUrl) ; Get RootUrl
		RestUrl := rootUrl  . "/wiki/rest/api/content/" .  pageId
		sResponse := Confluence_Get(RestUrl)
		sPat = "title":"([^"]*)"
	} Else { ; server/DC
		sResponse := Confluence_Get(sUrl)
		sPat = sU)<meta name="ajs-page-id" content="(.*)">
	}

	If !RegExMatch(sResponse, sPat, sMatch) {
		TrayTipAutoHide("Confluence Error!","Getting Page Title by API failed!")  	
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
		If InStr(sRootUrl,".atlassian.net") {
			sQuery := Trim(sQuery) 
			sParam := Query2Param(sQuery,sSpace)
			sUrl := sRootUrl . "/search" . sParam
		} Else { ; Server/DC
			sCQL := Query2CQL(sQuery,sSpace)
			sUrl := sRootUrl . "/dosearchsite.action?cql=" . sCQL
		}
	}
	Atlasy_OpenUrl(sUrl)
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Confluence_GetRootUrl() {
; Get Confluence Root Url
; From current Browser window (Jira or Confluence opened)
; From PowerTools Setting (JiraRootUrl for Cloud or ConfluenceRootUrl for Server)

; From Browser Url
If Browser_WinActive() {
	sUrl := Browser_GetUrl()
	If Jira_IsUrl(sUrl) {
		RegExMatch(sUrl,"https?://[^/]*",JiraRootUrl)
		If InStr(JiraRootUrl,".atlassian.net")
			return JiraRootUrl . "/wiki"
	}
	If Confluence_IsUrl(sUrl) {
		RegExMatch(sUrl,"https?://[^/]*",ConfluenceRootUrl)
		If InStr(ConfluenceRootUrl,".atlassian.net")
			ConfluenceRootUrl := ConfluenceRootUrl . "/wiki"
		return ConfluenceRootUrl
	}
}
JiraRootUrl := PowerTools_GetSetting("JiraRootUrl")
If InStr(JiraRootUrl,".atlassian.net")
	return JiraRootUrl . "/wiki"

If FileExist("PowerTools.ini") {
	IniRead, ConfluenceRootUrl, PowerTools.ini,Confluence,ConfluenceRootUrl
	If !(ConfluenceRootUrl="ERROR") ; Key found
		return ConfluenceRootUrl
}

} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Confluence_QuickOpen(sSearch,sSpace :="") { ;@fun_confluence_quickopen@
; called by Altasy c -o keyword
	If (sSpace ="") {
		pat := "^\-?s\s([^\s]*)"
		If RegExMatch(sSearch,pat,sMatch) {
			sSpace := sMatch1
			sSearch := Trim(RegExReplace(sSearch,pat))
		}
	}
	cql := Query2CQL(sSearch,sSpace)
	cql := uriEncode(cql)
	rootUrl := Confluence_GetRootUrl()
	restUrl := rootUrl . "/rest/api/content/search?cql=" . cql
	response := Confluence_Get(restUrl)
	JsonObj := Jxon_Load(response)
	resultsObj := JsonObj["results"]
	pageId := resultsObj[1]["id"]
	If (pageId="") {
		TrayTip, Warning, Confluence no results found!,,2
		;TrayTipAutoHide("Confluence: No result","No results found matching the query!")  
		return
	}
	url := rootUrl . "/pages/viewpage.action?pageId=" . pageId
	Atlasy_OpenUrl(url)

	; /history e.g. <root>/wiki/rest/api/content/131432476/history will get lastUpdated:by, "when": "2024-01-15T13:24:58.927Z" and previousVersion:by->"Display Name"
	; createdBy->{publicName}, createdDate

} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Confluence_QuickSearch(sSearch,sSpace :="") {
; quick open first result of input search
; called by Altasy c -o keyword
	If (sSpace ="") {
		pat := "^\-?s\s([^\s]*)"
		If RegExMatch(sSearch,pat,sMatch) {
            sSpace := sMatch1
			sSearch := Trim(RegExReplace(sSearch,pat))
        }
	}
	cql := Query2CQL(sSearch,sSpace)
	cql := uriEncode(cql)

	rootUrl := Confluence_GetRootUrl()
	restUrl := rootUrl . "/rest/api/content/search?cql=" . cql
	response := Confluence_Get(restUrl)
	JsonObj := Jxon_Load(response)
	resultsObj := JsonObj["results"]
	For i, page in resultsObj 
	{
		title := page["title"]
		id := page["id"]
		pageRestUrl := rootUrl . "/rest/api/content/" . id
	}

	; /history e.g. <root>/wiki/rest/api/content/131432476/history will get lastUpdated:by, "when": "2024-01-15T13:24:58.927Z" and previousVersion:by->"Display Name"
	; createdBy->{publicName}, createdDate

} ; eofun
; -------------------------------------------------------------------------------------------------------------------


Confluence_Search(sUrl:=""){
If (sUrl = "")
	sUrl := Confluence_GetRootUrl()
If InStr(sUrl,".atlassian.net") ; Cloud
	ConfluenceSearch_Cloud(sUrl)
Else 
	ConfluenceSearch_Server(sUrl)
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

ConfluenceSearch_Cloud(sUrl) {
	; $root/wiki/search?text=ux&labels=cloud%2Cjira&spaces=PMT&type=page

If RegExMatch(sUrl,"U)/spaces/(.*)/",sMatch) {
	sSpace := sMatch1
}

InputBox, sSearch , Confluence Search, Enter search string (use # for labels):,,640,125,,,,,%sDefSearch% 
If ErrorLevel
	return
sSearch := Trim(sSearch) 
sParam := Query2Param(sSearch,sSpace)
RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
sSearchUrl := sRootUrl . "/wiki/search" . sParam


} ; eofun

; -------------------------------------------------------------------------------------------------------------------

ConfluenceSearch_Server(sUrl){
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
	If ErrorLevel
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


Query2Param(sSearch,sSpace) {
;.atlassian.net/wiki/search?spaces=ES&type=page&labels=meeting-notes%2Ctoto&text=test
	sQuote = "

	; Convert labels to params
	sPat := "#([^#\s]*)" 
	Pos=1
	While Pos :=    RegExMatch(sSearch, sPat, label,Pos+StrLen(label)) {
		sLabels := sLabels . "%2C" . label1
	} ; end while
	
	If !(sLabels="") {
		sLabels := RegExReplace(sLabels,"^%2C","&labels=")
	}

	; remove labels from search string
	sSearch := RegExReplace(sSearch, sPat , "")
	sSearch := Trim(sSearch)

	; Check for leading wildcard -> convert to regexp
	If RegExMatch(sSearch,"^\*") {
		sSearch := "%2F" . sSearch "%2F" ; %2F is / encoded
		sSearch := StrReplace(sSearch,"*",".*")
	}
	
	If sSearch ; not empty
		sParam := sParam . "&text=" . sSearch
	If sSpace
		sParam := sParam . "&spaces=" . sSpace

	If sLabels ; not empty
		sParam := sParam . sLabels

	sParam := RegExReplace(sParam,"^&","?")
	return sParam
} ; eofun
; -------------------------------------------------------------------------------------------------------------------



Query2CQL(sSearch,sSpace) {
	sQuote = "

	; Shortcuts for labels
	sSearch := RegExReplace(sSearch,"#t(\s?)","#tool$1")
	sSearch := RegExReplace(sSearch,"#pt(\s?)","#powertool$1")
	sSearch := RegExReplace(sSearch,"#m(\s?)","#method$1")

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
Confluence_VersionComment(pageId,versionNumber:="",message:="") {
; https://developer.atlassian.com/cloud/confluence/rest/v1/api-group-content-versions/#api-wiki-rest-api-content-id-version-post

rootUrl := Confluence_GetRootUrl()
If (versionNumber="") { ; get version number from current page
	url := rootUrl . "/api/v2/pages/" . pageId
	resp := Confluence_Get(url)
	respJson := Jxon_Load(resp)
	version := respJson["version"]
	versionNumber := version["number"]
}

; status: historical or current

; Create dummy/empty version (required for restoring non-current version)
; PUT
empty_versionNumber := versionNumber + 1
bodyData = {"id": "%pageId%","status": "current","title": "Dummy (PowerTool:Confluence:VersionComment)","body": {"representation": "storage", "value": ""},"version": { "number": %empty_versionNumber%,"message": "empty"}}
url := rootUrl . "/api/v2/pages/" . pageId
WebRequest := Confluence_WebRequest("PUT",url,bodyData)

; Restore page
If (message="") { ; prompt user for comment version
	InputBox, message , Version Comment, Enter version description:,,640,125,,,,, %sDefSearch%
	if ErrorLevel
		return
}
bodyData = {"operationKey": "restore","params": {"versionNumber": %versionNumber%,"message": "%message%","restoreTitle": true}}
url := rootUrl . "/rest/api/content/" . pageId . "/version"
resp := Confluence_Post(url,bodyData)

; Delete empty/dummy version
; https://developer.atlassian.com/cloud/confluence/rest/v1/api-group-content-versions/#api-wiki-rest-api-content-id-version-versionnumber-delete
; DELETE /wiki/rest/api/content/{id}/version/{versionNumber}
url := rootUrl . "/rest/api/content/" . pageId . "/version/" . empty_versionNumber
WebRequest := Confluence_WebRequest("DELETE",url)

; Delete original version
url := rootUrl . "/rest/api/content/" . pageId . "/version/" . versionNumber
WebRequest := Confluence_WebRequest("DELETE",url)


; Copy Page Content
; https://community.atlassian.com/t5/Confluence-questions/Re-How-to-edit-the-page-content-using-rest-api/qaq-p/905680/comment-id/121163#M121163

; GET <INSTANCE>/rest/api/content/<PAGEID>?expand=body.storage,version
url := rootUrl . "/rest/api/content/" . pageId . "?expand=body.storage" 

} ; eofun


; -------------------------------------------------------------------------------------------------------------------
Confluence_GetCurrentVersion(pageId) {
	rootUrl := Confluence_GetRootUrl()
	; Get page current version information from pageId
	; "version": {"number": 663,"message": "","minorEdit": false,"authorId": "6141c0a0eaef3400697a1834","createdAt": "2024-01-31T07:45:30.869Z"}
	url := rootUrl . "/api/v2/pages/" . pageId 
	resp := Confluence_Get(url)
	respJsonObj :=  Jxon_Load(resp)
	return respJsonObj["version"] 
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Confluence_GetPageInfo(pageId,rootUrl:="") {
If (rootUrl="")
	rootUrl := Confluence_GetRootUrl()
; Get page current version information from pageId
; "version": {"number": 663,"message": "","minorEdit": false,"authorId": "6141c0a0eaef3400697a1834","createdAt": "2024-01-31T07:45:30.869Z"}
url := rootUrl . "/api/v2/pages/" . pageId 
resp := Confluence_Get(url)
return  Jxon_Load(resp)
} ; eofun
	
; -------------------------------------------------------------------------------------------------------------------
Confluence_Version2Link(pageId,versionNumber) {
rootUrl := Confluence_GetRootUrl()
; From version number get page Id
; https://developer.atlassian.com/cloud/confluence/rest/v2/api-group-version/#api-pages-id-versions-get
url := rootUrl . "/api/" . pageId . "/versions"

resp := Confluence_Get(url)

; look for number=versionNumber in results array
respJsonObj :=  Jxon_Load(resp)
results := respJsonObj["results"]

For i, ver in results
{
	url := 
	If (ver["number"] = versionNumber) {
		pageId := ver["page"]["id"]
		break
	}
}

If (pageId ="")
	Return
link := rootUrl . "/pages/viewpage.action?pageId=" . pageId
return link
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Confluence_Reorder(sUrl:="") {
	If (sUrl="")
		sUrl:= Browser_GetUrl()
	If (sUrl="")
		return
	spaceKey := Confluence_GetSpace(sUrl)
	If (spaceKey="")
		return
	RegExMatch(sUrl,"https?://[^/]*",rootUrl)
	rootUrl := RegExReplace(rootUrl,"\.atlassian\.net$",".atlassian.net/wiki") ; append wiki to Cloud root url if no
	
	sUrl := rootUrl . "/pages/reorderpages.action?key=" . spaceKey
	Atlasy_OpenUrl(sUrl)
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Confluence_Redirect(sUrl,tgtRootUrl:="") {
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
		IniRead, tgtRootUrl, PowerTools.ini,Confluence,ConfluenceRootUrl
		If (tgtRootUrl="ERROR") { 
			PowerTools_ErrDlg("ConfluenceRootUrl key not found in PowerTools.ini file [Confluence] section!")
			return
		}
		
	} Else { ; from server to Cloud
		Loop, Parse, JiraRootUrls,`,
		{
			If InStr(A_LoopField,".atlassian.net")
				JiraCloudRootUrl := A_LoopField
		}
		If (JiraCloudRootUrl ="") {
			PowerTools_ErrDlg("JiraRootUrls does not contain a Cloud url!")
			return
		}
		tgtRootUrl := JiraCloudRootUrl . "/wiki"
	}

} Else {
	tgtRootUrl := RegExReplace(tgtRootUrl,"/$") ; remove trailing sep
	tgtRootUrl := RegExReplace(tgtRootUrl,"\.atlassian\.net$",".atlassian.net/wiki") ; append wiki to Cloud root url if no
}

; https://community.atlassian.com/t5/Confluence-questions/Post-migration-to-the-cloud-Redirection/qaq-p/1118935

If InStr(sUrl,".atlassian.net") { ; cloud
	If RegExMatch(sUrl,"/wiki/spaces/([^/]*)/pages/([^/]*)/([^/#]*)",sMatch) {
		spaceKey := sMatch1
		;pageId := sMatch2
		pageTitle := sMatch3
	} Else {
		TrayTipAutoHide("Confluence Error!","Getting Page Title from Url failed!")  
		return
	}
} Else { ; server/DC
	sResponse := Confluence_Get(sUrl)
	; "ajs-space-key" "ajs-page-title"
	; Get Page Name
	sPat = sU)<meta name="ajs-page-title" content="(.*)">
	If !RegExMatch(sResponse, sPat, sMatch) {
		TrayTipAutoHide("Confluence Error!","Getting Page Title by API failed!")  
		return
	}
	pageTitle := sMatch1

	; Get space key
	sPat = sU)<meta name="ajs-space-key" content="(.*)">
	If !RegExMatch(sResponse, sPat, sMatch) {
		TrayTipAutoHide("Confluence Error!","Getting Space Key by API failed!")  	
		return
	}
	spaceKey := sMatch1
}

sUrl := tgtRootUrl . "/display/" . spaceKey . "/" . pageTitle
return sUrl

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Confluence_GetVerLink(url:="") { ; @fun_Confluence_GetVerLink@
; Get Link information from current page Url
; sHtml := Confluence_GetVerLink(url:="")
If (url="")
	url := Browser_GetUrl()
If (url="")
	return
; Get current pageId
pageId := Confluence_GetPageId(url)
rootUrl := Confluence_GetRootUrl()
pageInfo := Confluence_GetPageInfo(pageId,rootUrl)
; If Confluence Document History page, get parentId
If RegExMatch(pageInfo["title"],"i)^(Document|Work Product) History") {
	pageId := pageInfo["parentId"]
	pageInfo := Confluence_GetPageInfo(pageId,rootUrl)
}
;MsgBox % Jxon_Dump(pageInfo) ; DBG
version := pageInfo["version"]
sText := version["createdAt"] . " (v." . version["number"] . ")"
sLink :=  rootUrl . "/pages/viewpage.action?pageId=" . pageId . "&pageVersion=" . version["number"]
sHtml := "<a href=""" . sLink . """>" . sText . "</a>"
return sHtml
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
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
