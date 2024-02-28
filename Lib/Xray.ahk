#Include <Jira>
#Include <Atlasy>

XrayCApiRootUrl := "https://eu.r4j-cloud.easesolutions.com"

; ---------------------------------------------------------------------------------

Xray_GetIssueKey(){
; sIssueKey := R4J_GetIssueKey()
; Get IssueKey from current selection or from Browser url
; returns only one IssueKey even if multiple are selected
    sSelection := Clip_GetSelection()
    If (sSelection = "") {
        If !Jira_IsWinActive()
            return
        sUrl := Browser_GetUrl()
        return Jira_Url2IssueKey(sUrl)
    } Else { ; selection
        sPat := "([A-Z]{3,})-([\d]{1,})"
        If RegExMatch(sSelection,sPat,sMatch)
            return sMatch1 . "-" . sMatch2
    }

} ; eofun
; ---------------------------------------------------------------------------------
Xray_OpenIssueSelection() {
    ; Called by Hotkey Ctrl+Shift+V
    sSelection := Clip_GetSelection()
    If (sSelection = "") {
        If !Jira_IsWinActive()
            return
        sUrl := Browser_GetUrl()
        RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
        ; issue detailed view
        If RegExMatch(sUrl,"/browse/([A-Z]*-\d*)",sIssueKey) {
            R4J_OpenIssue(sIssueKey1)
            return
        } Else If R4J_IsUrl(sUrl) { ; tree view -> jira detailed view
            sUrl := R4J_R2J(sUrl)
            return sUrl
        } 
    } Else {
        ; Loop on selection issue keys
        sPat := "([A-Z]{3,})-([\d]{1,})"
        Pos = 1 
        While Pos := RegExMatch(sSelection,sPat,sMatch,Pos+StrLen(sMatch)){ 
            sIssueKey := sMatch1 . "-" . sMatch2
            If InStr(sIssueKeyList,sIssueKey . ";")
                continue
            sIssueKeyList := sIssueKeyList . sIssueKey . ";"
            R4J_OpenIssue(sIssueKey)
        }
        return SubStr(sIssueKeyList,1,-1) ; remove ending ;
    }
} ; eofun
; ---------------------------------------------------------------------------------

Xray_X2J(sUrl) {
    RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
    If InStr(sUrl,".atlassian.net") {
        ; .atlassian.net/plugins/servlet/ac/com.easesolutions.jira.plugins.requirements/requirements-page-jira?project.id=15315#!tree?issueKey=RDMO-5
        RegExMatch(sUrl,".*/plugins/servlet/ac/com\.easesolutions\.jira\.plugins\.requirements/.*\?issueKey=([A-Z]{3,}-[\d]{1,})" ,sMatch) 
    } Else 
        RegExMatch(sUrl,".*/plugins/servlet/com\.easesolutions\.jira\.plugins\.requirements/project\?detail=[A-Z]*&issueKey=([A-Z]{3,}-[\d]{1,})" ,sMatch) 
    sUrl := sRootUrl . "/browse/" . sMatch1
    
    Atlasy_OpenUrl(sUrl)
    return sUrl
} ; eofun
; ---------------------------------------------------------------------------------

Xray_IsWinActive() {
    If !Browser_IsWinActive()
        return False
    sUrl := Browser_GetUrl()  
    return Xray_IsUrl(sUrl)
} ; eofun
; ---------------------------------------------------------------------------------

Xray_OpenIssues(IssueArray) {
    If IsObject(IssueArray) {
        for index, element in IssueArray 
            {
                 Xray_OpenIssue(element)
            }
    } Else {
        Xray_OpenIssue(IssueArray)
    }
    
} ; eofun
; ---------------------------------------------------------------------------------


Xray_Open(kw:="gs",pj:="",JiraRootUrl :="") {
    
    If (JiraRootUrl="")
        JiraRootUrl := Jira_GetRootUrl()
    IsCloud := InStr(JiraRootUrl, "atlassian.net")
    ; First switch for kw without JiraRootUrl required
    Switch kw {
        Case "doc":

            Switch pj {
                Case "rn":
                    sUrl := "https://docs.getxray.app/display/XRAYCLOUD/Release+Notes"
                Case "gs":
                    Xray_Open("gs")
                    return
                Default:
                    If IsCloud
                        sUrl := "https://docs.getxray.app/category/xc"
                    Else
                        sUrl := "https://docs.getxray.app/category/xs"
            }
            
        
        Case "gs":
            If IsCloud
                sUrl := JiraRootUrl . "/plugins/servlet/ac/com.xpandit.plugins.xray/xray-get-started"
            Else
                sUrl := JiraRootUrl . "/secure/XrayGetStarted.jspa"
        }
    
    If !(sUrl ="") {
        Atlasy_OpenUrl(sUrl)
        return
    }

    If (kw="")
        kw := "r" ; default to repository
    If (pj="")
        pj := Xray_GetProjectDef()

    
    If IsCloud { ; cloud

        Switch kw
        {
            Case "r","p", "e":
                sRootUrl := JiraRootUrl . "/projects/" . pj . "?selectedItem=com.atlassian.plugins.atlassian-connect-plugin:com.xpandit.plugins.xray__testing-board"
            Case "trace","m","te","tp","tr","ts","t","c","cov":
                sRootUrl := JiraRootUrl . "/plugins/servlet/ac/com.xpandit.plugins.xray"
                pjid := Jira_ProjectKey2Id(pj)
        }
        
        switch kw
        {
            Case "r":
                sUrl := sRootUrl . "#!page=test-repository&selectedFolder="
            Case "p":
                sUrl := sRootUrl . "#!page=test-plans"
            Case "e":
                sUrl := sRootUrl . "#!page=test-executions"
            ;---------------------- 
            Case "trace": ; traceability report
                sUrl := sRootUrl . "/traceability-report-page?project.key=" . pj . "&project.id=" . pjid
            Case "m":
                sUrl := sRootUrl . "/testplans-metrics-report-page?project.key=" . pj . "&project.id=" . pjid
            Case "te":
                sUrl := sRootUrl . "/testexecs-report-page?project.key=" . pj . "&project.id=" . pjid
            Case "tp":
                sUrl := sRootUrl . "/testplans-report-page?project.key=" . pj . "&project.id=" . pjid
            Case "tr":
                sUrl := sRootUrl . "/testruns-list-report-page?project.key=" . pj . "&project.id=" . pjid
            Case "ts":
                sUrl := sRootUrl . "/testsets-report-page?project.key=" . pj . "&project.id=" . pjid
            Case "t":
                sUrl := sRootUrl . "/tests-report-page?project.key=" . pj . "&project.id=" . pjid
            Case "c","cov":
                sUrl := sRootUrl . "/test-coverage-report-page?project.key=" . pj . "&project.id=" . pjid
        }
        

    } Else { ; server/dc TODO not implemented
        sUrl := JiraRootUrl . "/secure/XrayTestRepositoryAction!default.jspa?entityKey=" . sInput . "&path=%5Craven_all_tests"
    }

    Atlasy_OpenUrl(sUrl)

} ; eofun
; ---------------------------------------------------------------------------------

Xray_GetProjectDef() {
    ; Project := Xray_GetProjectDef()
    ; Return Xray default project key
    
    If InStr(A_ScriptName,"e.CoSys")
        Project := "T" . PowerTools_GetSetting("ECProject")
    Else
        Project := PowerTools_GetSetting("JiraProject")

    return Project
} ; eofun

Xray_OpenIssue(sIssueKey) {
    ; Open Issue in Xray Repository tree
    sJiraRootUrl := Jira_IssueKey2RootUrl(sIssueKey)
    sProjectKey := RegExReplace(sIssueKey,"\-.*")
    If InStr(sJiraRootUrl,".atlassian.net") { ; cloud version
        sProjectId := Jira_ProjectKey2Id(sProjectKey)
        sUrl := sJiraRootUrl . "/plugins/servlet/ac/com.easesolutions.jira.plugins.requirements/requirements-page-jira?project.id=" . sProjectId . "#!tree?issueKey=" . sIssueKey
    } Else
        sUrl := sJiraRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/project?detail=" . sProjectKey . "&issueKey=" . sIssueKey

    Atlasy_OpenUrl(sUrl)
} ; eofun

Xray_J2X(sIssueKey) {
    Xray_OpenIssue(sIssueKey)
} ; eofun


Xray_OpenProject(sProjectKey,view := "d") {
    ; Open project R4J tree
    sJiraRootUrl := Jira_IssueKey2RootUrl(sProjectKey)
    If InStr(sJiraRootUrl,".atlassian.net") { ; cloud version
        sProjectId := Jira_ProjectKey2Id(sProjectKey)
        sUrl := sJiraRootUrl . "/plugins/servlet/ac/com.easesolutions.jira.plugins.requirements/requirements-page-jira?project.id=" . sProjectId 
    } Else ; server/dc
        sUrl := sJiraRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/project?detail=" . sProjectKey 
    

    Switch view {
        Case "d":
            sUrl := sUrl . "#!tree"
        Case "c":
            sUrl := sUrl . "#!coverage"
        Case "t":
            sUrl := sUrl . "#!traceability"
    }

    Atlasy_OpenUrl(sUrl)
} ; eofun


; -------------------------------------------------------------------------------------------------------------------
Xray_IsUrl(sUrl){
return  InStr(sUrl,"/com.easesolutions.jira.plugins.requirements/") 
} ;eofun
; -------------------------------------------------------------------------------------------------------------------

Xray_Url2IssueKey(sUrl){
    ; /com.easesolutions.jira.plugins.requirements/project?detail=RDMO&issueKey=RDMO-14
    sKeyPat := "([A-Z]{3,}\-\d{1,})"
    if RegExMatch(sUrl,"&issueKey=" . sKeyPat,sMatch) ; url ending with Issue Key
        return sMatch1
    sUrl := RegExReplace(sUrl,"\?.*$","") ; remove optional arguments after ? in url
    if RegExMatch(sUrl,"/" . sKeyPat . "$",sMatch) ; url ending with Issue Key
        return sMatch1
} ; eofun


; -------------------------------------------------------------------------------------------------------------------



; ----------------------------------------------------------------------
Xray_Auth(sUrl:="",sToken:="") {
; Get Auth String for R4J authentification for a specific instance
; sAuth := R4J_GetApiToken(sUrl:="",sToken:="")
; Default Url is Setting JiraRootUrl
    
    
If !RegExMatch(sUrl,"^http")  ; missing root url or default url
    sUrl := Jira_GetRootUrl() . sUrl

; Read Jira PowerTools.ini setting
If !FileExist("PowerTools.ini") {
    PowerTools_ErrDlg("No PowerTools.ini file found!")
    return
}

App := "Xray"
IniRead, Auth, PowerTools.ini,%App%,%App%Auth
If (Auth="ERROR") { ; 
    PowerTools_ErrDlg(App . "Auth key not found in PowerTools.ini file [" . App . "] section!")
    return
}

JsonObj := Jxon_Load(Auth)
For i, inst in JsonObj 
{
    url := inst["url"]
    If InStr(sUrl,url) {
        ApiToken := inst["apitoken"]
        If (ApiToken="") { 
            PowerTools_ErrDlg("ApiToken is not defined in PowerTools.ini file [" . App . "] section, " . App . "Auth key for url '" . url . "'!")
            return
        }
        break
    }
}
If (ApiToken="") { ; Section [Jira] Key JiraAuth not found
    RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
    PowerTools_ErrDlg("No instance defined in PowerTools.ini file [" . App . "] section, " . App . "Auth key for url '" . sRootUrl . "'!")
    return
} 	

return ApiToken

} ; eofun
; ----------------------------------------------------------------------


Xray_Get(sPath,sToken:=""){
; sResponse := Xray_Get(sPath)
; sPath starting with /

; ; https://easesolutions.atlassian.net/wiki/spaces/R4JC/pages/2250506241/REST+API
If (sToken="") {
    sToken := Xray_Auth()
}

If (sToken="") {
	TrayTip, Error, Xray Authentication not defined!,,3
	return
}

WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")

sUrl := "https://eu.r4j-cloud.easesolutions.com/rest/api/1" . sPath
WebRequest.Open("GET", sUrl, false) ; Async=false

sAuth := "JWT " . sToken
WebRequest.setRequestHeader("Authorization", sAuth) 
WebRequest.setRequestHeader("Content-Type", "application/json")
WebRequest.Send()        
return WebRequest.responseText
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

Xray_Post(sPath,sBody,sToken:=""){
; sResponse := Xray_Post(sPath,sBody)
; sPath without starting /

; ; https://easesolutions.atlassian.net/wiki/spaces/R4JC/pages/2250506241/REST+API
If (sToken="") {
    sToken := Xray_Auth()
}

If (sToken="") {
    TrayTip, Error, Xray Authentication not defined!,,3
    return
}

WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")

sUrl := "https://eu.r4j-cloud.easesolutions.com/rest/api/1" . sPath
WebRequest.Open("POST", sUrl, false) ; Async=false

sAuth := "JWT " . sToken
WebRequest.setRequestHeader("Authorization", sAuth) 
WebRequest.setRequestHeader("Content-Type", "application/json")
If (sBody = "")
	WebRequest.Send() 
Else
	WebRequest.Send(sBody)       
return WebRequest.responseText
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

