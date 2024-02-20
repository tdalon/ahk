#Include <Jira>
#Include <Atlasy>
#Include Jxon.ahk

; ---------------------------------------------------------------------------------

R4J_GetIssueKey(){
; sIssueKey := R4J_GetIssueKey()
; Get IssueKey from current selection or from Browser url
; returns only one IssueKey even if multiple are selected
sSelection := Clip_GetSelection()
If (sSelection = "") {
    If !Jira_IsWinActive()
        return
    sUrl := Browser_GetUrl()
    If R4J_IsUrl(sUrl) {
        issueKey := R4J_Url2IssueKey(sUrl)
        If (issueKey != "")
            return issueKey
    }
    return Jira_Url2IssueKey(sUrl)
} Else { ; selection
    sPat := "([A-Z]{3,})-([\d]{1,})"
    If RegExMatch(sSelection,sPat,sMatch)
        return sMatch1 . "-" . sMatch2
}

} ; eofun
; ---------------------------------------------------------------------------------
R4J_OpenIssueSelection() {
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

R4J_R2J(sUrl) {
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

R4J_IsWinActive(){
    If !Browser_IsWinActive()
        return False
    sUrl := Browser_GetUrl()  
    return R4J_IsUrl(sUrl)
} ; eofun
; ---------------------------------------------------------------------------------

R4J_OpenIssues(IssueArray) {
    If IsObject(IssueArray) {
        for index, element in IssueArray 
            {
                 R4J_OpenIssue(element)
            }
    } Else {
        R4J_OpenIssue(IssueArray)
    }
    
} ; eofun
; ---------------------------------------------------------------------------------


R4J_OpenIssue(sIssueKey) {
    ; Open Issue in project R4J tree
    sJiraRootUrl := Jira_IssueKey2RootUrl(sIssueKey)
    sProjectKey := RegExReplace(sIssueKey,"\-.*")
    If InStr(sJiraRootUrl,".atlassian.net") { ; cloud version
        sProjectId := Jira_ProjectKey2Id(sProjectKey)
        sUrl := sJiraRootUrl . "/plugins/servlet/ac/com.easesolutions.jira.plugins.requirements/requirements-page-jira?project.id=" . sProjectId . "#!tree?issueKey=" . sIssueKey
    } Else
        sUrl := sJiraRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/project?detail=" . sProjectKey . "&issueKey=" . sIssueKey

    Atlasy_OpenUrl(sUrl)
} ; eofun
; ---------------------------------------------------------------------------------

R4J_J2R(sIssueKey) {
    R4J_OpenIssue(sIssueKey)
} ; eofun

; ---------------------------------------------------------------------------------

R4J_OpenProject(sProjectKey,view := "d",JiraRootUrl:="") {
    ; Open project R4J tree
    If (JiraRootUrl = "")
        JiraRootUrl := Jira_GetRootUrl()
    If InStr(JiraRootUrl,".atlassian.net") { ; cloud version
        sProjectId := Jira_ProjectKey2Id(sProjectKey)
        sUrl := JiraRootUrl . "/plugins/servlet/ac/com.easesolutions.jira.plugins.requirements/requirements-page-jira?project.id=" . sProjectId 
    } Else ; server/dc
        sUrl := JiraRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements" 
    
    Switch view {
        Case "d","doc","tree":
            If InStr(JiraRootUrl,".atlassian.net")
                sUrl := sUrl . "#!tree"
            Else
                sUrl := sUrl . "/project?detail=" . sProjectKey
        Case "c","cov","coverage":
            If InStr(JiraRootUrl,".atlassian.net")
                sUrl := sUrl . "#!coverage"
            Else
                sUrl := sUrl . "/coverage?prj=" . sProjectKey
        Case "t","trace","traceability":
            If InStr(JiraRootUrl,".atlassian.net")
                sUrl := sUrl . "#!traceability"
            Else
                sUrl := sUrl . "/tracematrix?prj=" . sProjectKey
            
    }

    Atlasy_OpenUrl(sUrl)
} ; eofun


; -------------------------------------------------------------------------------------------------------------------
R4J_IsUrl(sUrl){
; valid both for server and cloud
return  InStr(sUrl,"/com.easesolutions.jira.plugins.requirements/") 
} ;eofun
; -------------------------------------------------------------------------------------------------------------------

R4J_Url2IssueKey(sUrl) {
; issueKey := R4J_Url2IssueKey(sUrl)

If R4J_IsUrl(sUrl) {
    sKeyPat := "([A-Z]{3,}\-\d{1,})"
    If RegExMatch(sUrl,"(&|\?)issueKey=" . sKeyPat,sMatch) ; cloud can have multiple ? in url as parameter
        return sMatch2
    ; Root level: return project key
    If RegExMatch(sUrl,"\?project\.id=(.*)#",sMatch) ; cloud can have multiple ? in url as parameter 
        return Jira_ProjectId2Key(sMatch1)
} Else If Jira_IsUrl(sUrl) {
    return Jira_Url2IssueKey(sUrl)
}

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

R4J_Url2ProjectKey(sUrl) {
; projectKey := R4J_Url2ProjectKey(sUrl)
    If R4J_IsUrl(sUrl) {
        If RegExMatch(sUrl,"prj=([A-Z]*)",sMatch) ; ; server version
            return sMatch1
        If RegExMatch(sUrl,"project\.id=(\d*)",sMatch) ; cloud version
            return Jira_ProjectId2Key(sMatch1)

        IssueKey := R4J_Url2IssueKey(sUrl)
        If !(IssueKey="")
            return RegExReplace(IssueKey,"-.*")
    } Else If Jira_IsUrl(sUrl) {
        return Jira_Url2ProjectKey(sUrl)
    }
} ; eofun


; -------------------------------------------------------------------------------------------------------------------

R4J_GetProjectDef() {
    ; Project := R4J_GetProjectDef()
    ; Return R4J default project key
    
    ; Get project from current browser window url
    If Browser_WinActive() {
        sUrl := Browser_GetUrl()
        Project := R4J_Url2ProjectKey(sUrl)
        If !(Project="")
            return Project
    }

    If InStr(A_ScriptName,"e.CoSys")
        ECPRoject := PowerTools_GetSetting("ECProject")
        If !(ECProject ="")
            Project := "R" . ECProject
    Else
        Project := PowerTools_GetSetting("JiraProject")

    return Project
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
; Architecture functions with prefix R4J_Arch_
; -------------------------------------------------------------------------------------------------------------------

R4J_Arch_OpenIssue(sIssueKey) {
sJiraRootUrl := Jira_IssueKey2RootUrl(sIssueKey)
sProjectKey := RegExReplace(sIssueKey,"\-.*")
sProjectKey := "O" . SubStr(sProjectKey,2)

sUrl := sJiraRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/project?detail=" . sProjectKey . "&issueKey=" . sIssueKey
Run, %sUrl% 
return sUrl
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

R4J_Arch_OpenIssueSelection() {
; Called by Hotkey Ctrl+Shift+V
ClipSaved := ClipboardAll

; Try first if selection is manually set (Triple click doesn't work in Outlook #35)
Clipboard := ""		;Clear the clipboard
Send, ^c			;Copy (Ctrl+C)


sSelection := Clipboard
Clipboard := ClipSaved ; restore clipboard

If (sSelection = "") {
    If Jira_IsWinActive() {
        sUrl := Browser_GetUrl()
        RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
        ; issue detailed view
        If RegExMatch(sUrl,StrReplace(sRootUrl,".","\.") . "/browse/([A-Z]*)-(\d*)",sIssueKey) 
        {
            sIssueKey := sIssueKey1 . "-" . sIssueKey2
            sUrl := R4J_Arch_OpenIssue(sIssueKey)
            return sUrl
        } 
    }
    return
}

; Loop on issue keys
sPat := "([A-Z]{3,})-([\d]{1,})"
Pos = 1 
While Pos := RegExMatch(sSelection,sPat,sMatch,Pos+StrLen(sMatch)){ 
    sIssueKey := sMatch1 . "-" . sMatch2
    If InStr(sIssueKeyList,sIssueKey . ";")
        continue
    sIssueKeyList := sIssueKeyList . sIssueKey . ";"
    R4J_Arch_OpenIssue(sIssueKey)
}
return SubStr(sIssueKeyList,1,-1) ; remove ending ;
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
 
R4J_CopyPathJql(sIssueKey:="") {
    sJql := R4J_GetPathJql(sIssueKey)
    If (sJql="") ; Cancel
        return
    Clipboard := sJql
    TrayTipAutoHide("R4J PowerTool","Jql '" . sJql . "' was copied to the clipboard!")
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

R4J_OpenPathJql(sIssueKey:="") {
    If (sIssueKey="") 
        sIssueKey := R4J_GetIssueKey()
    If (sIssueKey="") {
        TrayTip, Error, Jira Issue could not be identified!,,3
        return
    }
    jiraRootUrl := Jira_GetRootUrl()
    IsCloud := Jira_IsCloud(jiraRootUrl)
    If !InStr(sIssueKey,"-") { ; root level 
        If (IsCloud) 
            sJql := "r4jPath in ('" . sIssueKey . "')"
        Else 
            sJql := "issue in requirementsPath('" . sIssueKey . "')"
    } Else {
        sJql := R4J_GetPathJql(sIssueKey)
        sJql := sJql . " or key = " . sIssueKey
    }
    
    searchUrl := jiraRootUrl . "/issues/?jql=" . sJql
    ;MsgBox %sJql%`n%searchUrl% ; DBG
    Atlasy_OpenUrl(searchUrl) 
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
R4J_CopyChildrenJql() { ; @fun_r4j_CopyChildrenJql@
If GetKeyState("Ctrl") and !GetKeyState("Shift") { ; open help
    PowerTools_OpenDoc("r4j_CopyChildrenJql")
    return
} 
issueKeys := Jira_GetIssueKeys()
IsCloud := Jira_IsCloud(jiraRootUrl)
/* 
If (issueKeys.Length() > 1) {
    op := ButtonBox("Jql combine","Choose your logical operator:","OR|AND")
    If ( op = "ButtonBox_Cancel") or ( op = "Timeout")
        return
} 
*/
op := "OR"
for index, Key in issueKeys 
{
    If (IsCloud) 
        Jql_i := "r4jPath in ('" . Key . "')"
    Else 
        Jql_i := "issue in requirementsPath('" . Key . "')"
    If (index=1)
        Jql := Jql_i
    Else
        Jql := Jql . " " . op . " " . Jql_i
} ; end for   
Clipboard := Jql
TrayTipAutoHide("R4J PowerTool","Jql '" . Jql . "' was copied to the clipboard!")
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------


R4J_GetPathJql(sIssueKey:="") {
If GetKeyState("Ctrl") and !GetKeyState("Shift") { ; open help
    PowerTools_OpenDoc("r4j_CopyPathJql")
    return
}

If (sIssueKey="") 
    sIssueKey := R4J_GetIssueKey()
If (sIssueKey="") {
    TrayTip, Error, Jira Issue could not be identified!,,3
    return
}

If !InStr(sIssueKey,"-") ; root level
    sPath := sIssueKey
Else {
    sPath := R4J_GetPath(sIssueKey)
    If !sPath ; empty=error
        return
    ; Select Path : to issue or parent issue
    sPath := ListBox("R4J Get Path","Choose path to use:",sPath . "/" . sIssueKey . "|" . sPath ,1)
    If (sPath="") ; check if cancel
        return
}

IsCloud := Jira_IsCloud() ; check before ListBox to keep browser window active
If (IsCloud) 
    sJql := "r4jPath in ('" . sPath . "')"
Else 
    sJql := "issue in requirementsPath('" . sPath . "')"

sJql := StrReplace(sJql,"'","""")
return sJql
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

; ----------------------------------------------------------------------
R4J_GetApiToken(sUrl:="") {
; Get Auth String for R4J authentification for a specific instance
; sAuth := R4J_GetApiToken(sUrl:="")
; Default Url is Setting JiraRootUrl
    
    
If !RegExMatch(sUrl,"^http")  ; missing root url or default url
    sUrl := Jira_GetRootUrl() . sUrl

; Read Jira PowerTools.ini setting
If !FileExist("PowerTools.ini") {
    PowerTools_ErrDlg("No PowerTools.ini file found!")
    return
}

App := "R4J"
IniRead, Auth, PowerTools.ini,%App%,%App%Auth
If (Auth="ERROR") { 
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


R4J_Get(sPath,apitype:=""){
; R4J API Get for Cloud
; sResponse := R4J_Get(sPath)
; sPath starting with /
; ApiToken taken from Jira or R4J PowerTools.ini settings


; https://easesolutions.atlassian.net/wiki/spaces/R4JC/pages/2250506241/REST+API
WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
R4JCApiRootUrl := "https://eu.r4j-cloud.easesolutions.com"
If RegExMatch(apitype,"^i") {
    sUrl := R4JCApiRootUrl . "/rest/internal/1/"
} Else
    sUrl := R4JCApiRootUrl . "/rest/api/1/"
    
sUrl := sUrl . RegExReplace(sPath,"^/") 
WebRequest.Open("GET", sUrl, false) ; Async=false

JiraRootUrl := Jira_GetRootUrl()
If !InStr(JiraRootUrl,".atlassian.net") {
    TrayTip, Error, Default Jira Url is not Cloud!,,3
	return
}
    
sAuth := Jira_BasicAuth()
If !(sAuth="") {
    sAuth := "Basic " . sAuth
    WebRequest.setRequestHeader("Authorization", sAuth)
    WebRequest.setRequestHeader("X-Atlassian-Base-Url", JiraRootUrl)
} Else {
    sToken := R4J_GetApiToken()
    If !(sToken="")
        WebRequest.setRequestHeader("Authorization", "JWT " . sToken) 
}

If (sAuth="") and (sToken="") {
	TrayTip, Error, R4J or Jira API Authentication is not set in PowerTools.ini!,,3
	return
}

WebRequest.setRequestHeader("Content-Type", "application/json")
WebRequest.Send()     

return WebRequest.responseText
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

R4J_Post(sPath,sBody, apitype:=""){
; sResponse := R4J_Post(sPath,sBody:="")
; sPath without starting /

JiraRootUrl := Jira_GetRootUrl()
If !InStr(JiraRootUrl,".atlassian.net") {
    TrayTip, Error, Default Jira Url is not Cloud!,,3
	return
}

; https://easesolutions.atlassian.net/wiki/spaces/R4JC/pages/2250506241/REST+API

If RegExMatch(apitype,"^i") {
    sUrl := "https://eu.r4j-cloud.easesolutions.com/rest/internal/1/"
} Else
    sUrl := "https://eu.r4j-cloud.easesolutions.com/rest/api/1/"
    
sUrl := sUrl . RegExReplace(sPath,"^/")
WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("POST", sUrl, false) ; Async=false
    
sAuth := Jira_BasicAuth()

If !(sAuth="") {
    sAuth := "Basic " . sAuth
    WebRequest.setRequestHeader("Authorization", sAuth)
    WebRequest.setRequestHeader("X-Atlassian-Base-Url", JiraRootUrl)
} Else {
    sToken := R4J_GetApiToken()
    If !(sToken="")
        WebRequest.setRequestHeader("Authorization", "JWT " . sToken) 
}

If (sAuth="") and (sToken="") {
	TrayTip, Error, R4J or Jira API Authentication is not set in PowerTools.ini!,,3
	return
}

If (sBody = "")
	WebRequest.Send() 
Else
	WebRequest.Send(sBody)       
return WebRequest.responseText
} ; eofun

; ----------------------------------------------------------------------
R4J_WebRequest(sReqType,sPath,sBody:="",apitype:=""){
; Syntax: WebRequest := R4J_WebRequest(sReqType,sUrl,sBody:="")
; Output WebRequest with fields Status and ResponseText
           
JiraRootUrl := Jira_GetRootUrl()
If !InStr(JiraRootUrl,".atlassian.net") {
    TrayTip, Error, Default Jira Url is not Cloud!,,3
	return
}

If RegExMatch(apitype,"^i") {
    sUrl := "https://eu.r4j-cloud.easesolutions.com/rest/internal/1/"
} Else
    sUrl := "https://eu.r4j-cloud.easesolutions.com/rest/api/1/"
sUrl := sUrl . RegExReplace(sPath,"^/") 
WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open(sReqType, sUrl, false) ; Async=false
    
sAuth := Jira_BasicAuth()
If !(sAuth="") {
    sAuth := "Basic " . sAuth
    WebRequest.setRequestHeader("Authorization", sAuth)
    WebRequest.setRequestHeader("X-Atlassian-Base-Url", JiraRootUrl)
} Else {
    sToken := R4J_GetApiToken()
    If !(sToken="")
        WebRequest.setRequestHeader("Authorization", "JWT " . sToken) 
}

; MsgBox Auth:%sAuth%`nToken:%sToken%
If (sAuth="") and (sToken="") {
	TrayTip, Error, R4J or Jira API Authentication is not set in PowerTools.ini!,,3
	return
}

If (sBody = "")
    WebRequest.Send() 
Else {
    WebRequest.setRequestHeader("Content-Type", "application/json")
    WebRequest.Send(sBody) 
}      
WebRequest.WaitForResponse()
stat := WebRequest.Status
If !(stat = "200") and !(stat = "201") {
    resp := WebRequest.responseText
    MsgBox 0x10, Error, Error on %sUrl% - %sReqType%: %stat%`n%resp%
    ;return
}
return WebRequest
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
R4JC_IssueKey2TreeId(IssueKey,projectId:="") {
    issueId := Jira_IssueKey2Id(IssueKey)
    If (projectId="")
        projectId := Jira_ProjectKey2Id(RegExReplace(IssueKey,"\-.*"))
    
    sBody = {"treeProjectId":"%projectId%","jql":"key=%IssueKey%"}
    restPath := "tree/search" 
    Request := R4J_WebRequest("POST",restPath,sBody)
   
    ;JsonObj := Jxon_Load(Request.responseText)
    ;return JsonObj["values"][1][items][1]["id"]
    sPat = "id":(\d*)
    RegExMatch(Request.responseText,sPat,sMatch)
    return sMatch1

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

R4J_GetPath(sIssueKey:="",sProjectKey := "") {
; sPath := R4J_GetPath(sIssueKey,sProjectKey := "")
; Returned Path does not include current Issue 

If (sIssueKey="") 
    sIssueKey := R4J_GetIssueKey()
If (sIssueKey="") {
    TrayTip, Error, Jira Issue could not be identified!,,3
    return
}

If !InStr(sIssueKey,"-") ; root level
    return sIssueKey

If (sProjectKey="") {
    If Browser_IsWinActive() {
        sUrl := Browser_GetUrl()
        If R4J_IsUrl(sUrl)
            sProjectKey := R4J_Url2ProjectKey(sUrl)
    }
}
If (sProjectKey ="")
    sProjectKey := RegExReplace(sIssueKey,"\-.*")

JiraRootUrl := Jira_GetRootUrl()

If InStr(JiraRootUrl,".atlassian.net") { ; Cloud
    restUrl := JiraRootUrl . "/rest/api/latest/issue/" . sIssueKey . "/properties/com.easesolutions.requirements:treeIssuePath"
    sResponse := Jira_Get(restUrl)
    sPat = U)"parents":\[(.*)\]
    If !RegExMatch(sResponse,sPat,sMatch) {
        TrayTip, Error, parents not found in response %sResponse% from %restUrl%!,,3
        return
    }    

    Loop, Parse, sMatch1, `, ; multiple path
        {
            sPat = ^"%sProjectKey%(/|")
            If RegExMatch(A_LoopField,sPat,sMatch) { ; match path starting with projectKey
                sPat = "
                path := StrReplace(A_LoopField, sPat) ; remove quotes
                break
            }  
        }
    If (!path) {
        TrayTip, Error, Can not get path for %sIssueKey%!,,3
        return
    }
} Else { ; server/DC

    sUrl :=  JiraRootUrl . "/rest/com.easesolutions.jira.plugins.requirements/1.0/issue/req-path?jql=key=" . sIssueKey
    sResponse := Jira_Post(sUrl)

    Json := Jxon_Load(sResponse)
    JsonPaths := Json[1]["paths"]

    For i, p in JsonPaths 
    {
        If (p["prjKey"] = sProjectKey) {
            pathArray := p["path"]
            break
        }
    }       

    If !(pathArray) 
        return ; empty
        
    Loop, % pathArray.MaxIndex()   ; concat string array
        path .= pathArray[A_Index]"/"  

    ; Prepend Project name as root
    sUrl := "/rest/api/2/project/" . sProjectKey
    sResponse := Jira_Get(sUrl)
    JsonObj := Jxon_Load(sResponse)
    ProjectName := JsonObj["name"]

    path := ProjectName . "/" path
    path := RegExReplace(path,"/$") ;remove trailing /
}
return path
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
; R4J Coverage Views Functionality based on 
; https://easesolutions.atlassian.net/wiki/spaces/REQ4J/pages/73302077/REST+API+Coverage
; https://easesolutions.atlassian.net/wiki/spaces/REQ4J/pages/1490616345/REST+API+2.0
; -------------------------------------------------------------------------------------------------------------------



R4J_View_Import(viewtype := "c"){
; Import Views from exported json files
If GetKeyState("Ctrl")  {
    PowerTools_OpenDoc("r4j_CV")
    return
}

static S_CV_DestFolder := A_ScriptDir

; Get ProjectKey 
DefProjectKey := R4J_GetProjectDef()
sProjectKey := Jira_InputProjectKey("Copy Views to Project?","Enter target Project Key:",DefProjectKey)
If (sProjectKey == "")
    return

; Get ProjectId from ProjectKey
sProjectId := Jira_ProjectKey2Id(sProjectKey)

; IsPublic
MsgBox, 0x24,Public?, Do you want to import the views as Public?	
isPublic := 0
IfMsgBox Yes 
    isPublic := 1

; Select Json Files to Import
FileSelectFile, JsonFiles, M3, %S_CV_DestFolder%, Filter .json file?, Json File (*.json)
If ErrorLevel 
    Return

; Loop on selected Files
Loop, parse, JsonFiles, `n
{
    if (A_Index = 1) { ; skip first display folder
        S_CV_DestFolder := A_LoopField
        Continue
    }
    ; Read Json File
    JsonFile := S_CV_DestFolder . "\" . A_LoopField
    FileRead, Json, %JsonFile%
    If ErrorLevel {
        TrayTip, Error, Reading file %JsonFile%!,,3
        return
    }
    
    If Jira_IsCloud() & RegExMatch(PathX(JsonFile).File,"U)^(.*)_",sMatch)
        {
            srcProjectId := Jira_ProjectKey2Id(sMatch1)
        }
    R4J_View_Save(viewtype,sProjectId,Json,isPublic,srcProjectId)
        
}

; Open Views in Browser R4J
R4J_OpenProject(sProjectKey,viewtype)

} ; eofun


; -------------------------------------------------------------------------------------------------------------------
R4J_View_Delete(viewtype:="c") { ; @fun_r4j_view_delete@
If GetKeyState("Ctrl")  {
    PowerTools_OpenDoc("r4j_CV") 
    return
}

; Get Jira auth parameters
JiraPassword := Jira_GetPassword()
If (JiraPassword = "") ; cancel
    return	
JiraRootUrl := Jira_GetRootUrl()
If (JiraRootUrl = "")
    return

sProjectKey := Jira_InputProjectKey("Delete Views: Project?","Enter Project Key:")
If (sProjectKey == "")
    return
; Get ProjectId from ProjectKey
sProjectId := Jira_ProjectKey2Id(sProjectKey)

CVArray := R4J_View_GetViews(viewtype,sProjectId,"Select views to delete:")
If (CVArray="")
    return

Loop % CVArray.Length()
{
    viewId := CVArray[A_Index]["id"]
    If InStr(JiraRootUrl,".atlassian.net") { ; cloud
        path := "/coverage/projects/" . sProjectId . "/views/" . viewId
        WebResponse := R4J_WebRequest("DELETE",path)
    } Else { ;server/dc
        sUrl := JiraRootUrl . "/rest/com.easesolutions.jira.plugins.coverage/1.0/filter/deleteFilter?projectId=" . sProjectId . "&filterId=" . CVId
        WebResponse := Jira_WebRequest("DELETE",sUrl,,JiraPassword)
    }
    
    
}

If (WebResponse.Status != 201) {
    txt := WebResponse.responseText
    TrayTip, Error:Save View,%path%`n%txt%,,3  
}
; Open Views in Browser R4J
R4J_OpenProject(sProjectKey,"cov")

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
R4J_View_Save(viewtype, tgtProjectId,JsonView,isPublic:=1,srcProjectId:="") { ; @fun_r4j_view_save@
; Save R4J Coverage or Traceability View
; R4J_View_Save(viewtype:="c",tgtProjectId,JsonView,isPublic:=1,srcProjectId:="")
; Json Response String

JiraRootUrl := Jira_GetRootUrl()

If (srcProjectId = "") and !InStr(JiraRootUrl,".atlassian.net") { ; server/DC projectId is stored in Json View
    ;; Get srcProjectName, srcProjectKey
    ;;; Get srcProjectId
    sPat = "projectId":([^,]*)
    RegExMatch(Json, sPat,sMatch)
    srcProjectId := sMatch1
} 

If (srcProjectId = "") {
    TrayTip, Error:Save View,No source Project Id!,,3  
    return
}

; Transform Json to Body

;; Replace ProjectId (not for cloud)
If !InStr(JiraRootUrl,".atlassian.net") {
    sPat = "projectId":[^,]*
    sRep = "projectId":%tgtProjectId%
    sBody := RegExReplace(JsonView, sPat,sRep)

    ;; isPublic
    sPat = "isPublic":[^}]* ; last parameter
    sRep = "isPublic":%isPublic%
    sBody := RegExReplace(sBody, sPat,sRep)
    
} Else { ; cloud 
    ; replace isPublic by parentId because parent determines the visibility
    parentId := R4JC_View_GetParentId(tgtProjectId, isPublic,viewtype)
    ; replace first id by parentId
    sPat = U)^{"id":.*,
    sRep = {"parentId":%parentId%,
    sBody:=RegExReplace(JsonView,sPat,sRep)

}

; Replace in Jql ProjectKey and RootPath (make filter relative)
;; Get tgtProjectName, tgtProjectKey
sUrl := JiraRootUrl . "/rest/api/latest/project/" . tgtProjectId
sResponse := Jira_Get(sUrl)
JsonObj := Jxon_Load(sResponse)

tgtProjectName := JsonObj["name"]
tgtProjectKey := JsonObj["key"]


sUrl := JiraRootUrl . "/rest/api/latest/project/" . srcProjectId
sResponse := Jira_Get(sUrl)
JsonObj := Jxon_Load(sResponse)
srcProjectName := JsonObj["name"]
srcProjectKey := JsonObj["key"]

;; Replace in Json string srcRootPath -> tgtRootPath & srcProjectKey -> tgtProjectKey whole word only 
; (\b word boundary see https://www.autohotkey.com/docs/v1/misc/RegEx-QuickRef.htm)


If !InStr(JiraRootUrl,".atlassian.net") { ; server|dc
    ; Replace Project Name (not cloud)
    n = i)\b%srcProjectName%\b
    sBody := RegExReplace(sBody,n,tgtProjectName)

    ; Replace Project Folder (before ProjectKey)
    n = \"%srcProjectId%,%srcProjectKey%\"
    r = \"%tgtProjectId%,%tgtProjectKey%\"
    sBody := StrReplace(sBody,n,r)
}

; Replace ProjectKey
n = i)\b%srcProjectKey%\b
sBody := RegExReplace(sBody,n,tgtProjectKey) ; ignore case because project key is lower case in Json view Jql

;MsgBox Body with replaced PK %srcProjectKey%->%tgtProjectKey%:`n%sBody% ; DBG

;; Replace Test ProjectKey
If RegExMatch(srcProjectKey,"^REQTEMPL$")
    srcTestProjectKey := "TESTTEMPL"
Else {
    srcTestProjectKey := RegExReplace(srcProjectKey,"i)^R(.*)","T$1")
}
tgtTestProjectKey := RegExReplace(tgtProjectKey,"i)^R(.*)","T$1")

n = i)\b%srcTestProjectKey%\b
sBody := RegExReplace(sBody,n,tgtTestProjectKey) ; ignore case

; Save Filter via REST API
; REST POST /rest/com.easesolutions.jira.plugins.coverage/1.0/filter/saveFilter
JiraRootUrl := Jira_GetRootUrl()
If InStr(JiraRootUrl,".atlassian.net") { ; Cloud version
    path_create := "/coverage/projects/" . tgtProjectId . "/views"
    WebResponse := R4J_WebRequest("POST",path_create,sBody)
} Else {  ; server/datacenter
    sUrl := JiraRootUrl . "/rest/com.easesolutions.jira.plugins.coverage/1.0/filter/saveFilter" 
    WebResponse := Jira_WebRequest("POST",sUrl,sBody)
}

; Implemented only for Cloud
; If name already exists, aks user if he wants to overwrite or choose another name
NameConflict:
If InStr(JiraRootUrl,".atlassian.net") & (WebResponse.Status = 409 ) { ; improvement: response could return viewId
    JsonObj := Jxon_Load(sBody)
    name := JsonObj["name"]

    OnMessage(0x44, "OnNameConflictMsgBox")
    MsgBox, 0x23, View with same name already exists,Do you want to Overwrite or create with other name?
    OnMessage(0x44, "")
    IfMsgBox Cancel 
        return

    IfMsgBox Yes ; Overwrite
    { 
        viewId := R4JC_View_Name2Id(name, tgtProjectId, JsonObj["parentId"])
        path_update := "/coverage/projects/" . tgtProjectId . "/views/" . viewId
        WebResponse := R4J_WebRequest("PUT",path_update,sBody)
    } Else { ; Rename
        InputBox, name, View Name Conflict, Enter a different View Name,, 250, 125, , , , ,%name%
        If ErrorLevel
            return
        sPat = "name":[^,]*
        sRep = "name":%name%
        sBody := RegExReplace(sBody, sPat,sRep)
        WebResponse := R4J_WebRequest("POST",path_create,sBody)
        If (WebResponse.Status = 409)
            GoTo, NameConflict
    }  
    
}

If (WebResponse.Status != 201) {
    txt := WebResponse.responseText
    TrayTip, Error:Save View,%path%`n%txt%,,3  
}
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
OnNameConflictMsgBox() {
    DetectHiddenWindows, On
    Process, Exist
    If (WinExist("ahk_class #32770 ahk_pid " . ErrorLevel)) {
        ControlSetText Button1, Overwrite
		ControlSetText Button2, Rename
    }
}

; -------------------------------------------------------------------------------------------------------------------
R4J_View_Copy(viewtype:="c",srcProjectId:="", tgtProjectId :="") { ; @fun_r4j_view_copy@
; Copy R4J Coverage View from one source Project to a target Project
; User will be prompted to select views to copy

If (srcProjectId = "") or (tgtProjectId ="") 
    DefProjectKey := R4J_GetProjectDef()
; Get ProjectKey 
If (srcProjectId = "") { 
    sProjectKey := Jira_InputProjectKey("Get Views: Project?","Enter source Project Key:",DefProjectKey)
    If (sProjectKey == "")
        return
    ; Get ProjectId from ProjectKey
    srcProjectId := Jira_ProjectKey2Id(sProjectKey)
}
CVArray := R4J_View_GetViews(viewtype,srcProjectId,"Select views to Copy:")
If (CVArray="")
    return

If (tgtProjectId ="") {
    ; Get Target Project
    tgtProjectKey := Jira_InputProjectKey("Copy Views to Project?","Enter target Project Key:",DefProjectKey)
    If (tgtProjectKey == "")
        return

    ; Get ProjectId from ProjectKey
    tgtProjectId := Jira_ProjectKey2Id(tgtProjectKey)
}

; IsPublic
MsgBox, 0x24,Public?, Do you want to import the views as Public?	
isPublic := 0
IfMsgBox Yes 
    isPublic := 1

for i,cv in CVArray
{
    sResponse := R4J_View_GetConfig(viewtype,srcProjectId, cv["id"])  
    R4J_View_Save(viewtype,tgtProjectId,sResponse,isPublic,srcProjectId)
} ; end for Loop views

; Open Coverage Views in Browser R4J
R4J_OpenProject(tgtProjectKey,"cov")

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
R4J_View_GetConfig(viewtype,projectIdKey, viewId) {
;    Json := R4J_View_GetConfig(viewtype,projectId, viewId) 
; Get Coverage or Traceability View Json Configuration
; return Json string view configuration

JiraRootUrl := Jira_GetRootUrl()
If (viewtype="c")
    viewtype := "coverage"
Else ; "t" traceability view  
    viewtype := "traceability"
If InStr(JiraRootUrl,".atlassian.net") { ; Cloud 
    path := viewtype . "/projects/" . projectIdKey . "/views/" . viewId
    sResponse := R4J_Get(path)
} Else { ; Server/DC
    ; api v2 does not work for coverage -> needs to handle differently with v1
    Switch viewtype
    {
        Case "coverage": 
            
            /*  Old method with v1 api and ProjectId
            sUrl := JiraRootUrl . "/rest/com.easesolutions.jira.plugins." . viewtype . "/1.0/filter/getFilter?projectId=" . projectIdKey . "&filterId=" . viewId
             ; Bug in API response "projectId":0
             s = "projectId":0,"isPublic"
             r = "projectId":%projectIdKey%,"isPublic"
             sResponse := StrReplace(sResponse,s,r)  
             */

            sUrl := JiraRootUrl . "/rest/com.easesolutions.jira.plugins.coverage/2.0/" . "projects/" . projectIdKey . "/" . viewtype . "/views/" . viewId
            sResponse := Jira_Get(sUrl)
           
        Case "traceability":
            sUrl := JiraRootUrl . "/rest/com.easesolutions.jira.plugins.requirements/2.0/" . "projects/" . projectIdKey . "/" . viewtype . "/views/" . viewId
            sResponse := Jira_Get(sUrl)
    }       
}
return sResponse
} ; eofun

; -------------------------------------------------------------------------------------------------------------------


R4J_View_Export(viewtype:="c"){
static S_CV_DestFolder := A_ScriptDir

DefProject := R4J_GetProjectDef()
srcProjectKey := Jira_InputProjectKey("Export Views: Project?","Enter source Project Key:",DefProject)
If (srcProjectKey == "")
    return
IsCloud := Jira_IsCloud()
If IsCloud { ; or (viewtype="c")  ; bug api 2 coverage
    srcProjectId := Jira_ProjectKey2Id(srcProjectKey)
    CVArray := R4J_View_GetViews(viewtype,srcProjectId,"Select Views to Export:")
} Else
    CVArray := R4J_View_GetViews(viewtype,srcProjectKey,"Select Views to Export:")
  
If (CVArray="")
    return

; Select destination folder
FileSelectFolder, DestFolder, *%S_CV_DestFolder%, , Choose your export destination folder:
If ErrorLevel ; Cancel
    return
S_CV_DestFolder := DestFolder

FormatTime sDateStamp ,,yyyyMMddTHHmm

Loop % CVArray.Length()
    {
        CVName := CVArray[A_Index]["name"]
        CVId := CVArray[A_Index]["id"]
        CVName := RegExReplace(CVName,"[><\\/\*\?\|:]","") ; remove special characters for file name \<>|?*:
        If (viewtype="c")
            DestFile := DestFolder . "\" . srcProjectKey . "_" . CVName  . "_CoverageView_" . sDateStamp . ".json"
        Else ; "t" traceability view  
            DestFile := DestFolder . "\" . srcProjectKey . "_" . CVName  . "_TraceabilityView_" . sDateStamp . ".json"
        
        
        If IsCloud ; or (viewtype = "c")
            sResponse := R4J_View_GetConfig(viewtype,srcProjectId, CVId)
        Else
            sResponse := R4J_View_GetConfig(viewtype,srcProjectKey, CVId)
        FileAppend, %sResponse%, %DestFile%	
    }
    Run, %DestFolder%
} ; eofun

; -------------------------------------------------------------------------------------------------------------------


R4J_View_GetViews(viewtype:="c",sProjectIdKey:="",sPrompt:="Select Views:") { ; @fun_r4j_view_getviews@
; Get Coverage Views of a Project
; 	CVArray := R4J_View_GetViews(viewtype:="c",sProjectIdKey:="",sPrompt:="Select Views:")
; For server, sProjectIdKey shall be the project Key for traceability and project Id for coverage
; For Cloud, sProjectIdKey shall be the project Id
; User is prompted to enter Project Key if ProjectIdKey is not provided
; Output
;   CVArray: array of objects with values "id" and "name"  e.g. id of first view is CVArray[1][id]
;   empty if no views defined

If GetKeyState("Ctrl")  {
    PowerTools_OpenDoc("r4j_ViewGet")
    return
}
JiraRootUrl := Jira_GetRootUrl() 

; Get ProjectKey 
If (sProjectIdKey == "") {
    DefProjectKey := R4J_GetProjectDef()
    sProjectIdKey := Jira_InputProjectKey("Get Views: Project?","Enter source Project Key:",DefProjectKey)
    If (sProjectIdKey == "")
        return
    If InStr(JiraRootUrl,".atlassian.net") ;or (viewtype = "coverage") ; Get ProjectId from ProjectKey. Server coverage requires v1 with projectId
        sProjectIdKey := Jira_ProjectKey2Id(sProjectKey)
} Else {
    If InStr(JiraRootUrl,".atlassian.net") ;or (viewtype = "coverage")
        sProjectIdKey := Jira_ProjectKey2Id(sProjectIdKey)
}

; Get Coverage View Ids

If (viewtype="c")
    viewtype := "coverage"
Else ; "t" traceability view  
    viewtype := "traceability"
If InStr(JiraRootUrl,".atlassian.net") { ; Cloud
    viewFolderId := R4JC_View_GetParentId(sProjectIdKey)
    
    path := viewtype . "/projects/" . sProjectIdKey . "/views/" . viewFolderId . "/children"
    sResponse := R4J_Get(path) 
    ;MsgBox %path%`n%sResponse% ; DBG
    Json := Jxon_Load(sResponse)
    JsonViews := Json["values"]
} Else { ; server/dc
    
    
    ; api v2 does not work for coverage -> needs to handle differently with v1
    Switch viewtype
    {
        Case "coverage": 
            ;sUrl := JiraRootUrl . "/rest/com.easesolutions.jira.plugins." . viewtype . "/1.0/filter/getFilters?projectId=" . sProjectIdKey
            ;sResponse := Jira_Get(sUrl)
            ;MsgBox Views: %sResponse% ; DBG
            ;Json := Jxon_Load(sResponse)
            ;JsonViews := Json["publicFilters"]

            sUrl := JiraRootUrl . "/rest/com.easesolutions.jira.plugins.coverage/2.0/projects/" . sProjectIdKey . "/" . viewtype . "/views"
            sResponse := Jira_Get(sUrl)
            JsonViews := Jxon_Load(sResponse)
        Case "traceability":
            sUrl := JiraRootUrl . "/rest/com.easesolutions.jira.plugins.requirements/2.0/projects/" . sProjectIdKey . "/" . viewtype . "/views"
            sResponse := Jira_Get(sUrl)
            JsonViews := Jxon_Load(sResponse)
    }
     
    
}
;MsgBox % Jxon_Dump(JsonViews)
If (ObjCount(JsonViews) = 0) {
    TrayTip, Warning, No public views defined!,,2
    return
}

static S_CV_Ids := []

CVAllArray := []

For i,v in JsonViews 
{
    sel := HasVal(S_CV_Ids,v["id"]) 
    CVAllArray.Push({"name":v["name"],"id":v["id"],"select":sel})
}

CVArray := ListView("Views Selector",sPrompt,CVAllArray,"View")
If (CVArray == "ListViewCancel")
    return

; Store selected View Ids
S_CV_Ids := []
Loop , % CVArray.Length()
{
    S_CV_Ids[A_Index] := CVArray[A_Index]["id"]
}

return CVArray

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
R4JC_View_GetParentId(projectId, isPublic := true,viewtype:="c") {
; Get Public Views Folder Id (Id change for each project)
If (viewtype="c")
    path := "/coverage/"
Else ; "t" traceability view  
    path := "/traceability/"
path := path . "projects/" . projectId . "/views/-1/children"
sResponse := R4J_Get(path)
Json := Jxon_Load(sResponse)
JsonViews := Json["values"]
For i, v in JsonViews 
{   
    If (v["name"] = "Public Views") & isPublic
        return v["id"]
    If (v["name"] = "Private Views") & !isPublic
        return v["id"]
}

} ; eof
; -------------------------------------------------------------------------------------------------------------------
R4JC_View_Name2Id(name, projectId, parentId,viewtype:="c") {
; Get ViewId from View Name
; viewId := R4JC_View_Name2Id(name, projectId, parentId,viewtype)
; viewtype "c" for coverage (default)
;          "t" for traceability

; called by: R4J_View_Save
    If (viewtype="c")
        path := "/coverage/"
    Else ; "t" traceability view  
        path := "/traceability/"
    path := path . "projects/" . projectId . "/views/" . parentId . "/children"
    sResponse := R4J_Get(path) 
    ;MsgBox %path%`n%sResponse% ; DBG
    Json := Jxon_Load(sResponse)
    JsonViews := Json["values"]

    For i,v in JsonViews 
    {
        If (v["name"] = name) 
            return v["id"]
    }
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
R4J_View_Migrate() {

; replace requirementPath Jql

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

R4J_CV_Report(Debug:=False) {
If GetKeyState("Ctrl")  {
    PowerTools_OpenDoc("r4j_CVReport")
    return
}

JiraRootUrl := Jira_GetRootUrl()
If (JiraRootUrl = "")
    return
If Jira_IsCloud(JiraRootUrl) {
    PowerTools_ErrDlg("Coverage Report is not implemented for Cloud version!")
    return
}

static S_CR_DestFolder := A_ScriptDir
static S_JsonFile := "r4j"

; Set environment variables
JiraPassword := Jira_GetPassword()
If (JiraPassword = "") ; cancel
    return	
EnvSet, JIRA_PASSWORD, %JiraPassword%


EnvSet, JIRA_HOST_URL, %JiraRootUrl%

JiraUserName := Jira_GetUserName()
If (JiraUserName = "")
    return
EnvSet, JIRA_USER, %JiraUserName%

; Get ProjectKey 
sProjectKey := Jira_InputProjectKey("Coverage Report: Project?","Enter Project Key:")
If (sProjectKey == "")
    return

; Get ProjectId from ProjectKey
sProjectId := Jira_ProjectKey2Id(sProjectKey)
EnvSet, JIRA_PROJECT_ID, %sProjectId%

; Get Coverage View Ids
CVArray := R4J_View_GetViews(sProjectId,"Select Views to be used in the Report:")
If (CVArray=="") ; user cancelled
    return

; Select destination folder
FileSelectFolder, DestFolder, *%S_CR_DestFolder%, , Choose your export destination folder:
If ErrorLevel ; Cancel
    return
S_CR_DestFolder := DestFolder
; Select fields.json file
FileSelectFile, JsonFile, 1, %S_JsonFile%, Fields.json file?, Json File (*.json)
If !ErrorLevel
    S_JsonFile := JsonFile

FormatTime sDateStamp ,,yyyyMMddTHHmm
ExeFile :=  A_ScriptDir . "\r4j\coverage_exporter.exe"
sCmd0 := ExeFile . " generate-report"

If (JsonFile !="")
    sCmd0 := sCmd0 . " --fields-config """ . JsonFile . """"

Loop % CVArray.Length()
{
    CVName := CVArray[A_Index]["name"]
    DestFile := DestFolder . "\" . sProjectKey . "_" . CVName  . "_Coverage_Report_" . sDateStamp . ".xlsx"
    CVArray[A_Index]["file"] := DestFile

    EnvSet, R4J_CV_ID, % CVArray[A_Index]["id"]

    sCmd := sCmd0 . " -f """ . DestFile . """"
    If (Debug) {
        sCmd = %ComSpec% /c echo %sCmd% && %sCmd% && pause ; add pause to display log
        RunWait, %sCmd%
    } Else
        Run %sCmd%
}

; Merge Files?
MsgBox, 0x24,Merge Reports?, Do you want to merge all Excel files into one file with one report per Sheet?	
IfMsgBox Yes 
{
    DestFile := DestFolder . "\" . sProjectKey .  "_Coverage_Report_" . sDateStamp . ".xlsx" 

    oExcel := ComObjCreate("Excel.Application") ;handle
    oExcel.SheetsInNewWorkbook := 1
    xlWorkbook := oExcel.Workbooks.Add ;add a new workbook
    infoSheet := xlWorkbook.Sheets(1)
    infoSheet.Name := "Info"
    oExcel.Visible := False 

    infoSheet.Range("A1").Value := "Creator"
    infoSheet.Range("B1").Value := JiraUserName
    infoSheet.Range("A2").Value := "Date"
    infoSheet.Range("B2").Value := sDateStamp

    Loop % CVArray.Length()
    {
        CVName := CVArray[A_Index]["name"]
        ; truncate Sheet name to 31 characters
        SheetName := SubStr(CVName,1,31)
        srcFile := CVArray[A_Index]["file"]
        nRow := A_Index + 3

        CVLink := JiraRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/coverage?prj=" . sProjectKey . "#" . CVArray[A_Index]["id"]
        sFormula = =HYPERLINK("%CVLink%","View in R4J") 
        infoSheet.Range("A" . nRow).Formula := sFormula ; escape quotes

        If FileExist(srcFile) {
            ; Copy Coverage Sheet
            srcWorkbook:=oExcel.Workbooks.Open(srcFile)
            srcWorkbook.Sheets("Coverage").Copy(,xlWorkbook.Sheets(xlWorkbook.Sheets.Count)) ; after last sheet
            reportSheet := xlWorkbook.Sheets(xlWorkbook.Sheets.Count)
            reportSheet.Name := SheetName

            ; Fine format report sheet with Filter and Freeze
            reportSheet.Rows(5).AutoFilter()
            reportSheet.Rows(6).Select
            oExcel.ActiveWindow.FreezePanes := True

            ; Add values to info sheet
            sFormula = =HYPERLINK("#'%SheetName%'!A1","%CVName%") ; https://www.ablebits.com/office-addins-blog/excel-insert-hyperlink/ requires # '' for blank and link to cell
            infoSheet.Range("B" . nRow).Formula := sFormula ; escape quotes
            sFormula = =SUBSTITUTE(CONCAT('%SheetName%'!4:4),"Coverage:","")
            infoSheet.Range("C" . nRow).Formula := sFormula 
            srcWorkbook.Close(False)
        } Else {
            ; Fill Info sheet
            infoSheet.Range("B" . nRow).Value := CVName
            infoSheet.Range("C" . nRow).Value := "FAILED"
        }
        infoSheet.Columns(2).AutoFit()
        
    }
    xlWorkbook.SaveAs(DestFile,xlWorkbookDefault := 51)
    xlWorkbook.Sheets("Info").Activate()
    oExcel.Visible := True
    ;xlWorkbook.Close(False)

} ; end IfMsgBox



} ; eofun

; -------------------------------------------------------------------------------------------------------------------
R4J_Remove_Deleted(sProjectKey){
; TODO

; Filter issues with Resolution in ("Deleted")

; Loop on issues, remove from tree
Loop 
{
    R4J_Remove_Issue(sIssueKey, sProjectKey)
} ; end Loop

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

R4J_Remove_Issue(sIssueKey,sProjectKey) {
; TODO

; Identify Parent type 
;; Get Tree Info
; GET /rest/com.easesolutions.jira.plugins.requirements/1.0/tree/%sProjectKey%

sUrl:= "/rest/com.easesolutions.jira.plugins.requirements/1.0/tree/" . sProjectKey
sResponse := Jira_Get(sUrl)
MsgBox % sResponse ; DBG


;/rest/com.easesolutions.jira.plugins.requirements/1.0/issue/req-path?jql=issuekey=
; If Parent is an issue= path contains only "prjKey":%sProjectKey%","path":["folderName","parentKey"]
; [{"issueKey":"RDMO-6","paths":[{"prjKey":"RDMO","path":["System Requirements","RDMO-5"]}]}]


; DELETE /rest/com.easesolutions.jira.plugins.requirements/1.0/child-req/%sProjectKey%/%sParentKey%?childKey=%sIssueKey%

; If Parent is a folder = path contains only "path"["folderName"]
;; Get folder Id
; GET /rest/com.easesolutions.jira.plugins.requirements/1.0/tree/%sProjectKey%/folders
;sMatch= "name":"Folder 1.1","id":10}

; DELETE /rest/com.easesolutions.jira.plugins.requirements/1.0/tree/%sProjectKey%/folderissue/%sFolderId%?issueKey=%sIssueKey%

; If Root Issue
; DELETE /rest/com.easesolutions.jira.plugins.requirements/1.0/tree/%sProjectKey%/rootissue?issuekey=


} ; eofun


; -------------------------------------------------------------------------------------------------------------------
R4J_Redirect(sUrl,tgtRootUrl:="") {
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


; coverage views
; $rootc/plugins/servlet/ac/com.easesolutions.jira.plugins.requirements/requirements-page-jira?project.id=16330#!coverage -> $roots/plugins/servlet/com.easesolutions.jira.plugins.requirements/coverage?prj=REQTEMPL
If (!InStr(tgtUrl,".atlassian.net") & InStr(sUrl,".atlassian.net")) { ; cloud to server
    If RegexMatch(sUrl,"project\.id=(\d*)#!coverage",sMatch) {
        prjKey := Jira_ProjectId2Key(sMatch1,srcRootUrl)
        tgtUrl := tgtRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/coverage?prj=" . prjKey
        return tgtUrl
    }
    
} Else If (InStr(tgtUrl,".atlassian.net") & !InStr(sUrl,".atlassian.net")) { ; server to cloud
    If RegexMatch(sUrl,"/coverage\?prj=([A-Z\d]*)",sMatch) {
        prjId := Jira_ProjectKey2Id(sMatch1,tgtRootUrl)
        tgtUrl := tgtRootUrl . "/plugins/servlet/ac/com.easesolutions.jira.plugins.requirements/requirements-page-jira?project.id=" . prjId . "#!coverage"
        return tgtUrl
    }
}

; traceability
; $roots/plugins/servlet/com.easesolutions.jira.plugins.requirements/tracematrix?prj=RTHB
; /tracematrix?prj=RTHB#14&CRS%20FTTI%20-%3E%20SRS%20FTTI&true
; $rootc/plugins/servlet/ac/com.easesolutions.jira.plugins.requirements/requirements-page-jira?project.id=16310#!traceability?viewId=704
If (!InStr(tgtUrl,".atlassian.net") & InStr(sUrl,".atlassian.net")) { ; cloud to server
    If RegexMatch(sUrl,"project\.id=(\d*)#!traceability",sMatch) {
        prjKey := Jira_ProjectId2Key(sMatch1,srcRootUrl)
        tgtUrl := tgtRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/tracematrix?prj=" . prjKey
        return tgtUrl
    }
    
} Else If (InStr(tgtUrl,".atlassian.net") & !InStr(sUrl,".atlassian.net")) { ; server to cloud
    If RegexMatch(sUrl,"/tracematrix\?prj=([A-Z\d]*)",sMatch) {
        prjId := Jira_ProjectKey2Id(sMatch1,tgtRootUrl)
        tgtUrl := tgtRootUrl . "/plugins/servlet/ac/com.easesolutions.jira.plugins.requirements/requirements-page-jira?project.id=" . prjId . "#!traceability"
        return tgtUrl
    }
}

; tree
; $rootc/plugins/servlet/ac/com.easesolutions.jira.plugins.requirements/requirements-page-jira?project.id=16310#!tree?issueKey=RTHB-3348
; $roots/servlet/com.easesolutions.jira.plugins.requirements/project?detail=RTHB&issueKey=RTHB-2904
If (!InStr(tgtUrl,".atlassian.net") & InStr(sUrl,".atlassian.net")) { ; cloud to server
    If RegexMatch(sUrl,"project\.id=(\d*)#",sMatch) {
        prjKey := Jira_ProjectId2Key(sMatch1,srcRootUrl)
        tgtUrl := tgtRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/project?detail=" . prjKey
    }
    If RegexMatch(sUrl,"\?issueKey=(.*)",sMatch)
        tgtUrl := tgtUrl . "&issueKey=" . sMatch1
    
} Else If (InStr(tgtUrl,".atlassian.net") & !InStr(sUrl,".atlassian.net")) { ; server to cloud
    If RegexMatch(sUrl,"detail=([^&]*)",sMatch) {
        prjId := Jira_ProjectKey2Id(sMatch1,tgtRootUrl)
        tgtUrl := tgtRootUrl . "/plugins/servlet/ac/com.easesolutions.jira.plugins.requirements/requirements-page-jira?project.id=" . prjId . "#!tree"
    }
    If RegexMatch(sUrl,"&issueKey=(.*)",sMatch)
        tgtUrl := tgtUrl . "?issueKey=" . sMatch1
}

return tgtUrl

} ; eofun
; -------------------------------------------------------------------------------------------------------------------


R4J_Migrate_Jql(Jql) {
; Convert Jql requirementsPath to r4jPath 

; (issue in requirementsPath("Triathlon HV BMS_Req/CRS - Customer Requirement Specification/rthb-24/rthb-25") OR  key = RTHB-2861) and issuetype = "Customer Requirement" and "Safety Level" != QM and key != rthb-1704 and key != RTHB-26 and key != RTHB-2861

keyPat := "i)([A-Z\d]{3,}\-[\d]{1,})"
NewJql := Jql
; Loop on issue keys
sPat = U)issue (?:not |)in requirementsPath\("(.*)"\)
Pos = 1 
While Pos := RegExMatch(Jql,sPat,matchPath,Pos+StrLen(matchPath))
{ 
    If InStr(matchPath," not ")
        sNot := "not " 
    Else
        sNot :=""
    NewR4jPath := matchPath1
    Pos2:= 0
    NewPos2 :=0
    While Pos2 := InStr(matchPath1 . "/","/",,Pos2 + 1)
    {
        subPath := SubStr(matchPath1,1,Pos2 - 1)
        ;MsgBox %matchPath1%`n%subPath% ; DBG
        NewPos2 := InStr(NewR4jPath,"/",,NewPos2 + 1)
        subNewPath := SubStr(NewR4jPath,1,NewPos2 - 1)
        If !(InStr(subPath,"/")) { ; root
            prjName := Jira_ProjectName2Key(subPath)
            NewR4jPath := RegExReplace(NewR4jPath,"^" . subPath, prjName)
        } Else {
            RegExMatch(subPath,"/([^/]*)$",nextMatch)
            ; Get key from summary
            ;MsgBox %subPath%`n%matchPath1%`n%nextMatch1% ; DBG
            If !RegExMatch(nextMatch1,keyPat) { ; IsKey
                parNewPath := RegExReplace(subNewPath,"/([^/]*)$")
                rest = /rest/api/latest/search?jql=r4jPath in ("%parNewPath%") and summary~"\"%nextMatch1%\""
                sResponse := Jira_Get(rest)
                sPat = U)"key":"(.*)"
                If !RegExMatch(sResponse,sPat,keyMatch) {
                    MsgBox %rest% could not find any issue!`n%sResponse%
                    break
                }
                    
                NewR4jPath := RegExReplace(NewR4jPath, "/" . nextMatch1 . "$","/" . keyMatch1) ; ending
                NewR4jPath := RegExReplace(NewR4jPath, "/" . nextMatch1 . "/","/" . keyMatch1 . "/")
            }
        }
    }
    NewR4jPathJql = r4jPath %sNot%in ("%NewR4jPath%")
    NewJql := StrReplace(NewJql, matchPath, NewR4jPathJql)
}
return NewJql
} ; eofun


R4J_Copy_Jql(Jql,prjKey) {
; Copy Jql r4jPath from one project to another

; (issue in requirementsPath("Triathlon HV BMS_Req/CRS - Customer Requirement Specification/rthb-24/rthb-25") OR  key = RTHB-2861) and issuetype = "Customer Requirement" and "Safety Level" != QM and key != rthb-1704 and key != RTHB-26 and key != RTHB-2861

    keyPat := "i)([A-Z\d]{3,}\-[\d]{1,})"
    NewJql := Jql
    ; Loop on issue keys
    sPat = Ui)r4jPath in \("(.*)"\)|r4jPath\s?=\s?"(.*)"
    Pos = 1 
    While Pos := RegExMatch(Jql,sPat,matchPath,Pos+StrLen(matchPath))
    { 
        NewR4jPath := matchPath1
        Pos2:= 0
        NewPos2 :=0
        While Pos2 := InStr(matchPath1 . "/","/",,Pos2 + 1)
        {
            subPath := SubStr(matchPath1,1,Pos2 - 1)
            ;MsgBox %matchPath1%`n%subPath% ; DBG
            NewPos2 := InStr(NewR4jPath,"/",,NewPos2 + 1)
            subNewPath := SubStr(NewR4jPath,1,NewPos2 - 1)
            If !(InStr(subPath,"/")) { ; root
                NewR4jPath := RegExReplace(NewR4jPath,"^" . subPath, prjKey)
            } Else {
                RegExMatch(subPath,"/([^/]*)$",nextMatch)
                ; Get summary from src key
                summary := Jira_GetIssueField(nextMatch1,"summary") 
                ; get tgt Key from path and summary
                parNewPath := RegExReplace(subNewPath,"/([^/]*)$")
                rest = /rest/api/latest/search?jql=r4jPath in ("%parNewPath%") and summary~"\"%summary%\""
                sResponse := Jira_Get(rest)
                sPat = U)"key":"(.*)"
                If !RegExMatch(sResponse,sPat,keyMatch) {
                    MsgBox %rest% could not find any issue!`n%sResponse%
                    break
                }
                    
                NewR4jPath := RegExReplace(NewR4jPath, "/" . nextMatch1 . "$","/" . keyMatch1) ; ending
                NewR4jPath := RegExReplace(NewR4jPath, "/" . nextMatch1 . "/","/" . keyMatch1 . "/")

            }
        }
        NewR4jPathJql = r4jPath in ("%NewR4jPath%")
        NewJql := StrReplace(NewJql, matchPath, NewR4jPathJql)
    }
    return NewJql
} ; eofun

