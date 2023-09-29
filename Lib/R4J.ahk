#Include <Jira>
#Include <Atlasy>

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
    If RegExMatch(sUrl,StrReplace(sRootUrl,".","\.") . "/browse/([A-Z]*-\d*)",sIssueKey) 
    {
        R4J_OpenIssue(sIssueKey1)
        return
    } ; tree view 
    Else If RegExMatch(sUrl,StrReplace(sRootUrl,".","\.") . "/plugins/servlet/com\.easesolutions\.jira\.plugins\.requirements/project\?detail=[A-Z]*" . "&issueKey=(.*)" ,sMatch) 
    {
        sUrl := sRootUrl . "/browse/" . sMatch1
        Run, %sUrl% 
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


R4J_OpenProject(sProjectKey) {
    ; Open project R4J tree
    sJiraRootUrl := Jira_IssueKey2RootUrl(sProjectKey)
    If InStr(sJiraRootUrl,".atlassian.net") { ; cloud version
        sProjectId := Jira_ProjectKey2Id(sProjectKey)
        sUrl := sJiraRootUrl . "/plugins/servlet/ac/com.easesolutions.jira.plugins.requirements/requirements-page-jira?project.id=" . sProjectId 
    } Else
        sUrl := sJiraRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/project?detail=" . sProjectKey . "#!tree"
    Atlasy_OpenUrl(sUrl)
} ; eofun


; -------------------------------------------------------------------------------------------------------------------
R4J_IsUrl(sUrl){
return  InStr(sUrl,"/com.easesolutions.jira.plugins.requirements/") 
} ;eofun
; -------------------------------------------------------------------------------------------------------------------

R4J_Url2IssueKey(sUrl){
; /com.easesolutions.jira.plugins.requirements/project?detail=RDMO&issueKey=RDMO-14
sKeyPat := "([A-Z]{3,}\-\d{1,})"
if RegExMatch(sUrl,"&issueKey=" . sKeyPat,sMatch) ; url ending with Issue Key
	return sMatch1
sUrl := RegExReplace(sUrl,"\?.*$","") ; remove optional arguments after ? in url
if RegExMatch(sUrl,"/" . sKeyPat . "$",sMatch) ; url ending with Issue Key
    return sMatch1
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

R4J_CopyPathJql(sIssueKey,sProjectKey := "") {
If GetKeyState("Ctrl") and !GetKeyState("Shift") { ; open help
    PowerTools_OpenDoc("r4j_CopyPathJql")
    return
}
sPath := R4J_GetPath(sIssueKey,sProjectKey)
sPath := ListBox("R4J Copy Path","Choose path to use:",sPath . "/" . sIssueKey . "|" . sPath ,1)
If (sPath="") ; check if cancel
    return

; Select Path : to issue or parent issue
sJql := "issue in requirementsPath('" . sPath . "')"
sJql := StrReplace(sJql,"'","""")
Clipboard := sJql
TrayTipAutoHide("R4J PowerTool","Jql '" . sJql . "' was copied to the clipboard!")

} ; eofun

; -------------------------------------------------------------------------------------------------------------------

R4J_GetPath(sIssueKey,sProjectKey := "") {
; sPath := R4J_GetPath(sIssueKey,sProjectKey := "")

If (sProjectKey ="") {
    sProjectKey := RegExReplace(sIssueKey,"\-.*")
}

sRestUrl := Jira_GetRootUrl() . "/rest/com.easesolutions.jira.plugins.requirements/1.0/issue/req-path?jql=key=" . sIssueKey
sResponse := Jira_Post(sRestUrl)

Json := Jxon_Load(sResponse)
JsonPaths := Json[1]["paths"]

For i, p in JsonPaths 
{
    If (p["prjKey"] = sProjectKey) {
        pathArray := p["path"]
        Break
    }
}       

If !(pathArray) 
    return ; empty
    
Loop, % pathArray.MaxIndex()   ; concat string array
{
    path .= pathArray[A_Index]"/"
}    

; Prepend Project name as root
sUrl := "/rest/api/2/project/" . sProjectKey
sResponse := Jira_Get(sUrl)
JsonObj := Jxon_Load(sResponse)
ProjectName := JsonObj["name"]

path := ProjectName . "/" path
path := RegExReplace(path,"/$") ;remove trailing /
return path

} ; eofun


; -------------------------------------------------------------------------------------------------------------------
; R4J Coverage Views Functionality based on https://easesolutions.atlassian.net/wiki/spaces/REQ4J/pages/73302077/REST+API+Coverage
; -------------------------------------------------------------------------------------------------------------------

#Include Jxon.ahk

R4J_CV_Import(){
If GetKeyState("Ctrl")  {
    PowerTools_OpenDoc("r4j_CV")
    return
}

static S_CV_DestFolder := A_ScriptDir . "/cv"

; Get ProjectKey 
sProjectKey := Jira_InputProjectKey("Copy Views to Project?","Enter target Project Key:")
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

    R4J_CV_Save(sProjectId,Json,isPublic)
        
}

; Open Coverage Views in Browser R4J
sUrl := JiraRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/coverage?prj=" . sProjectKey 
Run, "%sUrl%"

} ; eofun


; -------------------------------------------------------------------------------------------------------------------
R4J_CV_Delete(){
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

sProjectKey := Jira_InputProjectKey("Delete Coverage Views: Project?","Enter Project Key:")
If (sProjectKey == "")
    return
; Get ProjectId from ProjectKey
sProjectId := Jira_ProjectKey2Id(sProjectKey)

CVArray := R4J_CV_GetViews(sProjectId,"Select views to delete:")
If (CVArray="")
    return

Loop % CVArray.Length()
{
    CVId := CVArray[A_Index]["id"]
    sUrl := JiraRootUrl . "/rest/com.easesolutions.jira.plugins.coverage/1.0/filter/deleteFilter?projectId=" . sProjectId . "&filterId=" . CVId
    Jira_WebRequest("DELETE",sUrl,,JiraPassword)
}

; Open Coverage Views in Browser R4J
sUrl := JiraRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/coverage?prj=" . sProjectKey 
Run, "%sUrl%"

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
R4J_CV_Save(tgtProjectId,Json,isPublic:=1) {

; Transform Json to Body
;; Replace ProjectId 
sPat = "projectId":[^,]*
sRep = "projectId":%tgtProjectId%
sBody := RegExReplace(Json, sPat,sRep)

;; isPublic 
sPat = "isPublic":[^}]* ; last parameter
sRep = "isPublic":%isPublic%
sBody := RegExReplace(sBody, sPat,sRep)


; Replace in Jql ProjectKey and RootPath (make filter relative)
;; Get tgtProjectName, tgtProjectKey
sUrl := "/rest/api/2/project/" . tgtProjectId
sResponse := Jira_Get(sUrl)
JsonObj := Jxon_Load(sResponse)

tgtProjectName := JsonObj["name"]
tgtProjectKey := JsonObj["key"]


;; Get srcProjectName, srcProjectKey
;;; Get srcProjectId
sPat = "projectId":([^,]*)
RegExMatch(Json, sPat,sMatch)
srcProjectId := sMatch1

sUrl := "/rest/api/2/project/" . srcProjectId
sResponse := Jira_Get(sUrl)
JsonObj := Jxon_Load(sResponse)
srcProjectName := JsonObj.name
srcProjectKey := JsonObj.key

;; Replace in Json string srcRootPath -> tgtRootPath & srcProjectKey -> tgtProjectKey whole word only 
; (\b word boundary see https://www.autohotkey.com/docs/v1/misc/RegEx-QuickRef.htm)
n = i)\b%srcProjectName%\b
sBody := RegExReplace(sBody,n,tgtProjectName)


;; Replace Project Folder (before ProjectKey)
n = \"%srcProjectId%,%srcProjectKey%\"
r = \"%tgtProjectId%,%tgtProjectKey%\"
sBody := StrReplace(sBody,n,r)

;; Replace ProjectKey
n = i)\b%srcProjectKey%\b
sBody := RegExReplace(sBody,n,tgtProjectKey) ; ignore case

;; Replace Test ProjectKey
If RegExMatch(srcProjectKey,"^REQTEMPL$")
    srcTestProjectKey := "TESTTEMPL"
Else {
    srcTestProjectKey := RegExReplace(srcProjectKey,"^R(.*)","T$1")
}
tgtTestProjectKey := RegExReplace(tgtProjectKey,"^R(.*)","T$1")
n = i)\b%srcTestProjectKey%\b
sBody := RegExReplace(sBody,n,tgtTestProjectKey) ; ignore case


; Save Filter via REST API
; REST POST /rest/com.easesolutions.jira.plugins.coverage/1.0/filter/saveFilter
sUrl := "/rest/com.easesolutions.jira.plugins.coverage/1.0/filter/saveFilter" 
WebResponse := Jira_WebRequest("POST",sUrl,sBody)

If WebResponse.Status <> 201 {
    txt := WebResponse.responseText
    TrayTip, Error:SaveFilter,%txt%,,3  ; DBG
}
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
R4J_CV_Copy(){

; Get Jira auth parameters
JiraPassword := Jira_GetPassword()
If (JiraPassword = "") ; cancel
    return	
JiraRootUrl := Jira_GetRootUrl()
If (JiraRootUrl = "")
    return


; Get ProjectKey 
srcProjectKey := Jira_InputProjectKey("Get Coverage Views: Project?","Enter source Project Key:")
If (srcProjectKey == "")
    return
; Get ProjectId from ProjectKey
srcProjectId := Jira_ProjectKey2Id(srcProjectKey)


CVArray := R4J_CV_GetViews(srcProjectId,"Select views to Copy:")
If (CVArray="")
    return

; Get Target Project
tgtProjectKey := Jira_InputProjectKey("Copy Views to Project?","Enter target Project Key:")
If (tgtProjectKey == "")
    return

; Get ProjectId from ProjectKey
tgtProjectId := Jira_ProjectKey2Id(tgtProjectKey)

; IsPublic
MsgBox, 0x24,Public?, Do you want to import the views as Public?	
isPublic := 0
IfMsgBox Yes 
    isPublic := 1

Loop % CVArray.Length()
{
    CVId := CVArray[A_Index]["id"]
    
    sUrl := JiraRootUrl . "/rest/com.easesolutions.jira.plugins.coverage/1.0/filter/getFilter?projectId=" . srcProjectId . "&filterId=" . CVId
    sResponse := Jira_Get(sUrl)
    ; Bug in API response "projectId":0
    s = "projectId":0,"isPublic"
    r = "projectId":%srcProjectId%,"isPublic"
    sResponse := StrReplace(sResponse,s,r)
    
    R4J_CV_Save(tgtProjectId,sResponse,isPublic)

} ; end Loop views

; Open Coverage Views in Browser R4J
sUrl := JiraRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/coverage?prj=" . tgtProjectKey 
Run, "%sUrl%"

} ; eofun

; -------------------------------------------------------------------------------------------------------------------

R4J_CV_Export(){
static S_CV_DestFolder := A_ScriptDir

srcProjectKey := Jira_InputProjectKey("Get Coverage Views: Project?","Enter source Project Key:")
If (srcProjectKey == "")
    return
; Get ProjectId from ProjectKey
srcProjectId := Jira_ProjectKey2Id(srcProjectKey)

CVArray := R4J_CV_GetViews(srcProjectId,"Select Views to Export:")
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
        DestFile := DestFolder . "\" . srcProjectKey . "_" . CVName  . "_Coverage_Filter_" . sDateStamp . ".json"
        
        sUrl := JiraRootUrl . "/rest/com.easesolutions.jira.plugins.coverage/1.0/filter/getFilter?projectId=" . srcProjectId . "&filterId=" . CVId
        sResponse := Jira_Get(sUrl,JiraPassword)
        ; Bug in API response "projectId":0
        s = "projectId":0,"isPublic"
        r = "projectId":%srcProjectId%,"isPublic"
        sResponse := StrReplace(sResponse,s,r)
        FileAppend, %sResponse%, %DestFile%	
    }
} ; eofun

; -------------------------------------------------------------------------------------------------------------------


R4J_CV_GetViews(sProjectId,sPrompt:="Select Views:"){
; Get Coverage Views of a Project
; 	CVArray := R4J_CV_GetViews(ProjectId*,sPrompt*)
; User is prompted to enter Project Key if ProjectId is not provided
; Output
;   CVArray: array of objects with values "id" and "name"

; Call: R4J_CV_GetFilters
If GetKeyState("Ctrl")  {
    PowerTools_OpenDoc("r4j_CVGet")
    return
}

; Get ProjectKey 
If (sProjectId == "") {
    sProjectKey := Jira_InputProjectKey("Get Coverage Views: Project?","Enter source Project Key:")
    If (sProjectKey == "")
        return
    ; Get ProjectId from ProjectKey
    sProjectId := Jira_ProjectKey2Id(sProjectKey)
}

; Get Coverage View Ids
return R4J_CV_GetFilters(sProjectId)

} ; eofun

; -------------------------------------------------------------------------------------------------------------------

R4J_CV_Report(Debug:=False) {
If GetKeyState("Ctrl")  {
    PowerTools_OpenDoc("r4j_CVReport")
    return
}

static S_CR_DestFolder := A_ScriptDir
static S_JsonFile := "r4j"

; Set environment variables
JiraPassword := Jira_GetPassword()
If (JiraPassword = "") ; cancel
    return	
EnvSet, JIRA_PASSWORD, %JiraPassword%

JiraRootUrl := Jira_GetRootUrl()
If (JiraRootUrl = "")
    return
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
CVArray := R4J_CV_GetFilters(sProjectId,"Select Views to be used in the Report:")
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
R4J_CV_GetFilters(sProjectId,sPrompt:="Select Views:" ){
; CVArray := R4J_CV_GetFilters(sProjectId,sPrompt*)
; return an array of view objects with properties "id","name" for the input Project
; e.g. id of first view is CVArray[1][id]
; user is prompted for selection

sUrl := Jira_GetRootUrl() . "/rest/com.easesolutions.jira.plugins.coverage/1.0/filter/getFilters?projectId=" . sProjectId
sResponse := Jira_Get(sUrl)

static S_CV_Ids := []

CVAllArray := []
sPat = "id":"([^"]*)","name":"([^"]*)"
Pos := 1
While Pos := RegExMatch(sResponse,sPat,sMatch,Pos+StrLen(sMatch)){
    sel := HasVal(S_CV_Ids,sMatch1) 
    CVAllArray.Push({"name":sMatch2,"id":sMatch1,"select":sel})
}

CVArray := ListView("Coverage Views Selector",sPrompt,CVAllArray,"View")
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
R4J_OpenCoverage(sProjectKey){
JiraRootUrl := Jira_GetRootUrl()
sUrl := JiraRootUrl . "/plugins/servlet/com.easesolutions.jira.plugins.requirements/coverage?prj=" . sProjectKey 
Run, "%sUrl%"
}

; -------------------------------------------------------------------------------------------------------------------
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


/rest/com.easesolutions.jira.plugins.requirements/1.0/issue/req-path?jql=issuekey=
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

HasVal(haystack, needle) {
    for index, value in haystack
        if (value = needle)
            return index
    if !IsObject(haystack)
        throw Exception("Bad haystack!", -1, haystack)
    return 0
}