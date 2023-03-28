; Library for SharePoint Utilities

; -------------------------------------------------------------------------------------------------------------------

; Sharepoint_CleanUrl - Get clean sharepoint document library Url from browser url
; Syntax:
;	newurl := Sharepoint_CleanUrl(url)
;
#Include <UriDecode>
#Include <People>
; for People_GetMyOUid Personal OneDrive url

; Calls: uriDecode
; Called by: CleanUrl
SharePoint_CleanUrl(url){
	; remove ending filesep 
	If (SubStr(url,0) == "/") ; file or url
		url := SubStr(url,1,StrLen(url)-1)	

	If InStr(url,"_vti_history") ; special case hardlink for old sharepoints
	{
		url := uriDecode(url)
		RegExMatch(url,"\?url=(.*)",newurl)
		return newurl1
	}

	; For new spo links
	If RegExMatch(url ,":[a-z]:/r/") { ; New SPO links
		url := RegExReplace(url ,":[a-z]:/r/","")
		; keep durable link
		If RegExMatch(url,"\?d=")
			url := RegExReplace(url ,"&.*","")
		Else
			url := RegExReplace(url ,"\?.*","")
		return url
	}

	url := StrReplace(url,"?Web=1","") ; old sharepoint link opening in browser


	; For old SP links
	RegExMatch(url,"https://([^/]*)",rooturl) 
	rooturl = https://%rooturl1%

	If !RegExMatch(url,"(?:\?|&)RootFolder=([^&]*)",RootFolder) 
		If !RegExMatch(url,"(?:\?|&)id=([^&]*)",RootFolder) 
			return RegExReplace(url,"/Forms/AllItems.aspx.*$","")
		
	; exclude & 
	; non-capturing group starts with ?: see https://autohotkey.com/docs/misc/RegEx-QuickRef.htm
		
	; decode url			
	RootFolder:= uriDecode(RootFolder1)
	;msgbox %RootFolder%	
	
	newurl := rooturl . RootFolder
	;MsgBox %newurl%
	return newurl	
}
; -------------------------------------------------------------------------------------------------------------------

; SharePoint_IsUrl(url)
SharePoint_IsUrl(url){
If RegExMatch(url,"https://[^\.]+\.sharepoint\.com/.*") or InStr(url,"https://mspe.")  
{
	; url with a few letters followed by one number
    ;  or InStr(url,"https://tenantname.sharepoint.com/") or InStr(url,"https://tenantname-my.sharepoint.com/") 
	return true
	}
Else {
	return false
	}
}

; -------------------------------------------------------------------------------------------------------------------
; NOT USED
SharePoint_Link2Text(sLink){
; linktext := SharePoint_Link2Text(sLink)
sLink := RegExReplace(sLink, "\?.*$","") ; remove e.g. ?d=
If RegExMatch(sLink,"^https://[^/]*/[^/]*/(.*)",sMatch) {
	sMatch1 := uriDecode(sMatch1)
	linktext := sMatch1
	;MsgBox %linktext%
	If Not InStr(linktext,"/") ; item in root level= no breadcrumb
		return linktext
	linktext := StrReplace(sMatch1,"/"," > ") ; Breadcrumb navigation for Teams link to folder

	; Choose how to display: with breadcrumb or only last level
	FileName := RegExReplace(sMatch1,".*/","")
	linktext := ListBox("IntelliPaste: File Link","Choose how to display",linktext . "|" . FileName,1)
	return linktext
}
} ; eofun
; -------------------------------------------------------------------------------------------------------------------


; -------------------------------------------------------------------------------------------------------------------
; Called by IntelliPaste
SharePoint_Link2Html(sLink) {
; sHtml := SharePoint_Link2Html(sLink)
; User will be prompted if he wants a breadcumb format or just the file name as Html link display text

sLink := RegExReplace(sLink, "\?.*$","") ; remove e.g. ?d=

; Prompt user for link text format: breadcrumb vs. filename
If !RegExMatch(sLink,"^(https://[^/]*/[^/]*)/(.*)",sMatch) 
	return
	
sMatch2 := uriDecode(sMatch2)
linktext := sMatch2
;MsgBox %linktext%
If InStr(linktext,"/") { ; item in root level= no breadcrumb
	linktext := StrReplace(sMatch2,"/"," > ") ; Breadcrumb navigation for Teams link to folder
	; Choose how to display: with breadcrumb or only last level
	FileName := RegExReplace(sMatch2,".*/","")
	linktext := ListBox("IntelliPaste: File Link","Choose how to display",linktext . "|" . FileName,1)
	If (linktext = "") ; user cancelled
		return
}


If !InStr(linktext,">") {
	sHtml := "<a href=""" . sLink . """>" . linktext . "</a>"
	return sHtml
}

; use breadcrumbs
sLink := sMatch1
Loop, Parse, sMatch2, /
{
	sLink = %sLink%/%A_LoopField%
	sHtml = %sHtml% > <a href="%sLink%">%A_LoopField%</a>
}
sHtml := SubStr(sHtml,3) ; remove starting >

return sHtml

} ; eofun

; -------------------------------------------------------------------------------------------------------------------

GetRootUrl(sUrl){
    RegExReplace(sUrl,"https?://[^/]",rootUrl)
    return rootUrl
}

; -------------------------------------------------------------------------------------------------------------------

SharePoint_GetSyncIniFile(sFile := ""){
; sIniFile := SharePoint_GetSyncIniFile(sFile :="")
EnvGet, sOneDriveDir , onedrive
If (sFile = "") {
	sOneDriveDir := StrReplace(sOneDriveDir,"OneDrive - ","")
	sIniFile = %sOneDriveDir%\SPsync.ini
} Else {
	FoundPos := InStr(sOneDriveDir, "\" , , -1)
	sOneDriveDirPar = SubStr(sOneDriveDir,1,FoundPos-1)
	If !InStr(sFile, sOneDrivePar)
		return
	FoundPos := InStr(sFile, "\" , StrLen(sOneDriveDirPar) )
	sDir := SubStr(sFile,1,FoundPos-1)
	sIniFile = %sDir%\SPsync.ini
	
} 
return sIniFile
}
; -------------------------------------------------------------------------------------------------------------------

SharePoint_GetSyncDir(){
; not used
    EnvGet, sOneDriveDir , onedrive
	sOneDriveDir := StrReplace(sOneDriveDir,"OneDrive - ","")
    return sOneDriveDir
}
; -------------------------------------------------------------------------------------------------------------------
SharePoint_UpdateSync(){
; SharePoint_UpdateSync()
; Update SPSync.ini file creating temporary Excel files in each sync directory

If GetKeyState("Ctrl") and !GetKeyState("Shift") {
	Run, "https://tdalon.blogspot.com/2023/02/sharepoint-sync-get-url.html" 
	return
}

sIniFile := SharePoint_GetSyncIniFile()
FileRead, IniContent, %sIniFile%

oExcel := ComObjCreate("Excel.Application") 
oExcel.Visible := False ; DBG
oExcel.DisplayAlerts := false

Loop, Reg, HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive, K
{
	
	RegRead MountPoint, HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive\%A_LoopRegName%, MountPoint
	MountPoint := StrReplace(MountPoint,"\\","\")

	; Exclude Personal OneDrive
	If InStr(MountPoint,"\OneDrive -")
		Continue

	FoundPos := InStr(MountPoint, "\" , , -1)
	sOneDriveDir = SubStr(MountPoint,1,FoundPos-1)	

	If InStr(IniContent,MountPoint . A_Tab) ; already mapped
		Continue
	xlFile := MountPoint . "\SPsync.xlsx"
	; Create Excel file under MountPoint
	
	xlWorkbook := oExcel.Workbooks.Add ;add a new workbook
	oSheet := oExcel.ActiveSheet
	oSheet.Range("A1").Formula := "=CELL(""filename"")" ; escape quotes
	oSheet.Range("A1").Dirty ; so that .Calculate updates the formula

	; Save Workbook https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saveas
	Try ; sometimes return error "Enable to get the SaveAs property" but still work
		xlWorkbook.SaveAs(xlFile,xlWorkbookDefault := 51)
	
	; Calculate Formula
	oSheet.Calculate ; works after Range set to Dirty ; no need Send {f9}

	; Get value
	UrlNamespace := oSheet.Range("A1").Value 
	;Find last /
	FoundPos:=InStr(UrlNamespace,"/",,0)
	UrlNamespace := SubStr(UrlNamespace,1,FoundPos-1)
	
	xlWorkbook.Close(False)
	
	; Delete temp file
	FileDelete, %xlFile%

	If (UrlNamespace = "") {
		sTrayTip := "Error in getting url path for synced location '" . MountPoint . "'!"
		TrayTip Check Mapping in SPsync.ini! , %sTrayTip%,,0x2
		Run "%sIniFile%"
	}
	
	FileAppend, %MountPoint%%A_Tab%%UrlNamespace%`n, %sIniFile%

} ; end Loop


oExcel.Quit()
Run "%sIniFile%"

} ; eofun


; -------------------------------------------------------------------------------------------------------------------
SharePoint_UpdateSyncVBA(xlFile := ""){
; SharePoint_UpdateSyncVBA()
; Update SPSync.ini file using SPSyncIni.xlsm with VBA code
; based on VBA macro: https://gist.github.com/guwidoe/6f0cbcd22850a360c623f235edd2dce2


If (xlFile = "") {
	xlFile := A_ScriptDir . "\SPSyncIni.xlsm"
}

oExcel := ComObjCreate("Excel.Application") 
oExcel.Visible := True ; DBG
oExcel.DisplayAlerts := false

xlWorkbook := oExcel.Workbooks.Open(xlFile) ; Open xlFile
oSheet := oExcel.ActiveSheet
;oSheet.Range("A1").Formula := "=CELL(""filename"")" ; escape quotes



tbl := oSheet.ListObjects("SPSync")

; Reset Table - but keep formulas: https://stackoverflow.com/questions/10220906/how-to-select-clear-table-contents-without-destroying-the-table
;    tbl.DataBodyRange.Delete
If (tbl.ListRows.Count >= 2) 
	tbl.DataBodyRange.Offset(1, 0).Resize(tbl.DataBodyRange.Rows.Count - 1, tbl.DataBodyRange.Columns.Count).Delete

rowCounter := 0	
Loop, Reg, HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive, K
{
	RegRead MountPoint, HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive\%A_LoopRegName%, MountPoint
	MountPoint := StrReplace(MountPoint,"\\","\")

	; Exclude Personal OneDrive
	If InStr(MountPoint,"\OneDrive -")
		Continue

	rowCounter := rowCounter + 1
    
    If rowCounter > 1 ; Create new row and paste formats and formula
        tbl.ListRows.Add ; Will copy previous row with same formatting and formulas

    tbl.DataBodyRange.Cells(rowCounter, 1).Value := MountPoint 
	; Takes pretty long for first evaluation   

} ; end Loop

; Save Workbook https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saveas
; Save as CSV file
sCsvFile := StrReplace(xlFile,".xlsm",".csv")
FileDelete, %sCsvFile%
xlWorkbook.SaveAs(sCsvFile,xlCSV := 6)

xlWorkbook.Close(False)
oExcel.Quit()

; Convert CSV to ini file
sIniFile := A_ScriptDir . "\PowerTools.ini"
IniDelete, %sIniFile%, SPsync
FileRead, IniContent, %sCsvFile%
IniContent := StrReplace(IniContent,",","=")
IniContent := StrReplace(IniContent,"`r","")
; delete header line
pos:=InStr(IniContent,"`n")
IniContent:=SubStr(IniContent,pos+1)

IniWrite, %IniContent%, %sIniFile%, SPsync

;Run, "%sIniFile%"
FileDelete, %sCsvFile%

} ; eofun



; -------------------------------------------------------------------------------------------------------------------

SharePoint_UpdateSyncIniFile(sIniFile:=""){
; showWarn := SharePoint_UpdateSyncIniFile(sIniFile:="")
If (sIniFile="")
	sIniFile := SharePoint_GetSyncIniFile()


If Not FileExist(sIniFile)
{
	TrayTip, NWS PowerTool, File %sIniFile% does not exist! File was created in "%sOneDriveDir%". 
	FileAppend, REM See documentation https://tdalon.github.io/ahk/Sync`n, %sIniFile% 
	FileAppend, REM Use a TAB to separate local root folder from SharePoint sync root url`n, %sIniFile%
	FileAppend, REM It might be the default mapping is wrong if you've synced from a subfolder not in the first level. Url shall not end with /`n, %sIniFile%

}

FileRead, IniContent, %sIniFile%

showWarn := False
Loop, Reg, HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive, K
{
	
	RegRead MountPoint, HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive\%A_LoopRegName%, MountPoint
	MountPoint := StrReplace(MountPoint,"\\","\")

	; Exclude Personal OneDrive
	If InStr(MountPoint,"\OneDrive -")
		Continue

	If Not InStr(IniContent,MountPoint . A_Tab) {
		RegExMatch(MountPoint,"[^\\]*$",sFolderName)

		RegRead UrlNamespace, HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive\%A_LoopRegName%, UrlNamespace
		
		sFolderName := RegExReplace(sFolderName,"^EXT - ","") ; Special case for EXT -
		

		If FolderName := RegExMatch(sFolderName,"^[^-]* - ([^-]*) - ([^-]*)$",sMatch) { ;  Private Channel
			If (sMatch1 = sMatch2) { ; root folder has same name
				UrlNamespace := SubStr(UrlNamespace,1,-1) ; remove trailing /	
			} Else {
				UrlNamespace := UrlNamespace . sMatch2
				showWarn := True
			}
		} Else {
			FolderName := RegExReplace(sFolderName,".*- ","")
			
			If Not (FolderName = "Documents") { ; not root level
				UrlNamespace := UrlNamespace . FolderName
				; For Teams SharePoint check for General channel folder to ignore displaying warning
				If RegExMatch(UrlNamespace,"sharepoint\.com/teams/team_")
					If Not (FolderName = "General")
						showWarn := True
				Else
					showWarn := True
			} Else ; root level -> remove trailing /
				UrlNamespace := SubStr(UrlNamespace,1,-1)

		}
		FileAppend, %MountPoint%%A_Tab%%UrlNamespace%`n, %sIniFile%

	}
} ; end Loop

If (showWarn) {
	sTrayTip = If you are not syncing on the root level, you need to check the default mapping!
	TrayTip Check Mapping in SPsync.ini! , %sTrayTip%,,0x2
	Run "%sIniFile%"
}

return showWarn

} ; eofun

; -------------------------------------------------------------------------------------------------------------------
SharePoint_Url2Sync(sUrl,sIniFile:=""){
; sFile := SharePoint_Url2Sync(sUrl,sIniFile*)
; returns empty if not sync'ed

If (sIniFile="")
	sIniFile := SharePoint_GetSyncIniFile()
If !FileExist(sIniFile) {
	SharePoint_UpdateSyncIniFile(sIniFile)
}


If RegExMatch(sUrl,"https://[^/]/[^/]/[^/]/[^/]*Documents",rooturl) { ; ?: non capturing group
	;MsgBox %newurl% %rooturl%
	needle := "(.*)\t" . rooturl . "(.*)"
	needle := StrReplace(needle," ","(?:%20| )")
	Loop, Read, %sIniFile%
	{
	If RegExMatch(A_LoopReadLine, needle, match) {	
		;MsgBox %rooturl% 1: %match1% 2: %match2%		
		sFile := StrReplace(sUrl, rooturl . match2,Trim(match1) ) ; . "/"
		sFile := StrReplace(sFile, "/", "\")
		
		; MsgBox %A_LoopReadLine% %needle% 
		return sFile
		}
	}
}

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

SharePoint_IsSPWithSync(sUrl){
; returns true if SPO SharePoint or mspe SharePoint
return RegExMatch(sUrl,"https://[^/]*\.sharepoint.com") or RegExMatch(sUrl,"https://mspe\..*")
} ; eofun


; -------------------------------------------------------------------------------------------------------------------
SharePoint_Sync2Url(sFile, doWarn:= true){
; sUrl := SharePoint_Sync2Url(sFile)
; 	returns empty if File is not located in a synced location or could not be found as of in the SPsync.ini file in which case an Update might be needed.
; Called by IntelliPaste->GetFileLink function

; Get File Link for Personal OneDrive
EnvGet, sOneDriveDir , onedrive
If InStr(sFile,sOneDriveDir . "\") { 
	RegRead, rooturl, HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive, UrlNamespace
	rooturl := SubStr(rooturl, 1 , -1) ; rooturl ends with / -> remove
	sFile := StrReplace(sFile, sOneDriveDir,rootUrl)
	sFile := StrReplace(sFile, "\", "/")
	return sFile
}

; Get File Link for SharePoint/OneDrive Synced location
sOneDriveDir := StrReplace(sOneDriveDir,"OneDrive - ","")
needle :=  StrReplace(sOneDriveDir,"\","\\") ; 
needle := needle "\\[^\\]*"
If Not (RegExMatch(sFile,needle,syncDir))
	Return

sIniFile = %sOneDriveDir%\SPSync.ini
If Not FileExist(sIniFile)
	SharePoint_UpdateSync()

FileRead, IniContent, %sIniFile%
If Not InStr(IniContent,syncDir . A_Tab)  {
	If doWarn
		TrayTip, NWS PowerTool, File SPsync.ini does not contain %syncDir%! You might need to update the Sync Ini file.,,0x23
	;Run "%sIniFile%"
	return
}

Loop, Read, %sIniFile%
{
	Array := StrSplit(A_LoopReadLine, A_Tab," `t",2)
	If !Array
		continue
	rootDir := StrReplace(Array[1],"??",".*") ; for emojis
	rootDirRe := StrReplace(rootDir,"\","\\") ; escape filesep
	If (RegExMatch(syncDir, rootDirRe)) {
		rootUrl := Array[2]
		sFile := StrReplace(sFile, syncDir,rootUrl)
		sFile := StrReplace(sFile, "\", "/")
		return sFile
	}
}	; End Loop		

} ; eofun