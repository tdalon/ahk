; Library for SharePoint Utilities

; -------------------------------------------------------------------------------------------------------------------

; Sharepoint_CleanUrl - Get clean sharepoint document library Url from browser url
; Syntax:
;	newurl := Sharepoint_CleanUrl(url)
;
#Include <UriDecode>
#Include <People>
; for People_GetMyOUid Personal OneDrive url

#Include <Teams>
; for Teams Name in IntelliPaste
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
If RegExMatch(url,"https://[a-z\-]+\.sharepoint\.com/.*") or InStr(url,"https://mspe.")  
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
; Called by: IntelliPaste -> IntelliHtml
SharePoint_Link2Text(sLink){
sLink := RegExReplace(sLink, "\?.*$","") ; remove e.g. ?d=
If RegExMatch(sLink,"^https://[^/]*/[^/]*/(.*)",sMatch) {
	sMatch1 := uriDecode(sMatch1)
	linktext := sMatch1
	If Not InStr(linktext,"/") ; item in root level= no breadcrumb
		return linktext
	linktext := StrReplace(sMatch1,"/"," > ") ; Breadcrumb navigation for Teams link to folder

	; Choose how to display: with breadcroumb or only last level
	FileName := RegExReplace(sMatch1,".*/","")
	Result := ListBox("IntelliPaste: File Link","Choose how to display",linktext . "|" . FileName,1)
	If Not (Result ="")
		linktext := Result
	return linktext
}
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

SharePoint_IntelliHtml(sLink){
; breadcrumb
If RegExMatch(sLink,"(^https://[^/]*/[^/]*)/(.*)",sMatch) {
	DocPath := uriDecode(sMatch2)
	sLink := sMatch1
	Loop, Parse, DocPath, /
	{
		sLink = %sLink%/%A_LoopField%
		linktext := RegExReplace(A_LoopField, "\?.*$","") ; remove e.g. ?d=
		sHtml = %sHtml% > <a href="%sLink%">%linktext%</a>
	}
	sHtml := SubStr(sHtml,3) ; remove starting >
} 
return sHtml

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

GetRootUrl(sUrl){
    RegExReplace(sUrl,"https?://[^/]",rootUrl)
    return rootUrl
}

; -------------------------------------------------------------------------------------------------------------------

SharePoint_GetSyncIniFile(sFile := ""){
	EnvGet, sOneDriveDir , onedrive
	If !sFile {
		FoundPos := InStr(sOneDriveDir, "\" , , -1)
		sOneDriveDirPar = SubStr(sOneDriveDir,1,FoundPos-1)
		If !InStr(sFile, sOneDrivePar)
			return
		FoundPos := InStr(sFile, "\" , StrLen(sOneDriveDirPar) )
		sDir := SubStr(sFile,1,FoundPos-1)
		sIniFile = %sDir%\SPsync.ini
		
	} Else { ; Default Tenant
		
		sOneDriveDir := StrReplace(sOneDriveDir,"OneDrive - ","")
		sIniFile = %sOneDriveDir%\SPsync.ini
	}
	return sIniFile

}
; -------------------------------------------------------------------------------------------------------------------

SharePoint_GetSyncDir(){
    EnvGet, sOneDriveDir , onedrive
	sOneDriveDir := StrReplace(sOneDriveDir,"OneDrive - ","")
    return sOneDriveDir
}
; -------------------------------------------------------------------------------------------------------------------
SharePoint_UpdateSync(){
; showWarn := SharePoint_UpdateSync()




FileRead, IniContent, %sIniFile%

showWarn := False
Loop, Reg, HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive, K
{
	
	RegRead MountPoint, HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive\%A_LoopRegName%, MountPoint
	MountPoint := StrReplace(MountPoint,"\\","\")

	; Exclude Personal OneDrive
	If InStr(MountPoint,"\OneDrive -")
		Continue

	FoundPos := InStr(MountPoint, "\" , , -1)
	sOneDriveDir = SubStr(MountPoint,1,FoundPos-1)
	sIniFile = %sOneDriveDir%\SPsync.ini

	If Not FileExist(sIniFile)
	{
		TrayTip, NWS PowerTool, File %sIniFile% does not exist! File was created. Fill it following user documentation.

		FileAppend, REM See documentation https://tdalon.github.io/ahk/Sync`n, %sIniFile% 
		FileAppend, REM Use a TAB to separate local root folder from SharePoint sync root url`n, %sIniFile%
		FileAppend, REM It might be the default mapping is wrong if you've synced from a subfolder not in the first level. Url shall not end with /`n, %sIniFile%
		Run "%sIniFile%"

	}

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

SharePoint_UpdateSyncIniFile(sIniFile:=""){
; showWarn := SharePoint_UpdateSyncIniFile(sIniFile:="")
If (sIniFile="")
	sIniFile := SharePoint_GetSyncIniFile()


If Not FileExist(sIniFile)
{
	TrayTip, NWS PowerTool, File %sIniFile% does not exist! File was created in "%sOneDriveDir%". Fill it following user documentation.

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


If RegExMatch(newurl,"https://[^/]/[^/]/[^/]/[^/]*Documents",rooturl) { ; ?: non capturing group
	;MsgBox %newurl% %rooturl%
	needle := "(.*)\t" rooturl "(.*)"
	needle := StrReplace(needle," ","(?:%20| )")
	Loop, Read, %sIniFile%
	{
	If RegExMatch(A_LoopReadLine, needle, match) {	
		;MsgBox %rooturl% 1: %match1% 2: %match2%		
		sFile := StrReplace(newurl, rooturl . match2,Trim(match1) ) ; . "/"
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
}


; -------------------------------------------------------------------------------------------------------------------
SharePoint_Sync2Url(sFile){
; sUrl := SharePoint_Sync2Url(sFile)
; returns empty if not sync'ed

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
FileRead, IniContent, %sIniFile%
If Not InStr(IniContent,syncDir . A_Tab) {
	doWarn := SharePoint_UpdateSyncIniFile(sIniFile)
	If (doWarn)
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

TrayTip, NWS PowerTool, File SPsync.ini is not properly filled for %syncDir%! Fill it following user documentation.,,0x23
Run "%sIniFile%"
return
} ; eofun