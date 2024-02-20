BigPicture_IsUrl(sUrl) {
    return  InStr(sUrl,"bigpicture")
} ; eofun

BigPicture_Redirect(sUrl,tgtRootUrl:="") {
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


; box
; $rootsv/plugins/servlet/softwareplant-bigpicture/#/box/PROG-58/g
; $rootc/plugins/servlet/ac/eu.softwareplant.bigpicture/bigpicture#!box/PROG-58/g

If (!InStr(tgtUrl,".atlassian.net") & InStr(sUrl,".atlassian.net")) { ; cloud to server
    tgtUrl := StrReplace(tgtUrl,"/ac/eu.softwareplant.bigpicture/bigpicture#!","/softwareplant-bigpicture/#/")
} Else If (InStr(tgtUrl,".atlassian.net") & !InStr(sUrl,".atlassian.net")) { ; server to cloud
    tgtUrl := StrReplace(tgtUrl,"/softwareplant-bigpicture/#/","/ac/eu.softwareplant.bigpicture/bigpicture#!")
}
return tgtUrl

} ; eofun