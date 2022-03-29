#Include <Clip>
Blogger_GetBlogId(sUrl){
; Truncate Url to root
RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
;sApiUrl := "https://www.googleapis.com/blogger/v3/blogs/byurl?url=" + sUrl 
sApiUrl := sRootUrl . "/feeds/posts/default"
sResponse := HttpGet(sApiUrl)
; Parse for <id>tag:blogger.com,1999:blog-7106641098407922697</id>
RegExMatch(sResponse,"<id>([^<]*)</id>",sId)
sId := RegExReplace(sId1,".*-","")
return sId
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Blogger_GetPostId(sUrl) {
; <meta content='7106641098407922697' itemprop='blogId'/>
; <meta content='966510408740448032' itemprop='postId'/>
 ; View Source
sResponse := HttpGet(sUrl)
RegExMatch(sResponse,"<meta content='(\d*)' itemprop='postId'/>", sMatch)
return sMatch1
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Blogger_Edit(sUrl){
sUrl := RegExReplace(sUrl,"#.*","") ; remove section
sResponse := HttpGet(sUrl)
If (sResponse ="")
    return
RegExMatch(sResponse,"<meta content='(\d*)' itemprop='blogId'/>", sBlogId)
If (sBlogId="") {
    MsgBox 0x10, Error, Error BlogId not found from Html content!
    return
}
RegExMatch(sResponse,"<meta content='(\d*)' itemprop='postId'/>", sPostId)
sEditUrl = https://www.blogger.com/blog/post/edit/%sBlogId1%/%sPostId1%
; Overwrite current window but stays in same browser profile
Send ^l
SendInput, %sEditUrl%{Enter}

} ; eofun
; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------

Blogger_IsUrl(sUrl){
return InStr(sUrl,".blogspot.com")
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Blogger_Search(sUrl,sDefSearch:=""){
RegExMatch(sUrl,"https?://[^/]*",sRootUrl)
ReRootUrl := StrReplace(sRootUrl,".","\.")
If (sDefSearch = "") {
    ; Get def search from input url
    If RegExMatch(sUrl,"/search/label/([^\?&]*)",sMatch) {
        sDefSearch = #%sMatch1%
    } Else If RegExMatch(sUrl,"\?q=([^&]*)",sDefSearch) {
       
        sDefSearch := sDefSearch1	
        sPat := "label:([^\+]*)" 
        Pos=1
        While Pos :=    RegExMatch(sDefSearch, sPat, tag,Pos+StrLen(tag)) 
            sTags := sTags . " #" . tag1 
        
        sDefSearch := RegExReplace(sDefSearch,sPat,"")
        sDefSearch := sTags . "+" . sDefSearch
        
        sDefSearch := StrReplace(sDefSearch,"%20"," ")
        sDefSearch := StrReplace(sDefSearch,"+"," ") 
        sDefSearch := Trim(sDefSearch) 
    }

    InputBox, sSearch , Blogger Search, Enter search string (use # for tags):,,640,125,,,,,%sDefSearch% 
	if ErrorLevel
		return
	sSearch := Trim(sSearch) 
}

sPat := "#([^\s]*)" 
sTags =
Pos=1
While Pos :=    RegExMatch(sSearch, sPat, tag,Pos+StrLen(tag)) 
    sTags := sTags . "+label:" . tag1
If Not (sTags = "") {
    sTags := SubStr(sTags,2) ; remove starting +
    sSearch := RegExReplace(sSearch,sPat,"")
    sSearch := Trim(sSearch) 
    If (sSearch = "")
        sSearch := sTags
    Else
        sSearch := sTags . "+" . sSearch
}
sSearchUrl := sRootUrl . "/search?q=" . sSearch . "&max-results=500"
;Run, %sSearchUrl%
Send ^n ; New Window
Sleep 500
Clip_Paste(sSearchUrl)
Send {Enter}
} ; eofun
; -------------------------------------------------------------------------------------------------------------------