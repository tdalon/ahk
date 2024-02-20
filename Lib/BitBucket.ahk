
BitBucket_CleanLink(sUrl){
; link := BitBucket_CleanLink(sUrl)
; link[1]: Link url
; link[2]: Link display text
; Paste link with Clip_PasteLink(link[1],link[2])
; Called by: IntelliPaste

If !RegExMatch(sUrl, "/projects/([^/]*)/repos/([^/]*)/browse/")
    return
linkText := RegExReplace(sUrl,"https?://[^/]*/projects/") ; remove root url
linkText := StrReplace(linkText, "/repos/" ,">")
linkText := StrReplace(linkText, "/browse/" ,">")
linkText :=  "BitBucket:" . linkText

return [sUrl, linkText]
}
