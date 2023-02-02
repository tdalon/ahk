; Library for Goodreads functionality

; Used by NWS.ahk QuickSearch
; ----------------------------------------------------------------------
Goodreads_IsUrl(sUrl){
return  InStr(sUrl,"goodreads.com")  
} ;eofun
; ----------------------------------------------------------------------


; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
; Goodreads QuickSearch - Search within Goodreads
; Called by: NWS.ahk (Win+F Hotkey)
; 
Goodreads_Search(sUrl){

static sGoodreadsSearch
sSearchUrl := "https://www.goodreads.com/review/list/4083745-thierry-dalon?" ; TODO - add setting


InputBox, sSearch , Goodreads Search, Enter search string (use # for tags/shelves):,,640,125,,,,,%sGoodreadsSearch% 
if ErrorLevel
	return
sSearch := Trim(sSearch) 

; ----
; Convert labels to Search pattern &shelf=lab1%2Clab2
sPat := "#([^#\s]*)" 
Pos=1
While Pos := RegExMatch(sSearch, sPat, label,Pos+StrLen(label)) {
	sSearchLabels := sSearchLabels . "%2C" . label1
} ; end while


If sSearchLabels { ; not empty 
    ; Remove first "%2C"
    sSearchLabels := SubStr(sSearchLabels,4)
    sSearchLabels := "&shelf=" . sSearchLabels
    sSearchUrl := sSearchUrl . sSearchLabels
}


; remove labels from search string
sQuerySearch := RegExReplace(sSearch, sPat , "")
sQuerySearch := Trim(sQuerySearch)

If sQuerySearch { ; not empty
	sSearchUrl := sSearchUrl . "&search%5Bquery%5D=" . sQuerySearch
} 

sGoodreadsSearch := sSearch

If InStr(sUrl,"www.goodreads.com/review/list/") ; list search view already -> use same window
	Send ^l
Else
	Send ^n ; New Window
Sleep 500
Clip_Paste(sSearchUrl)
Send {Enter}

} ; eofun Goodreads_Search