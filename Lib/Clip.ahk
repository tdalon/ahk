#Include <WinClipAPI>
#Include <WinClip>
; Uses WinClip For GeSelectionHtml->WinClip.GetHTML and Paste*->WinClip.Paste

GroupAdd, PlainEditor, ahk_exe Notepad.exe
GroupAdd, PlainEditor, ahk_exe notepad++.exe
GroupAdd, PlainEditor, ahk_exe atom.exe
; -------------------------------------------------------------------------------------------------------------------
Clip_Paste(sText) {
; Syntax: Clip_Paste(sText)
WinClip.Paste(sText)
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Clip_Set(sText){
; Syntax: Clip_Set(sText)
; 
Clipboard := sText
Clip_Wait()
} ; eofun
; -------------------------------------------------------------------------------------------------------------------
Clip_Restore(ClipBackup) {
Clip_Wait() ; in order not to overwrite running clipboard action like pasting
Clipboard:= ClipBackup
} ;eofun
; -------------------------------------------------------------------------------------------------------------------
Clip_Wait(){
Sleep, 150
while DllCall("user32\GetOpenClipboardWindow", "Ptr")
    Sleep, -1
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Clip_GrabWord(){
; Based on AHK Help Launcher. Credit RaptorX https://www.the-automator.com/autohotkey-webinar-notepad-a-solid-well-loved-but-dated-autohotkey-editor/
oldClip := clipboardAll
;If WinActive("ahk_exe Code.exe") or WinActive("ahk_exe Code - Insiders.exe") ; needs to select word with double click else copy whole line

clipboard =
sendinput ^c
Clip_Wait()
if (clipboard == "") ; no word selected -> use caret cursor location
{
    ; get char to the left of the cursor
    SendInput +{Left}^c{Right}
	clipwait 0
	LeftChar := regexmatch(clipboard, "\w")

	clipboard = 
	SendInput % (leftChar ? "^{left}" : "") "^+{right}^c"
	clipwait 0
	regexmatch(clipboard, "\w+", grab)
    return grab
    
    /* ; Ctrl+Right and then Ctrl+Shift+Left to select word
    SendInput ^{Right}
    SendInput ^+{Left}
    SendInput ^c
	clipwait 0
	grab := clipboard 
    */

}
else
	grab := trim(clipboard)

clipboard := oldClip ; restore clipboard
return grab
}

; -------------------------------------------------------------------------------------------------------------------
Clip_SetHtml(sHtml,sText:=""){
; Syntax: Clip_SetHtml(sHtml,sText:="",HtmlHead := "")
; If sHtml is a link (starts with http), the Html link will be wrapped around it sHtml=<a href="%sHtml%">%sText%</a>
; If no Text display is passed as argument sText := sHtml
If (sText = "")
    sText := sHtml
If RegExMatch(sHtml,"^http") {
    WinClip.SetText(sHtml)
    sHtml := "<a href=""" . sHtml . """>" . sText . "</a>"
} else {
    WinClip.SetText(sText)
}
;SetClipboardHTML(sHtml,HtmlHead,sText) ; does not work with WinClip.GetHtml
; WinClip.iSetHTML does not work (asked here https://www.autohotkey.com/boards/viewtopic.php?f=6&t=29314&p=393505#p393505)
WinClip.SetHTML(sHtml)
Clip_Wait()
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
Clip_PasteHtml2(sHtml,sText:="",restore := True) {
; Syntax: Clip_PasteHtml(sHtml,sText,restore := True)
; If sHtml is a link (starts with http), the Html link will be wrapped around it i.e. sHtml=<a href="%sHtml%">%sText%</a>
; Replaced by Clip_PasteHtml - see https://tdalon.blogspot.com/2023/03/confluence-intellipaste-issue.html
If (restore)
    ClipBackup := ClipboardAll
Clip_SetHtml(sHtml,sText)
WinClip.Paste()  
If (restore) 
    Clip_Restore(ClipBackup)
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------
Clip_PasteHtml(sHtml,restore := True) {
; Syntax: Clip_PasteHtml(sHtml,restore := True)
; If sHtml is a link (starts with http), the Html link will be wrapped around it i.e. sHtml=<a href="%sHtml%">%sText%</a>
If (restore)
    ClipBackup := ClipboardAll
SetClipboardHTML(sHtml)
WinClip.Paste()  
If (restore) 
    Clip_Restore(ClipBackup)
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

; -------------------------------------------------------------------------------------------------------------------
Clip_PasteLink(sUrl, sText:="", doEdit := True, restore:=True) {
; Syntax: Clip_PasteLink(sUrl,sText,doEdit,restore := True)
; 

If !sText
    sText := sUrl

If WinActive("ahk_group PlainEditor") {

    WinGetTitle, sTitle , A
    If InStr(sTitle,".md") { ; markdown format
        sLink =[%sText%](%sUrl%)
        Clip_Paste(sClipboard)
        return
    }
    Clip_Paste(sUrl)
    return
}

If (doEdit) {
    InputBox, sText , Display Link Text, Enter Link display text:,,640,125,,,,, %sText%
    if ErrorLevel ; Cancel
        return
}
sHtml =<a href="%sUrl%">%sText%</a>
If (restore)
    ClipBackup:= ClipboardAll
Clip_SetHtml(sHtml,sText)

; Paste
WinClip.Paste()  
If (restore)
    Clip_Restore(ClipBackup)
  
} ; eofun
; -------------------------------------------------------------------------------------------------------------------

Clip_GetSel2(restore:=True){
; SelArr := Clip_GetSel2
; returns an array with selection in both Html and Plain format
; sSelectionArr := Clip_GetSel2(restore:=True)
; sSelectionArr[1] : html format
; sSelectionArr[2] : plain text format

If (restore)
    ClipBackup:= ClipboardAll
Clipboard =
SendInput,^c    
Clip_Wait()
SelArr := []
SelArr[1] := WinClip.GetHTML()
SelArr[2] := clipboard

If (restore)
    Clipboard := ClipBackup 
return SelArr
} ; eofun
; -------------------------------------------------------------------------------------------------------------------


Clip_GetSel(type:="text"){
; Syntax:
; 	sSelection := Clip_GetSel("text"*|"html")
; Default "text"
Clipboard =
while(Clipboard){
  Sleep,10
}
SendInput,^c    
Clip_Wait()

If (type = "text") {
    sSelection := clipboard
} Else If (type ="html") {
    ;Clip_GetHtml(sSelection)
    sSelection := WinClip.GetHTML()
}
;sSelection := Trim(sSelection,"`n`r`t`s")

return sSelection
} 
; -------------------------------------------------------------------------------------------------------------------

Clip_GetSelectionHtml(restore:=True){
; Get selection in html format
; Syntax:
; 	sSelection := GetSelectionHtml()
If (restore)
    ClipBackup:= ClipboardAll
sSelection := Clip_GetSel("html")
If (restore)
    Clipboard := ClipBackup 
return sSelection
} 
; -------------------------------------------------------------------------------------------------------------------
Clip_GetSelection(restore:=True){
; Get selection in plain text format
; Syntax:
; 	sSelection := GetSelection(restore:=False)
If (restore)
    ClipBackup:= ClipboardAll
sSelection := Clip_GetSel("text")

If (restore)
    Clipboard := ClipBackup 
return sSelection 
} 
; -------------------------------------------------------------------------------------------------------------------

Clip_ReplaceSelection(sNeedle, sReplace:=""){
ClipBackup:= ClipboardAll
sSelection := Clip_GetSelection()
If !sSelection { ; no sSelection
    Clip_Restore(ClipBackup)
    Return
}
If InStr(sNeedle,"$0") {
    sNew := StrReplace(sNeedle,"$0",sSelection)
} Else {
    sNew := RegExReplace(sSelection, sNeedle, sReplace)
}
Clip_Paste(sNew)
Clip_Restore(ClipBackup)
}
; -------------------------------------------------------------------------------------------------------------------

Clip_GetHtml(){
sHtml := WinClip.GetHTML()
return sHtml
}

; -------------------------------------------------------------------------------------------------------------------
; ##################### NOT USED #################################

Clip_GetHtmlBuggy( byref Data ) { ; www.autohotkey.com/forum/viewtopic.php?p=392624#392624 -> DOES NOT WORK
 If CBID := DllCall( "RegisterClipboardFormat", Str,"HTML Format", UInt )
  If DllCall( "IsClipboardFormatAvailable", UInt,CBID ) <> 0
   If DllCall( "OpenClipboard", UInt,0 ) <> 0
    If hData := DllCall( "GetClipboardData", UInt,CBID, UInt )
       DataL := DllCall( "GlobalSize", UInt,hData, UInt )
     , pData := DllCall( "GlobalLock", UInt,hData, UInt )
     , VarSetCapacity( data, dataL * ( A_IsUnicode ? 2 : 1 ) ), StrGet := "StrGet"
     , A_IsUnicode ? Data := %StrGet%( pData, dataL, 0 )
                   : DllCall( "lstrcpyn", Str,Data, UInt,pData, UInt,DataL )
     , DllCall( "GlobalUnlock", UInt,hData )
 DllCall( "CloseClipboard" )
Return dataL ? dataL : 0
}

; SetHtml using SetClipboardHTML function
Clip_SetH(sHtml){
    SetClipboardHTML(sHtml)
} ; eofun

; -------------------------------------------------------------------------------------------------------------------
; https://www.autohotkey.com/boards/viewtopic.php?f=6&t=80706&sid=af626493fb4d8358c95469ef05c17563
; Drawback:  winclip.gethtml does not work if SetClipboardHTML before & not from fresh run -> revert to WinClip.SetHTML
SetClipboardHTML(HtmlBody, HtmlHead:="", AltText:="") {       ; v0.67 by SKAN on D393/D42B
Local  F, Html, pMem, Bytes, hMemHTM:=0, hMemTXT:=0, Res1:=1, Res2:=1   ; @ tiny.cc/t80706
Static CF_UNICODETEXT:=13,   CFID:=DllCall("RegisterClipboardFormat", "Str","HTML Format")

  If ! DllCall("OpenClipboard", "Ptr",A_ScriptHwnd)
    Return 0
  Else DllCall("EmptyClipboard")

  If (HtmlBody!="")
  {
      Html     := "Version:0.9`r`nStartHTML:00000000`r`nEndHTML:00000000`r`nStartFragment"
               . ":00000000`r`nEndFragment:00000000`r`n<!DOCTYPE>`r`n<html>`r`n<head>`r`n"
                         . HtmlHead . "`r`n</head>`r`n<body>`r`n<!--StartFragment -->`r`n"
                              . HtmlBody . "`r`n<!--EndFragment -->`r`n</body>`r`n</html>"

      Bytes    := StrPut(Html, "utf-8")
      hMemHTM  := DllCall("GlobalAlloc", "Int",0x42, "Ptr",Bytes+4, "Ptr")
      pMem     := DllCall("GlobalLock", "Ptr",hMemHTM, "Ptr")
      StrPut(Html, pMem, Bytes, "utf-8")

      F := DllCall("Shlwapi.dll\StrStrA", "Ptr",pMem, "AStr","<html>", "Ptr") - pMem
      StrPut(Format("{:08}", F), pMem+23, 8, "utf-8")
      F := DllCall("Shlwapi.dll\StrStrA", "Ptr",pMem, "AStr","</html>", "Ptr") - pMem
      StrPut(Format("{:08}", F), pMem+41, 8, "utf-8")
      F := DllCall("Shlwapi.dll\StrStrA", "Ptr",pMem, "AStr","<!--StartFra", "Ptr") - pMem
      StrPut(Format("{:08}", F), pMem+65, 8, "utf-8")
      F := DllCall("Shlwapi.dll\StrStrA", "Ptr",pMem, "AStr","<!--EndFragm", "Ptr") - pMem
      StrPut(Format("{:08}", F), pMem+87, 8, "utf-8")

      DllCall("GlobalUnlock", "Ptr",hMemHTM)
      Res1  := DllCall("SetClipboardData", "Int",CFID, "Ptr",hMemHTM)
  }

  If (AltText!="")
  {
      Bytes    := StrPut(AltText, "utf-16")
      hMemTXT  := DllCall("GlobalAlloc", "Int",0x42, "Ptr",(Bytes*2)+8, "Ptr")
      pMem     := DllCall("GlobalLock", "Ptr",hMemTXT, "Ptr")
      StrPut(AltText, pMem, Bytes, "utf-16")
      DllCall("GlobalUnlock", "Ptr",hMemTXT)
      Res2  := DllCall("SetClipboardData", "Int",CF_UNICODETEXT, "Ptr",hMemTXT)
  }

  DllCall("CloseClipboard")
  hMemHTM := hMemHTM ? DllCall("GlobalFree", "Ptr",hMemHTM) : 0

Return (Res1 & Res2)
} ;eofun
; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------


;--------------- SafePaste
SafePaste() {
    ; A way of pasting that only returns control when the paste is complete
    ; by jeeswg
    ; See https://www.autohotkey.com/boards/viewtopic.php?p=271514&sid=f898e28c59efcb6871c1dff403e663dd#p271517
    ; the point of this is that with a simple Ctrl + v, you don't know when the pasting is complete,
    ; so if you immediately reload the Clipboard, the new text may end up getting pasted...

    ControlGetFocus, vCtlClassNN, A 
    ControlGet, hCtl, Hwnd,, % vCtlClassNN, A 
    SendMessage, 0x302,,,, % "ahk_id " hCtl ;WM_PASTE := 0x302 
}


;https://autohotkey.com/board/topic/111817-robust-copy-and-paste-routine-function/ -> DOES NOT WORK
ClipPaste() {
sPasteKey := "vk56sc02F" ; Paste

SendInput, {Shift Down}{Shift Up}{Ctrl Down}{%sPasteKey% Down}

; wait for clipboard is ready
iStartTime := A_TickCount
Sleep, % 100
while (DllCall("GetOpenClipboardWindow") && (A_TickCount-iStartTime<1400)) ; timeout = 1400ms
    Sleep, % 100

SendInput, {%sPasteKey% Up}{Ctrl Up}

}

ClipCopy(piMode := 0)
; sMyVar := ClipCopy() ; will copy selected text via control + c
; sMyVar := ClipCopy(1) ; will cut selected text via control + x
{
    clpBackup := ClipboardAll

    Clipboard=

    if (piMode == 1)
        sCopyKey := "vk58sc02D" ; Cut
    else
        sCopyKey := "vk43sc02E" ; Copy

    SendInput, {Shift Down}{Shift Up}{Ctrl Down}{%sCopyKey% Down}
    ClipWait, 0.25
    SendInput, {%sCopyKey% Up}{Ctrl Up}

    sRet := Clipboard

    Clipboard := clpBackup

    return sRet
}