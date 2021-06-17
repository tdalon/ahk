; See documentation https://tdalon.blogspot.com/2020/12/chromy.html
#Include <Clip>

; Paste command with internal clipboard to be faster than sendinput

Edgy(A_Args[1])

Edgy(sInput){
FoundPos := InStr(sInput," ")   
sKeyword := SubStr(sInput,1,FoundPos-1)
sInput := SubStr(sInput,FoundPos+1)
;sInput := StrReplace(sInput, "#", "{#}")

If WinActive("ahk_exe msedge.exe") {   
    SendInput , ^t
} Else {
    Run, msedge.exe
}
SendBar(sKeyword,sInput)
} ; End function


; ---------------------------------------------
SendBar(sKeyword,sInput){
;MsgBox %sKeyword% %sInput%
SetTitleMatchMode, 1 ; start with
WinWaitActive, New Tab ahk_exe msedge.exe

SendInput %sKeyword%{tab}
Clip_Paste(sInput)
SendInput {enter}
}