#SingleInstance force
SetTitleMatchMode, 2
ComposeTitle = Compose Mail - 
    
If WinExist(ComposeTitle . " ahk_exe chrome.exe") {
    WinActivate
    return
}

; Loop on all Chrome Windows    
WinGet, id, List,ahk_exe Chrome.exe
found := False
tabSearch = Gmail
Loop, %id%
{
    hWnd := id%A_Index%
    WinActivate, ahk_id %hWnd%
    ; Loop  on all Tabs in Chrome Window
    WinGetTitle, firstTabTitle, A
    title := firstTabTitle
    Loop {
        if (InStr(title, tabSearch)>0){
            found = True
            break
        }
        Send ^{Tab} ; switch to next tab
        Sleep, 50
        WinGetTitle, title, A  ;get active window title
        if (title = firstTabTitle){
            break
        }

    } ; end Loop Tabs
    if (found)
        break
} ; end loop Chrome windows

If !(found) {
    Run run C:\Program Files (x86)\Google\Chrome\Application\chrome.exe https://mail.google.com/mail/ca/u/0/#inbox"
    WinWait, Inbox ahk_exe Chrome.exe
} 

SendInput +c ; Shift+C
;WinWait %ComposeTitle%
;winmove, A,, 1750, 303, 1725, 935; moving the window to my preferred position
return

