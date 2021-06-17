; ToStartup(sFile,True) : add sFile to Startup
; ToStartup(sFile,False) : remove sFile from Startup
; ToStartup(sFile) returns True if File Shortcut exists in startup and False else
ToStartup(sFile,Toggle := ""){
sLnk := RegExReplace(sFile,"\..*",".lnk")
sLnk := RegExReplace(sLnk,".*\\",A_Startup . "\")

If (Toggle="")
    return FileExist(sLnk)
Else If (Toggle = 1) {
    FileIcon := StrReplace(sFile,".ahk",".ico")
    If FileExist(FileIcon)
        FileCreateShortcut, %sFile%, %sLnk% ,,,,%FileIcon%		; will overwrite existing shortcut
    Else
        FileCreateShortcut, %sFile%, %sLnk% 		; will overwrite existing shortcut
    TrayTipAutoHide("Startup setting", "File ''" . RegExReplace(sLnk,".*\\","") . "'' was added to Startup!")
} Else If (Toggle=0) {
    If FileExist(sLnk) {
        FileDelete, %sLnk%
        TrayTipAutoHide("Startup setting", "File ''" . RegExReplace(sLnk,".*\\","") . "'' was removed from Startup!")
    }       
}
Run, %A_Startup%
}