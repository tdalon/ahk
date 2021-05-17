; AHK Library
; Includes Startup, Compile
AHK_Close(ScriptFullPath){
DetectHiddenWindows, On
WinClose, %ScriptFullPath% ahk_class AutoHotkey
}

; ---------------------------------------------------------------------- 
AHK_IsRunning(ScriptFullPath){
DetectHiddenWindows, On
IfWinExist, %ScriptFullPath%
    return True
return False
}

; ---------------------------------------------------------------------- 

AHK_Compile(ScriptFullPath,FileIcon:="",CompileDir:=""){
; See https://www.autohotkey.com/boards/viewtopic.php?t=60944

SplitPath A_AhkPath,, AhkDir

SplitPath,ScriptFullPath, OutFileName
ExeFileName := StrReplace(OutFileName,".ahk",".exe")
Run, taskkill.exe /F /IM %ExeFileName%
If (FileIcon="") {
    FileIcon := StrReplace(ScriptFullPath,".ahk",".ico")
}

FileRead, sCode, %ScriptFullPath%
sCode := RegExReplace(sCode,"LastCompiled =.*","LastCompiled = " . A_Now)
FileDelete, %ScriptFullPath%
FileAppend, %sCode%, %ScriptFullPath%

; sCmd := "`"" . AhkDir . "\Compiler\Ahk2Exe.exe`" /in `"" . ScriptFullPath . "`""
sCmd = "%AhkDir%\Compiler\Ahk2Exe.exe" /in "%ScriptFullPath%"
If FileExist(FileIcon)
    sCmd = %sCmd% /icon "%FileIcon%"

; Move File to CompileDir

If Not (CompileDir="") {
    If Not InStr(CompileDir,":") { ; relative path
        SplitPath, ScriptFullPath, ScriptFileName, OutDir
        CompileDir = %OutDir%\%CompileDir%
    }
    DestFile = %CompileDir%\%ExeFileName%
    sCmd = %sCmd% /out "%DestFile%"
}
RunWait %sCmd%
}


; ---------------------------------------------------------------------- 

AHK_Exit(ScriptFullPath:=""){
If !ScriptFullPath
    ScriptFullPath=A_ScriptFullPath

SplitPath,ScriptFullPath, FileName, OutDir, Extension
If (Extension = "ahk") {
    DetectHiddenWindows, On
    WinClose, %ScriptFullPath% ahk_class AutoHotkey
} Else {
    ;ExeFileName := StrReplace(OutFileName,".ahk",".exe")
    sCmd = taskkill /f /im "%FileName%"
    Run %sCmd%,,Hide
}
}

; ---------------------------------------------------------------------- 
AHK_Help(kw){
; Based on AHK Help Launcher. Credit RaptorX https://www.the-automator.com/autohotkey-webinar-notepad-a-solid-well-loved-but-dated-autohotkey-editor/
pwb := WinExist("ahk_id " pwbHandle) ? pwb : ComObjCreate("InternetExplorer.Application"), pwbHandle := pwb.hwnd
pwb.navigate("https://www.autohotkey.com/docs/")
pwb.addressbar:=false
pwb.ToolBar:=false
pwb.Statusbar:=false
pwb.visible := true

WinActivate, % "Ahk_id " pwb.hwnd

while (pwb.busy || pwb.ReadyState != 4)				;Wait for page to load
	sleep 50

while !(pwb.document.querySelectorAll("#left")[0])	; make sure the element exists before performing any more actions
	Sleep, 50

pwb.document.querySelectorAll("#left > div.search > div.input > input[type=search]")[0].value := kw
pwb.document.querySelectorAll("button[aria-label='Search tab']")[0].click()

ControlSend, Internet Explorer_Server1, {enter 2}, % "Ahk_id " pwb.hwnd
}