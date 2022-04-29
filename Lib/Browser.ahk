
Browser_WinActive(){
GroupAdd, Browser, ahk_exe iexplore.exe
GroupAdd, Browser, ahk_exe chrome.exe
GroupAdd, Browser, ahk_exe sidekick.exe
GroupAdd, Browser, ahk_exe firefox.exe
GroupAdd, Browser, ahk_exe palemoon.exe
GroupAdd, Browser, ahk_exe waterfox.exe
GroupAdd, Browser, ahk_exe vivaldi.exe
GroupAdd, Browser, ahk_exe msedge.exe	
GroupAdd, Browser, ahk_exe ApplicationFrameHost.exe ; Edge or IE with Win10
    
return WinActive("ahk_group Browser")
} ; eofun

; -----------------------------------------------------------------------------------------------------------------

; ------------------------------------------------------------------------------------------------------
; ------------------------------------------------------------------------------------------------------
Browser_GetUrl(WinTitle*)
; https://gist.github.com/anonymous1184/7cce378c9dfdaf733cb3ca6df345b140
; Version: 2022.02.27.1
{
	hWnd := WinExist(WinTitle*)
	if (!hWnd)
		throw Exception("Couldn't find the window.", -1)
	oAcc := Acc_ObjectFromWindow(hWnd)
	oAcc := GetUrl_Recurse(oAcc)
	return oAcc.accValue(0)
}

GetUrl_Recurse(oAcc)
{
	if (oAcc.accValue(0) ~= "^http")
		return oAcc
	for i,accChild in Acc_Children(oAcc) {
		oAcc := GetUrl_Recurse(accChild)
		if IsObject(oAcc)
			return oAcc
	}
}