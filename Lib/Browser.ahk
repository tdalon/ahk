
Browser_WinActive(){
GroupAdd, Browser, ahk_exe iexplore.exe
GroupAdd, Browser, ahk_exe chrome.exe
GroupAdd, Browser, ahk_exe firefox.exe
GroupAdd, Browser, ahk_exe palemoon.exe
GroupAdd, Browser, ahk_exe waterfox.exe
GroupAdd, Browser, ahk_exe vivaldi.exe
GroupAdd, Browser, ahk_exe msedge.exe	
GroupAdd, Browser, ahk_exe ApplicationFrameHost.exe ; Edge or IE with Win10
    
return WinActive("ahk_group Browser")
} ; eofun

; -------------------------------------------------------------------------------------------------------------------

;https://www.reddit.com/r/AutoHotkey/comments/m0i7tf/mind_helping_me_on_chaining_this_command/gvgwhie
Browser_GetUrl2(hWnd := "A")
{
    th:= WinExist(hWnd)
    WinGet, bp, ProcessName, ahk_id %th%
    if ( bp ="vivaldi.exe") { ; Acc does not work for vivaldi
        Send ^l
        sUrl := Clip_GetSelection()
        Click
    } Else {  
        accWindow := Acc_ObjectFromWindow(hWnd)
        sUrl := getAddressBar(accWindow).accValue(0)
        If !RegExMatch(sUrl,"https://.*")
            sUrl = https://%sUrl%
    }
    return sUrl
}
; ------------------------------------------------------------------------------------------------------
getAddressBar(accObj)
{
    if (accObj.accRole(0) == 42
    && accObj.accValue(0) != "")
        return accObj
    for i,accChild in Acc_Children(accObj)
        if IsObject(accObj := %A_ThisFunc%(accChild))
            return accObj
}
; ------------------------------------------------------------------------------------------------------



; ------------------------------------------------------------------------------------------------------
; ------------------------------------------------------------------------------------------------------

Browser_GetUrl(WinId :="A"){

th:= WinExist(WinId)
WinGet, bp, ProcessName, ahk_id %th%
if ( bp ="vivaldi.exe") { ; Acc does not work for vivaldi
    Send ^l
    sUrl := Clip_GetSelection()
    Click
} Else {
    accData:= GetAccData(WinId)
    sUrl := accData.2
}
return sUrl

} ; eofun
; -------------------------------------c------------------------------------------------------------------------------




; https://www.autohotkey.com/boards/viewtopic.php?f=6&t=85383

GetAccData(WinId:="A") { ;by RRR based on atnbueno's https://www.autohotkey.com/boards/viewtopic.php?f=6&t=3702
    Static w:= [], n:=0
    th:= WinExist(WinId), GetKeyState("Ctrl", "P")? (w:= [], n:= 0): ""
    WinGet, bp, ProcessName, ahk_id %th%
    if (bp!="iexplore.exe" && bp!="vivaldi.exe")
        for i, v in w
            if (th=v.1)
                Return [ GetAccObjectFromWindow(v.1).accName(0), ParseAccData(v.4).2 ]
    Return ( [(tr:= ParseAccData(GetAccObjectFromWindow(th))).1, tr.2]
                  , tr.2? ( w[++n]:= [], w[n].1:=th, w[n].2:=tr.1, w[n].3:=tr.2, w[n].4:=tr.3): "" ) ; save Obj history
}

ParseAccData(accObj, accData:="") {
    try   accData? "": accData:= [accObj.accName(0)]
    try   if accObj.accRole(0) = 42 && accObj.accName(0) && accObj.accValue(0)
              accData.2:= SubStr(u:=accObj.accValue(0), 1, 4)="http"? u: "https://" u, accData.3:= accObj ; modern browser omit http:// TD change to https
          For nChild, accChild in GetAccChildren(accObj)
              accData.2? "": ParseAccData(accChild, accData)
          Return accData
}

GetAccInit() {
    Static hw:= DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
}

GetAccObjectFromWindow(hWnd, idObject = 0) {
	If DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr"
		, -VarSetCapacity(IID, 16)+NumPut(idObject==0xFFFFFFF0? 0x46000000000000C0: 0x719B3800AA000C81
		, NumPut(idObject==0xFFFFFFF0? 0x0000000000020400: 0x11CF3C3D618736E0, IID, "Int64"), "Int64"), "Ptr*", pacc)=0
		Return ComObjEnwrap(9, pacc, 1)
}

GetAccQuery(objAcc) {
	Try Return ComObj(9, ComObjQuery(objAcc, "{618736e0-3c3d-11cf-810c-00aa00389b71}"), 1)
}

GetAccChildren(objAcc) {
try	If ComObjType(objAcc,"Name") != "IAccessible"
		ErrorLevel := "Invalid IAccessible Object"
	Else {
		cChildren:= objAcc.accChildCount, Children:= []
		If DllCall("oleacc\AccessibleChildren", "Ptr", ComObjValue(objAcc), "Int", 0, "Int", cChildren, "Ptr"
		  , VarSetCapacity(varChildren, cChildren * (8 + 2 * A_PtrSize), 0) * 0 + &varChildren, "Int*", cChildren) = 0 {
			Loop, % cChildren {
				i:= (A_Index - 1) * (A_PtrSize * 2 + 8) + 8, child:= NumGet(varChildren, i)
				Children.Insert(NumGet(varChildren, i - 8) = 9? GetAccQuery(child): child)
                NumGet(varChildren, i - 8) != 9? "": ObjRelease(child)
			}   Return (Children.MaxIndex()? Children: "")
		}	Else ErrorLevel := "AccessibleChildren DllCall Failed"
    }
}
