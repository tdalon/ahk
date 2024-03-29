; Browser Library
Browser_WinActive(){   
    GroupAdd, Browser, ahk_exe iexplore.exe
    GroupAdd, Browser, ahk_exe chrome.exe
    GroupAdd, Browser, ahk_exe sidekick.exe
    GroupAdd, Browser, ahk_exe firefox.exe
    GroupAdd, Browser, ahk_exe palemoon.exe
    GroupAdd, Browser, ahk_exe waterfox.exe
    GroupAdd, Browser, ahk_exe vivaldi.exe
    GroupAdd, Browser, ahk_exe msedge.exe	
    return WinActive("ahk_group Browser")
} ; eofun

Browser_IsWinActive(){   
    GroupAdd, Browser, ahk_exe iexplore.exe
    GroupAdd, Browser, ahk_exe chrome.exe
    GroupAdd, Browser, ahk_exe sidekick.exe
    GroupAdd, Browser, ahk_exe firefox.exe
    GroupAdd, Browser, ahk_exe palemoon.exe
    GroupAdd, Browser, ahk_exe waterfox.exe
    GroupAdd, Browser, ahk_exe vivaldi.exe
    GroupAdd, Browser, ahk_exe msedge.exe	
    return (WinActive("ahk_group Browser") != 0)
} ; eofun
    
Browser_WinWait(Timeout:=""){     
    GroupAdd, Browser, ahk_exe iexplore.exe
    GroupAdd, Browser, ahk_exe chrome.exe
    GroupAdd, Browser, ahk_exe sidekick.exe
    GroupAdd, Browser, ahk_exe firefox.exe
    GroupAdd, Browser, ahk_exe palemoon.exe
    GroupAdd, Browser, ahk_exe waterfox.exe
    GroupAdd, Browser, ahk_exe vivaldi.exe
    GroupAdd, Browser, ahk_exe msedge.exe	 
    WinWait, ahk_group Browser,,%Timeout%
} ; eofun
    
Browser_WinWaitActive(Timeout:=""){     
    GroupAdd, Browser, ahk_exe iexplore.exe
    GroupAdd, Browser, ahk_exe chrome.exe
    GroupAdd, Browser, ahk_exe sidekick.exe
    GroupAdd, Browser, ahk_exe firefox.exe
    GroupAdd, Browser, ahk_exe palemoon.exe
    GroupAdd, Browser, ahk_exe waterfox.exe
    GroupAdd, Browser, ahk_exe vivaldi.exe
    GroupAdd, Browser, ahk_exe msedge.exe	 
    WinWaitActive, ahk_group Browser,,%Timeout%
} ; eofun
    
Browser_WinExist(){
    GroupAdd, Browser, ahk_exe iexplore.exe
    GroupAdd, Browser, ahk_exe chrome.exe
    GroupAdd, Browser, ahk_exe sidekick.exe
    GroupAdd, Browser, ahk_exe firefox.exe
    GroupAdd, Browser, ahk_exe palemoon.exe
    GroupAdd, Browser, ahk_exe waterfox.exe
    GroupAdd, Browser, ahk_exe vivaldi.exe
    GroupAdd, Browser, ahk_exe msedge.exe	
    return WinExist("ahk_group Browser")
}
    
    
; -----------------------------------------------------------------------------------------------------------------

Browser_GetUrl(wTitle := "A"){
    sUrl:=GetUrl(wTitle)
    return sUrl
}


Browser_GetUrl_Uia(wTitle:="A") {
; Method based on UIAutomation ; From Descolada https://www.autohotkey.com/boards/viewtopic.php?f=6&t=3702&e=1&view=unread#p459451
; does not require Acc Lib
; Is bound to computer language settings i.e. English "Address and search bar"
    ErrorLevel := 0
    if !(wId := WinExist(wTitle)) {
        ErrorLevel := 1
        return
    }
    IUIAutomation := ComObjCreate(CLSID_CUIAutomation := "{ff48dba4-60ef-4201-aa87-54103eef594e}", IID_IUIAutomation := "{30cbe57d-d9d0-452a-ab13-7ac5ac4825ee}")
    DllCall(NumGet(NumGet(IUIAutomation+0)+6*A_PtrSize), "ptr", IUIAutomation, "ptr", wId, "ptr*", elementMain)   ; IUIAutomation::ElementFromHandle
    NumPut(addressbarStrPtr := DllCall("oleaut32\SysAllocString", "wstr", "Address and search bar", "ptr"),(VarSetCapacity(addressbar,8+2*A_PtrSize)+NumPut(8,addressbar,0,"short"))*0+&addressbar,8,"ptr")
    DllCall("oleaut32\SysFreeString", "ptr", addressbarStrPtr)
    if (A_PtrSize = 4) {
        DllCall(NumGet(NumGet(IUIAutomation+0)+23*A_PtrSize), "ptr", IUIAutomation, "int", 30005, "int64", NumGet(addressbar, 0, "int64"), "int64", NumGet(addressbar, 8, "int64"), "ptr*", addressbarCondition)   ; IUIAutomation::CreatePropertyCondition
    } else {
        DllCall(NumGet(NumGet(IUIAutomation+0)+23*A_PtrSize), "ptr", IUIAutomation, "int", 30005, "ptr", &addressbar, "ptr*", addressbarCondition)   ; IUIAutomation::CreatePropertyCondition
    }
    DllCall(NumGet(NumGet(elementMain+0)+5*A_PtrSize), "ptr", elementMain, "int", TreeScope_Descendants := 0x4, "ptr", addressbarCondition, "ptr*", currentURLElement) ; IUIAutomationElement::FindFirst
    DllCall(NumGet(NumGet(currentURLElement+0)+10*A_PtrSize),"ptr",currentURLElement,"uint",30045,"ptr",(VarSetCapacity(currentURL,8+2*A_PtrSize)+NumPut(0,currentURL,0,"short")+NumPut(0,currentURL,8,"ptr"))*0+&currentURL) ;IUIAutomationElement::GetCurrentPropertyValue
    ObjRelease(currentURLElement)
    ObjRelease(elementMain)
    ObjRelease(IUIAutomation)
    sUrl := StrGet(NumGet(currentURL,8,"ptr"),"utf-16")
    If !sUrl { ;empty
        ErrorLevel := 1
        return
    }

    If !RegExMatch(sUrl,"^https?://")
        sUrl := "https://" . sUrl
    return sUrl
}

; ------------------------------------------------------------------------------------------------------
; ------------------------------------------------------------------------------------------------------

Browser_GetUrl0(WinId :="A"){

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
; Version 6.45 ; modification marked with TD

GetAccData(WinId:="A") { ;by RRR based on atnbueno's https://www.autohotkey.com/boards/viewtopic.php?f=6&t=3702
    Static w:= [], n:=0
    Global vtr1, vtr2
    th:= WinExist(WinId), GetKeyState("Ctrl", "P")? (w:= [], n:= 0): ""
    ;WinGet, bp, ProcessName, ahk_id %th% ; TD not used/ commented out
    for i, v in w
        if (th=v.1)
            Return [ GetAccObjectFromWindow(v.1).accName(0), ParseAccData(v.4).2 ]
    Return ( [(tr:= ParseAccData(GetAccObjectFromWindow(th))).1, tr.2]
                , tr.2? ( w[++n]:= [], w[n].1:=th, w[n].2:=tr.1, w[n].3:=tr.2, w[n].4:=tr.3): "" ) ; save AccObj history
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
    SendMessage, WM_GETOBJECT := 0x003D, 0, 1, Chrome_RenderWidgetHostHWND1, % "ahk_id " WinExist("A") ; by malcev
    While DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr"
        , -VarSetCapacity(IID, 16)+NumPut(idObject==0xFFFFFFF0? 0x46000000000000C0: 0x719B3800AA000C81
        , NumPut(idObject==0xFFFFFFF0? 0x0000000000020400: 0x11CF3C3D618736E0, IID, "Int64"), "Int64"), "Ptr*", pacc)!=0
        && A_Index < 60
        sleep, 30
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



; ------------------------------------------------------------------------------------------------------
; ------------------------------------------------------------------------------------------------------
;https://www.reddit.com/r/AutoHotkey/comments/m0i7tf/mind_helping_me_on_chaining_this_command/gvgwhie
Browser_GetUrl2(hWnd := "A")
; Requires Acc Library
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
        If !RegExMatch(sUrl,"^https?://.*")
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












; Version: 2022.07.06.1
; https://gist.github.com/7cce378c9dfdaf733cb3ca6df345b140
GetUrl_Acc(WinTitle*)
{
    active := WinExist("A")
    if !(hWnd := WinExist(WinTitle*))
        return
    objId := -4
    WinGetClass wClass
    if (wClass ~= "Chrome") {
        WinGet pid, PID
        hWnd := WinExist("ahk_pid" pid)
        if (active != hWnd)
            objId := 0
    }
    oAcc := Acc_ObjectFromWindow(hWnd, objId)
    if (wClass ~= "Chrome") {
        try {
            SendMessage 0x003D, 0, 1, Chrome_RenderWidgetHostHWND1
            oAcc.accName(0)
        }
    }
    if (oAcc := GetUrl_Recurse(oAcc))
        return oAcc.accValue(0)
}

GetUrl_Recurse(oAcc)
{
    if (ComObjType(oAcc, "Name") != "IAccessible")
        return
    if (oAcc.accValue(0) ~= "^[\w-]+:")
        return oAcc
    for _,accChild in Acc_Children(oAcc) {
        oAcc := GetUrl_Recurse(accChild)
        if (IsObject(oAcc))
            return oAcc
    }
}


; Version: 2023.07.06.1
; https://gist.github.com/7cce378c9dfdaf733cb3ca6df345b140

GetUrl(WinTitle*) {
    static S_OK := 0, UIA_DocumentControlTypeId := 50030, UIA_EditControlTypeId := 50004, conditions := [], UIA_ControlTypePropertyId := 30003, TreeScope_Descendants := 4, UIA_ValueValuePropertyId := 30045
    hWnd := WinExist(WinTitle*)
    WinGetClass winClass
    IUIAutomation := eRoot := condition := eFirst := 0
    try {
        IUIAutomation := ComObjCreate("{ff48dba4-60ef-4201-aa87-54103eef594e}", "{30cbe57d-d9d0-452a-ab13-7ac5ac4825ee}") ; CLSID_CUIAutomation, IID_IUIAutomation
        HRESULT := DllCall(NumGet(NumGet(IUIAutomation + 0) + 6 * A_PtrSize), "Ptr", IUIAutomation, "Ptr", hWnd, "Ptr*", eRoot)
        if (HRESULT != S_OK) {
            throw Exception("IUIAutomation::ElementFromHandle`n`nHRESULT: " HRESULT)
        }
        ctrlTypeId := winClass ~= "Chrome" ? UIA_DocumentControlTypeId : UIA_EditControlTypeId
        if (conditions.HasKey(ctrlTypeId)) {
            condition := conditions[ctrlTypeId]
        } else {
            VarSetCapacity(value, 8 + 2 * A_PtrSize, 0)
            NumPut(3, value, 0, "UShort")
            NumPut(ctrlTypeId, value, 8, "Ptr")
            if (A_PtrSize = 8) {
                HRESULT := DllCall(NumGet(NumGet(IUIAutomation + 0) + 23 * A_PtrSize), "Ptr", IUIAutomation, "UInt", UIA_ControlTypePropertyId, "Ptr", &value, "Ptr*", condition)
            } else {
                HRESULT := DllCall(NumGet(NumGet(IUIAutomation + 0) + 23 * A_PtrSize), "Ptr", IUIAutomation, "UInt", UIA_ControlTypePropertyId, "UInt64", NumGet(value, 0, "UInt64"), "UInt64", NumGet(value, 8, "UInt64"), "Ptr*", condition)
            }
            if (HRESULT != S_OK) {
                throw Exception("IUIAutomation::CreatePropertyCondition`n`nHRESULT: " HRESULT)
            }
            conditions[ctrlTypeId] := condition
        }
        HRESULT := DllCall(NumGet(NumGet(eRoot + 0) + 5 * A_PtrSize), "Ptr", eRoot, "UInt", TreeScope_Descendants, "Ptr", condition, "Ptr*", eFirst)
        if (HRESULT != S_OK) {
            throw Exception("IUIAutomationElement::FindFirst`n`nHRESULT: " HRESULT)
        }
        VarSetCapacity(propertyValue, 8 + 2 * A_PtrSize, 0)
        NumPut(0, propertyValue, 0, "UShort")
        NumPut(0, propertyValue, 8, "Ptr")
        HRESULT := DllCall(NumGet(NumGet(eFirst + 0) + 10 * A_PtrSize), "Ptr", eFirst, "UInt", UIA_ValueValuePropertyId, "Ptr", &propertyValue)
        if (HRESULT != S_OK) {
            throw Exception("IUIAutomationElement::GetCurrentPropertyValue`n`nHRESULT: " HRESULT)
        }
        return StrGet(NumGet(propertyValue, 8), "UTF-16")
    } catch e {
        throw Exception("GetUrl() failed.", -1, e.Message)
    } finally {
        if (eFirst) {
            ObjRelease(eFirst)
        }
        if (eRoot) {
            ObjRelease(eRoot)
        }
        if (IUIAutomation) {
            ObjRelease(IUIAutomation)
        }
    }
}


; #### Not used
Browser_GetPageSource(){
    ; Window will flash. Alternative use HttpGet but requires password setting for Connections
    Send ^u ; Open code view via Ctrl+U Hotkey (Chrome)
    ; Copy all source code to clipboard
    ClipSaved := ClipboardAll
    Clipboard = 
    sleep, 300
    Send ^a
    sleep, 300
    Send ^c
    ClipWait, 0.5
    sHtml := Clipboard
    Send ^w
    ; Restore Clipboard
    Clipboard := ClipSaved
    return sHtml
}

Browser_GetPage(sUrl){
; Syntax:  sHtml := BrowserGetPage(sUrl*)
; Called by Connections->CNGetAtom and ConnectionsEnhancer
; If no input Url, copy current page
ClipboardSaved := ClipboardAll

If sUrl {
    Clipboard := sUrl
    Send ^t ; Open new tab
    Send ^v
    Send {Enter}
    Sleep 3000 ; time to load the page
    ; TODO: optimize timing
    ; https://jacksautohotkeyblog.wordpress.com/2018/05/02/waiting-for-a-web-page-to-load-into-a-browser-autohotkey-tips/
}

;SendInput %sUrl%{Enter}

Clipboard =
Send ^a
Sleep 500
Send ^c

ClipWait, 5 ; Wait max 5 seconds
If ErrorLevel
{
    MsgBox, The attempt to copy text onto the clipboard failed.
    Return
}
    
sHtml := Clipboard
If sUrl {
    Sleep 500
    Send ^w ; close window 
}
Else
    Click ; Deselect

Clipboard := ClipboardSaved
return sHtml

}