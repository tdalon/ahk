; Version: 2022.07.01.1 by anonymous1184
; https://gist.github.com/58d2b141be2608a2f7d03a982e552a71

; Private
Acc_Init(Function) {
    static hModule := DllCall("Kernel32\LoadLibrary", "Str","oleacc.dll", "Ptr")
    return DllCall("Kernel32\GetProcAddress", "Ptr",hModule, "AStr",Function, "Ptr")
}

Acc_ObjectFromEvent(ByRef ChildIdOut, hWnd, ObjectId, ChildId) {
    static address := Acc_Init("AccessibleObjectFromEvent")
    VarSetCapacity(varChild, A_PtrSize * 2 + 8, 0)
    hResult := DllCall(address, "Ptr",hWnd, "UInt",ObjectId, "UInt",ChildId
        , "Ptr*",pAcc:=0, "Ptr",&varChild)
    if (!hResult) {
        ChildIdOut := NumGet(varChild, 8, "UInt")
        return ComObj(9, pAcc, 1)
    }
}

Acc_ObjectFromPoint(ByRef ChildIdOut := "", x := 0, y := 0) {
    static address := Acc_Init("AccessibleObjectFromPoint")
    try
        point := x & 0xFFFFFFFF | y << 32
    catch
        DllCall("User32\GetCursorPos", "Int64*",point:=0)
    VarSetCapacity(varChild, A_PtrSize * 2 + 8, 0)
    hResult := DllCall(address, "Int64",point, "Ptr*",pAcc:=0, "Ptr",&varChild)
    if (!hResult) {
        ChildIdOut := NumGet(varChild, 8, "UInt")
        return ComObj(9, pAcc, 1)
    }
}

/* ObjectId
    0xFFFFFFFF = OBJID_SYSMENU
    0xFFFFFFFE = OBJID_TITLEBAR
    0xFFFFFFFD = OBJID_MENU
    0xFFFFFFFC = OBJID_CLIENT*
    0xFFFFFFFB = OBJID_VSCROLL
    0xFFFFFFFA = OBJID_HSCROLL
    0xFFFFFFF9 = OBJID_SIZEGRIP
    0xFFFFFFF8 = OBJID_CARET
    0xFFFFFFF7 = OBJID_CURSOR
    0xFFFFFFF6 = OBJID_ALERT
    0xFFFFFFF5 = OBJID_SOUND
    0xFFFFFFF4 = OBJID_QUERYCLASSNAMEIDX
    0xFFFFFFF0 = OBJID_NATIVEOM
    0x00000000 = OBJID_WINDOW
*/
Acc_ObjectFromWindow(hWnd, ObjectId := -4) {
    static address := Acc_Init("AccessibleObjectFromWindow")
    ObjectId &= 0xFFFFFFFF
    VarSetCapacity(IID, 16, 0)
    addr := ObjectId = 0xFFFFFFF0 ? 0x0000000000020400 : 0x11CF3C3D618736E0
    rIID := NumPut(addr, IID, "Int64")
    addr := ObjectId = 0xFFFFFFF0 ? 0x46000000000000C0 : 0x719B3800AA000C81
    rIID := NumPut(addr, rIID + 0, "Int64") - 16
    hResult := DllCall(address, "Ptr",hWnd, "UInt",ObjectId, "Ptr",rIID, "Ptr*",pAcc:=0)
    if (!hResult)
        return ComObj(9, pAcc, 1)
}

Acc_WindowFromObject(pAcc) {
    static address := Acc_Init("WindowFromAccessibleObject")
    if (IsObject(pAcc))
        pAcc := ComObjValue(pAcc)
    hResult := DllCall(address, "Ptr",pAcc, "Ptr*",hWnd:=0)
    if (!hResult)
        return hWnd
}

Acc_GetRoleText(nRole) {
    static address := Acc_Init("GetRoleTextW")
    nSize := DllCall(address, "UInt",nRole, "Ptr",0, "UInt",0)
    VarSetCapacity(sRole, ++nSize, 0)
    DllCall(address, "UInt",nRole, "Ptr",&sRole, "UInt",nSize)
    return StrGet(&sRole)
}

Acc_GetStateText(nState) {
    static address := Acc_Init("GetStateTextW")
    nSize := DllCall(address, "UInt",nState, "Ptr",0, "UInt",0)
    VarSetCapacity(sState, ++nSize, 0)
    DllCall(address, "UInt",nState, "Ptr",&sState, "UInt",nSize)
    return StrGet(&sState)
}

Acc_SetWinEventHook(EventMin, EventMax, Callback) {
    return DllCall("User32\SetWinEventHook", "UInt",EventMin, "UInt",EventMax
        , "Ptr",0, "Ptr",Callback, "UInt",0, "UInt",0, "UInt",0)
}

Acc_UnhookWinEvent(hHook) {
    return DllCall("User32\UnhookWinEvent", "Ptr",hHook)
}

/* Win Events
Callback := RegisterCallback("WinEventProc")
WinEventProc(hHook, Event, hWnd, ObjectId, ChildId, EventThread, EventTime) {
    Critical
    oAcc := Acc_ObjectFromEvent(ChildIdOut, hWnd, ObjectId, ChildId)
    ; Code Here
}
*/

; Written by jethrow

Acc_Role(oAcc, ChildId := 0) {
    try {
        role := oAcc.accRole(ChildId + 0)
        return Acc_GetRoleText(role)
    }
    return "invalid object"
}

Acc_State(oAcc, ChildId := 0) {
    try {
        state := oAcc.accState(ChildId + 0)
        return Acc_GetStateText(state)
    }
    return "invalid object"
}

Acc_Location(oAcc, ChildId := 0, ByRef Position := "") {
    x := 0, w := 0
    y := 0, h := 0
    ; 0x4003 = VT_BYREF | VT_I4
    x := ComObject(0x4003, &x), w := ComObject(0x4003, &w)
    y := ComObject(0x4003, &y), h := ComObject(0x4003, &h)
    try {
        oAcc.accLocation(x, y, w, h, ChildId + 0)
        x := NumGet(x, 0, "Int"), w := NumGet(w, 0, "Int")
        y := NumGet(y, 0, "Int"), h := NumGet(h, 0, "Int")
        Position := "x" x " y" y " w" w " h" h
        return {"x":x, "y":y, "w":w, "h":h, "pos":Position}
    }
}

Acc_Parent(oAcc) {
    try {
        if (oAcc.accParent)
            return Acc_Query(oAcc.accParent)
    }
}

Acc_Child(oAcc, ChildId := 0) {
    try {
        Child := oAcc.AccChild(ChildId + 0)
        if (Child)
            return Acc_Query(Child)
    }
}

; Private
Acc_Query(oAcc) {
    try {
        query := ComObjQuery(oAcc, "{618736E0-3C3D-11CF-810C-00AA00389B71}")
        return ComObj(9, query, 1)
    }
} ; Thanks Lexikos - www.autohotkey.com/forum/viewtopic.php?t=81731&p=509530#509530

; Private, deprecated
Acc_Error(Previous := "") {
    static setting := 0
    if (StrLen(Previous))
        setting := Previous
    return setting
}

Acc_Children(oAcc) {
    static address := Acc_Init("AccessibleChildren")
    if (ComObjType(oAcc, "Name") != "IAccessible")
        throw Exception("Invalid IAccessible Object", -1, oAcc)
    pAcc := ComObjValue(oAcc)
    size := A_PtrSize * 2 + 8
    VarSetCapacity(varChildren, oAcc.accChildCount * size, 0)
    hResult := DllCall(address, "Ptr",pAcc, "Int",0, "Int",oAcc.accChildCount
        , "Ptr",&varChildren, "Int*",obtained:=0)
    if (hResult)
        throw Exception("AccessibleChildren DllCall Failed", -1)
    children := []
    loop % obtained {
        i := (A_Index - 1) * size
        child := NumGet(varChildren, i + 8, "Int64")
        if (NumGet(varChildren, i, "Int64") = 9) {
            accChild := Acc_Query(child)
            children.Push(accChild)
            ObjRelease(child)
        } else {
            children.Push(child)
        }
    }
    return children
}

Acc_ChildrenByRole(oAcc, RoleText) {
    children := []
    for _,child in Acc_Children(oAcc) {
        if (Acc_Role(child) = RoleText)
            children.Push(child)
    }
    return children
}

/* Commands
    - Aliases:
    Action → DefaultAction
    DoAction → DoDefaultAction
    Keyboard → KeyboardShortcut
    - Properties:
    Child
    ChildCount
    DefaultAction
    Description
    Focus
    Help
    HelpTopic
    KeyboardShortcut
    Name
    Parent
    Role
    Selection
    State
    Value
    - Methods:
    DoDefaultAction
    Location
    - Other:
    Object
*/
Acc_Get(Command, ChildPath, ChildId := 0, Target*) {
    if (Command ~= "i)^(?:HitTest|Navigate|Select)$")
        throw Exception("Command not implemented", -1, Command)
    ChildPath := StrReplace(ChildPath, "_", " ")
    if (IsObject(Target[1])) {
        oAcc := Target[1]
    } else {
        hWnd := WinExist(Target*)
        oAcc := Acc_ObjectFromWindow(hWnd, 0)
    }
    if (ComObjType(oAcc, "Name") != "IAccessible")
        throw Exception("Cannot access an IAccessible Object", -1, oAcc)
    ChildPath := StrSplit(ChildPath, ".")
    for level,item in ChildPath {
        RegExMatch(item, "OS)(?<Role>\D+)(?<Index>\d*)", match)
        if (match) {
            item := match.Index ? match.Index : 1
            children := Acc_ChildrenByRole(oAcc, match.Role)
        } else {
            children := Acc_Children(oAcc)
        }
        if (children.HasKey(item)) {
            oAcc := children[item]
            continue
        }
        extra := match.Role
            ? "Role: " match.Role ", Index: " item
            : "Item: " item ", Level: " level
        throw Exception("Cannot access ChildPath Item", -1, extra)
    }
    switch (Command) {
        case "Action": Command := "DefaultAction"
        case "DoAction": Command := "DoDefaultAction"
        case "Keyboard": Command := "KeyboardShortcut"
        case "Object":
            return oAcc
    }
    switch (Command) {
        case "Location":
            out := Acc_Location(oAcc, ChildId).pos
        case "Parent":
            out := Acc_Parent(oAcc)
        case "Role", "State":
            out := Acc_%Command%(oAcc, ChildId)
        case "ChildCount", "Focus", "Selection":
            out := oAcc["acc" Command]
        default:
            out := oAcc["acc" Command](ChildId + 0)
    }
    return out
}