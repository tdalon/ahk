; CustomBoxes https://www.autohotkey.com/boards/viewtopic.php?f=6&t=35382
; AutoResize: https://www.autohotkey.com/boards/viewtopic.php?style=17&t=1403
; -------------------------------------------------------------------------------

#Include <LV_EX>

ListView(Title := "", Prompt := "", List:="", ColumnTitle:="", bSelectAll := False, AlwaysOnTop := False) {
;-------------------------------------------------------------------------------
; show a custom input box with a ListBox control
; return the text of the selected items ',' separated 
;       "ListViewCancel" if Cancelled
;---------------------------------------------------------------------------
; Title is the title for the GUI
; Prompt is the text to display
; List is a pipe delimited list of choices or an array of object with properties 'name' and 'select'
; ColumnTitle is a pipe delimited list of column title
; Array of boolean the size of List elements or Boolean True of False for Select all

; create GUI
Gui, MyListView:New,,%Title%


If Prompt  ; not empty
    Gui, Add, Text,xm, %Prompt%

Gui, Add, ListView,  xm y20 r15 w700 h350 Checked HwndHLV, %ColumnTitle% ; Create a ListView.    
    
If IsObject(List) {
    Loop , % List.Length()
        {
            sName := List[A_Index]["name"]
            LV_Add("",Trim(List[A_Index]["name"]))
            If List[A_Index]["select"]
                LV_Modify(A_Index, "Check")
        }
} Else {
    Loop, Parse, List, |
        {  
            LV_Add("",Trim(A_LoopField))
        } ; End Loop 
}

LV_ModifyCol()  ; Auto-adjust the column widths.

; MenuBar
Menu, ItemsMenu, Add, Check All`tCtrl+A, LVCheckAll
Menu, ItemsMenu, Add, Uncheck all`tCtrl+Shift+A, LVUncheckAll
Menu, MyMenuBar, Add, &Items, :ItemsMenu 
Gui, Menu, MyMenuBar

If !IsObject(Select) and Select
    LV_Modify(0, "Check")

LV_Size := LV_EX_CalcViewSize(HLV, LV_GetCount())
GuiControl, Move, %HLV%, % "W" . LV_Size.W + 12 "H" . LV_Size.H +10 
;GuiControl, Move, HLV, w%W% ; h%H%

If (AlwaysOnTop = True)
    Gui, +AlwaysOnTop

/* 
Gui, Add, Button, x0  w50 h21 Default, OK
Gui, Add, Button, x55 w50 h21, Cancel 
*/

h := LV_Size.H + 40
Gui, Add, Button, xm w60 h21 y%h%  Default, &OK
Gui, Add, Button, x+m wp h21, &Cancel

Gui, Show, AutoSize

; main wait loop
Gui, +LastFound
WinWaitClose


return Selection


    ;-----------------------------------
    ; event handlers
    ;-----------------------------------
    MyListViewButtonOK: ; "OK" button, {Enter} pressed
        GoTo Submit
    return

    MyListViewButtonCancel: ; "Cancel" button
    MyListViewxGuiClose:     ; {Alt+F4} pressed, [X] clicked
    MyListViewGuiEscape:    ; {Esc} pressed
        Selection := "ListViewCancel"
        Gui, Destroy
    return

    LVCheckAll:
    LV_Modify(0, "Check")  ; Uncheck all the checkboxes.
    return
    LVUncheckAll:
    LV_Modify(0, "-Check")  ; Uncheck all the checkboxes.
    return

    Submit:
    RowNumber := 0  ; This causes the first loop iteration to start the search at the top of the list.

    If IsObject(List) {
        Selection := []
        Loop
            {
                RowNumber := LV_GetNext(RowNumber,"Checked")  ; Resume the search at the row after that found by the previous iteration.
                if not RowNumber  ; The above returned zero, so there are no more selected rows.
                    break
                Selection.Push(List[RowNumber])
            }
    } Else {
        Loop
            {
                RowNumber := LV_GetNext(RowNumber,"Checked")  ; Resume the search at the row after that found by the previous iteration.
                if not RowNumber  ; The above returned zero, so there are no more selected rows.
                    break
                
                LV_GetText(Text, RowNumber)
                If !Selection 
                    Selection := Text
                Else
                    Selection := Selection . "," . Text
            }
    }
    Gui, Destroy
    return

}


; ------------------------------------------
; ListView with Search box by wildcard on the top
; Selection := ListView_Select(LVArray,Title:="",Name)
; LVArray: can be multidimensional array, each row matching a ListView row
; Name : Column Names separated by |
; Search will filter first column content
ListView_Select(LVArray,Title:="", Name := "Name") {

static LVSSearchValue
static LVSListView
Gui, ListView_Select:New,,%Title%
Gui, Add, Text, ,Search:
Gui, Add, Edit, w400 vLVSSearchValue gLVSSearch
Gui, Add, ListView, grid w400  AltSubmit vLVSListView hwndHLVSListView gLVSListView, %Name%


If InStr(Name,"|") { ; multiple columns
    for m, row in LVArray {
        args := {}
		for n, col in row {
            args[n] := LVArray[m,n]
        }
            
    LV_Add("",args*)
    }	
} Else {
    For k,v In LVArray
        LV_Add("", v)
}

; https://www.autohotkey.com/boards/viewtopic.php?t=83495
LV_ModifyCol()  ; Auto-adjust the column widths.

Gui, Show

ListView_WantReturn(HLVSListView) ; <<< added 

; main wait loop
Gui, +LastFound
WinWaitClose

return Selection


LVSSearch:
Gui,Submit,NoHide
GuiControl, -Redraw, LV
LV_Delete()

If (LVSSearchValue = "")
    sPat := ".*"
Else {
    sPat := StrReplace(LVSSearchValue,"*",".*")
    If (SubStr(LVSSearchValue,1,1) != "*")
        sPat := "^" . sPat
}

If InStr(Name,"|") { ; multiple columns
    for m, row in LVArray {
        args := {}
		for n, col in row {
            args[n] := LVArray[m,n]
        }
        If RegExMatch(args[1], "i)" . sPat)         
            LV_Add("",args*)
    }	
} Else {
    For k,v In LVArray
        If RegExMatch(v, "i)" . sPat) ; ignore case
            LV_Add("", v)
}

GuiControl, +Redraw, LV

LV_Modify(1, "Select")
Return

LVSListView:
if (A_GuiEvent = "DoubleClick")
    {
        Selection:= A_EventInfo  ; Get the text from the row's first field.
        Gui, Destroy
        return
    }
;Gui, ListView, %A_GuiControl% ; <<< added
If (A_GuiEvent == "K") && (A_EventInfo = 13) ; VK_RETURN = 13 (0x0D)
{
   Selection:=LV_GetNext()
   Gui, Destroy
   return
}
return


ListView_SelectGuiClose:     ; {Alt+F4} pressed, [X] clicked
ListView_SelectGuiEscape:    ; {Esc} pressed
    Selection := 0
    Gui, Destroy
return
}


; ==================================================================================================================================
; N.B.: Requires to use ListView Option AltSubmit
; Source: just me https://www.autohotkey.com/boards/viewtopic.php?t=83495
; LV_WantReturn
;     'Fakes' Return key processing for ListView controls which otherwise won't process it.
;     If enabled, control's g-label will be triggered with A_GuiEvent = K and A_EventInfo = 13
;     whenever the <Return> key is pressed while the control has the focus.
; Usage:
;     To register a control call the functions once and pass the controls HWND as the first and only parameter.
;     To deregister it, call the function again with the same HWND as the first and only parameter.
; ==================================================================================================================================
ListView_WantReturn(wParam, lParam := "", Msg := "", HWND := "") {
    Static Controls := []
         , MsgFunc := Func("ListView_WantReturn")
         , OnMsg := False
         , LVN_KEYDOWN := -155
   ; Message handler call -----------------------------------------------------------------------------------------------------------
    If (Msg = 256) { ; WM_KEYDOWM (0x0100)
       If (wParam = 13) && (Ctl := Controls[HWND]) {
          If !(lParam & 0x40000000) { ; don't send notifications for auto-repeated keydown events
             VarSetCapacity(NMKD, (A_PtrSize * 3) + 8, 0) ; NMLVKEYDOWN/NMTVKEYDOWN structure 64-bit
             , NumPut(HWND, NMKD, 0, "Ptr")
             , NumPut(Ctl.CID, NMKD, A_PtrSize, "Ptr")
             , NumPut(LVN_KEYDOWN, NMKD, A_PtrSize * 2, "Int")
             , NumPut(13, NMKD, A_PtrSize * 3, "UShort")
             , DllCall("SendMessage", "Ptr", Ctl.HGUI, "UInt", 0x004E, "Ptr", Ctl.CID, "Ptr", &NMKD)
          }
          Return 0
       }
    }
    ; User call ---------------------------------------------------------------------------------------------------------------------
    Else {
       If (Controls[wParam += 0]) { ; the control is already registered, remove it
          Controls.Delete(wParam)
          If ((Controls.Length() = 0) && OnMsg) {
             OnMessage(0x0100, MsgFunc, 0)
             OnMsg := False
          }
          Return True
       }
       If !DllCall("IsWindow", "Ptr", wParam, "UInt")
          Return False
       WinGetClass, ClassName, ahk_id %wParam%
       If (ClassName <> "SysListView32")
          Return False
       Controls[wParam] := {CID:  DllCall("GetDlgCtrlID", "Ptr", wParam, "Int")
                          , HGUI: DllCall("GetParent", "Ptr", wParam, "UPtr")}
       If !(OnMsg)
          OnMessage(0x0100, MsgFunc, -1)
       Return (OnMsg := True)
    }
 }
