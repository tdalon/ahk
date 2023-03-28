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

static LB ; used as a GUI control variable

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
;W := LB_EX_CalcWidth(HLV)
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



