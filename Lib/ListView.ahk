; CustomBoxes https://www.autohotkey.com/boards/viewtopic.php?f=6&t=35382
; AutoResize: https://www.autohotkey.com/boards/viewtopic.php?style=17&t=1403
; -------------------------------------------------------------------------------

#Include <LV_EX>

ListView(Title := "", Prompt := "", List:="", ColumnTitle:="", bSelectAll := True, AlwaysOnTop := False) {
;-------------------------------------------------------------------------------
; show a custom input box with a ListBox control
; return the text of the selected items , separated "ListViewCancel" if Cancelled
;---------------------------------------------------------------------------
; Title is the title for the GUI
; Prompt is the text to display
; List is a pipe delimited list of choices
; ColumnTitle is a pipe delimited list of column title
; bSelectAll If True all list elements are selected/ checked

static LB ; used as a GUI control variable

; create GUI
Gui, MyListView:New,,%Title%

Gui, Add, Button, x0 y0 w50 h21 Default, OK
Gui, Add, Button, x55 y0 w50 h21, Cancel


If Prompt { ; not empty
    Gui, Add, Text,x0, %Prompt%
    Gui, Add, ListView,  x0 y50 r15 w700 h350 Checked HwndHLV, %ColumnTitle% ; Create a ListView.
} Else
    ; ListView
    Gui, Add, ListView,  x0 y25 r15 w700 h350 Checked HwndHLV, %ColumnTitle% ; Create a ListView.

Loop, Parse, List, |
{  
    LV_Add("",Trim(A_LoopField))
} ; End Loop     

LV_ModifyCol()  ; Auto-adjust the column widths.

; MenuBar
Menu, ItemsMenu, Add, Check All`tCtrl+A, LVCheckAll
Menu, ItemsMenu, Add, Uncheck all`tCtrl+Shift+A, LVUncheckAll
Menu, MyMenuBar, Add, &Items, :ItemsMenu 
Gui, Menu, MyMenuBar

If bSelectAll
    LV_Modify(0, "Check")

LV_Size := LV_EX_CalcViewSize(HLV, LV_GetCount())
GuiControl, Move, %HLV%, % "W" . LV_Size.W + 12 "H" . LV_Size.H +10 


If (AlwaysOnTop = True)
    Gui, +AlwaysOnTop

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
    Loop
    {
        RowNumber := LV_GetNext(RowNumber,"Checked")  ; Resume the search at the row after that found by the previous iteration.
        if not RowNumber  ; The above returned zero, so there are no more selected rows.
            break
        LV_GetText(Text, RowNumber)
        If !Selection 
            Selection := Text
        Else
            Selection := Selection . ", " . Text
    }
    Gui, Destroy
    return

}



