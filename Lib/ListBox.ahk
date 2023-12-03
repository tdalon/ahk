; CustomBoxes https://www.autohotkey.com/boards/viewtopic.php?f=6&t=35382
; AutoResize: https://www.autohotkey.com/boards/viewtopic.php?style=17&t=1403
; -------------------------------------------------------------------------------

ListBox(Title := "", Prompt := "", List := "", Select := 0, AlwaysOnTop := True,returnIndex := False ) {
; LB := ListBox(Title := "", Prompt := "", List := "", Select := 0, AlwaysOnTop := True)
    ;-------------------------------------------------------------------------------
; show a custom input box with a ListBox control
; return the index of the selected item. Empty if cancelled
;---------------------------------------------------------------------------
; Title is the title for the GUI
; Prompt is the text to display
; List is a pipe delimited list of choices
; Select (if present) is the index of the preselected item. Default is 0 for no selection

static LB, OK, Cancel ; used as a GUI control variable

; create GUI
Gui, ListBox: New, ,%Title%
Gui, -MinimizeBox
;Gui, Margin, 30, 18
If Prompt ; not empty
    Gui, Add, Text,, %Prompt%

Loop Parse, List,| 
{
    rCnt++
    maxsize := StrLen(A_LoopField) > maxsize ? StrLen(A_LoopField) : maxsize
}
boxwidth := maxsize * 5  
If returnIndex
    Gui, Add, ListBox, +AltSubmit w%boxwidth% r%rCnt% vLB hwndHLB Choose%Select%, %List%
Else  
    Gui, Add, ListBox, w%boxwidth% r%rCnt% vLB hwndHLB Choose%Select%, %List%

;W := LB_EX_CalcWidth(HLB)
;GuiControl, Move, HLB, w%W% ; h%H%

Gui, Add, Button, w60 Default, &OK
Gui, Add, Button, x+m wp, &Cancel

;GuiControl, Move, OK, y%H%
;GuiControl, Move, Cancel, % "y" H

If (AlwaysOnTop = True)
    Gui, +AlwaysOnTop

Gui, Show , AutoSize

; main wait loop
Gui, +LastFound
WinWaitClose

return LB


    ;-----------------------------------
    ; event handlers
    ;-----------------------------------
    ListBoxButtonOK: ; "OK" button, {Enter} pressed
        Gui, Submit ; get Result from GUI
        Gui, Destroy
    return

    ListBoxButtonCancel: ; "Cancel" button
    ListBoxGuiClose:     ; {Alt+F4} pressed, [X] clicked
    ListBoxGuiEscape:    ; {Esc} pressed
        ;LB := "ListBoxCancel"
        LB =
        Gui, Destroy
    return
}



LB_EX_CalcWidth(HLB) { ; calculates the width of the list box needed to show the whole content.
   ; HLB - Handle to the ListBox.
   MaxW := 0
   ControlGet, Items, List, , , % "ahk_id " . HLB
   SendMessage, 0x0031, 0, 0, , % "ahk_id " . HLB ; WM_GETFONT
   HFONT := ErrorLevel
   HDC := DllCall("User32.dll\GetDC", "Ptr", HLB, "UPtr")
   DllCall("Gdi32.dll\SelectObject", "Ptr", HDC, "Ptr", HFONT)
   VarSetCapacity(SIZE, 8, 0)
   Loop, Parse, Items, `n
   {
      Txt := A_LoopField
      DllCall("Gdi32.dll\GetTextExtentPoint32", "Ptr", HDC, "Ptr", &Txt, "Int", StrLen(Txt), "UIntP", Width)
      If (Width > MaxW)
        MaxW := Width
   }
   DllCall("User32.dll\ReleaseDC", "Ptr", HLB, "Ptr", HDC)
   Return MaxW + 8 ; + 8 for the margins
}