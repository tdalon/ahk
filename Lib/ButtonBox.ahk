; CustomBoxes https://www.autohotkey.com/boards/viewtopic.php?f=6&t=35382
;-------------------------------------------------------------------------------
ButtonBox(Title := "", Prompt := "", List := "", Seconds := "",Def:=1,AlwaysOnTop:=True) {
;-------------------------------------------------------------------------------
    ; show a custom MsgBox with arbitrarily named buttons
    ; return the text of the button pressed. 
    ; If Timeout returns "Timeout". If Cancelled or Closed, returns "ButtonBox_Cancel"
    ;---------------------------------------------------------------------------
    ; Title is the title for the GUI
    ; Prompt is the text to display
    ; List is a pipe delimited list of captions for the buttons
    ; Seconds is the time in seconds to wait before timing out. Leave blank to wait indefinitely
    ; Def Number: default button selected. Default=1 for first button
    ; AlwaysOnTop. Default True. Gui is set AlwaysOnTop

    ; create GUI
    Gui, ButtonBox: New,, %Title%
    Gui, -MinimizeBox
    Gui, Margin, 30, 18
    Gui, Add, Text,, %Prompt%
    If (AlwaysOnTop=True)
        Gui,+AlwaysOnTop
    Loop, Parse, List, | 
    {  
        If (A_Index = Def)
            Gui, Add, Button, % (A_Index = 1 ? "" : "x+10") " gBtn Default" , %A_LoopField%
        Else
            Gui, Add, Button, % (A_Index = 1 ? "" : "x+10") " gBtn", %A_LoopField%
    }
    Gui, Show

    ; main wait loop
    Gui, +LastFound
    WinWaitClose,,, %Seconds%

    if (ErrorLevel = 1) {
        Result := "TimeOut"
        Gui, Destroy
    }

return Result


    ;-----------------------------------
    ; event handlers
    ;-----------------------------------
    Btn: ; all the buttons come here
        Result := A_GuiControl
        Gui, Destroy
    return

    ButtonBoxGuiClose:  ; {Alt+F4} pressed, [X] clicked
    ButtonBoxGuiEscape: ; {Esc} pressed
        Result := "ButtonBox_Cancel"
        Gui, Destroy
    return
}
