; -------------------------------------------------------------------------------------------------------------------
; -------------------------------------------------------------------------------------------------------------------
; https://stackoverflow.com/a/25800045/2043349 - modified
MultiLineInputBox(Text:="", Default:="", Caption:="Multi Line Input Box"){
    static
    ButtonOK:=ButtonCancel:= false
	Gui GuiMLIB:New,, % Caption
    Gui, add, Text, w600, % Text
    Gui, add, Edit, r10 w600 vMLIBEdit, % Default
    Gui, add, Button, w60 gMLIBOK , &OK
    Gui, add, Button, w60 x+10 gMLIBCancel, &Cancel

	Gui, Show
    while !(ButtonOK||ButtonCancel)
        continue
    if ButtonCancel {
        ErrorLevel := 1
        return
    }
    Gui, Submit
    ErrorLevel := 0
    return MLIBEdit
    ;----------------------
    MLIBOK:
    ButtonOK:= true
    return
    ;---------------------- 
    GuiMLIBGuiEscape:
	GuiMLIBGuiClose:
    MLIBCancel:
    ButtonCancel:= true
    
    Gui, Cancel
    return
}
