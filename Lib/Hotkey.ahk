; Source: https://autohotkey.com/board/topic/15577-function-hotkeygui-v04/
; Code: https://ahknet.autohotkey.com/~jballi/HotkeyGUI/v0.4/HotkeyGUI.ahk
; Documentation: https://ahknet.autohotkey.com/~jballi/HotkeyGUI/v0.4/HotkeyGUI.html -> dead link
;-----------------------------
;
; Function: HotkeyGUI
;
; Description:
;
;   This function displays a dialog that will allow the user to select a hotkey
;   without using the keyboard.  See the "Processing and Usage Notes" section
;   for more information.
;
; HG_HotKey := HotkeyGUI(Owner,Hotkey,Limit,OptionalAttrib,Title)
;
; Parameters:
;
;   p_Owner - The GUI owner of the HotkeyGUI window. [Optional] The default is
;       0 (no owner).  If not defined, the AlwaysOnTop attribute is added to the
;       HotkeyGUI window to make sure that the window is not lost.
;
;   p_Hotkey - The default hotkey value. [Optional] The default is blank.  To
;       only preselect modifiers and/or optional attributes, enter only the
;       associated characters.  For example, to only have the Ctrl and Shift
;       modifiers set as the default, enter "^+".
;
;   p_Limit - Hotkey limit. [Optional] The default is 0.  See the "Hotkey
;       Limits" section below for more information.
;
;   p_OptionalAttrib - Optional hotkey attributes. [Optional]  The default is
;       FALSE.  If set to TRUE, all items in the Optional Attributes group are
;       enabled and the user is allowed to add or remove optional attributes.
;
;   p_Title - Window title. [Optional]  The default is the current script name
;       (sans the extention) plus "Select Hotkey".
;
;
; Processing And Usage Notes:
;
; Stuff to know...
;
;    *  The function does not return until the HotkeyGUI window is closed
;       (Accept or Cancel).
;
;    *  A shift-only key (Ex: ~!@#$%^&*()_+{}|:"<>?) cannot be directly selected
;       as a key by this function.  To use a shift-only key, select the Shift
;       modifier and then select the non-shift version of the key.  For example,
;       to set the "(" key as a hotkey, select the Shift modifier and then
;       select the "9" key.  The net result is the "(" key.  In addition,
;       shift-only keys are not supported as values for the p_Hotkey parameter
;       as a default hotkey.  If a shift-only key is used, no default key will
;       be selected.
;
;    *  To resolve a minor AutoHotkey inconsistency, the "Pause" key and the
;       "Break" keys are automatically converted to the "CtrlBreak" key if the
;       Ctrl modifier is selected.  The "CtrlBreak" key is automatically
;       converted to the "Pause" key if the Ctrl modifier is not selected.
;
;
; Hotkey Limits:
;
;   The p_Limit parameter allows the developer to restrict the types of keys
;   that are selected.  The following limit values are available:
;
;       (start code)
;       Limit   Description
;       -----   -----------
;       1       Prevent unmodified keys
;       2       Prevent Shift-only keys
;       4       Prevent Ctrl-only keys
;       8       Prevent Alt-only keys
;       16      Prevent Win-only keys
;       32      Prevent Shift-Ctrl keys
;       64      Prevent Shift-Alt keys
;       128     Prevent Shift-Win keys
;       256     Prevent Shift-Ctrl-Alt keys
;       512     Prevent Shift-Ctrl-Win keys
;       1024    Prevent Shift-Win-Alt keys
;       (end)
;
;   To use a limit, enter the sum of one or more of these limit values.  For
;   example, a limit value of 1 will prevent unmodified keys from being used.
;   A limit value of 31 (1 + 2 + 4 + 8 + 16) will require that at least two
;   modifier keys be used.
;
;
; Returns:
;
;   If the function ends after the user has selected a valid key and the
;   "Accept" button is pressed, the function returns the selected key in the
;   standard AutoHotkey hotkey format and ErrorLevel is set to 0.
;   Example: Hotkey=^a  ErrorLevel=0
;
;   If the HotkeyGUI window is canceled (Cancel button, Close button, or Escape
;   key), the function returns the original hotkey value (p_Hotkey) and
;   Errorlevel is set to 1.
;
;   If the function is unable to create a HotkeyGUI window for any reason,
;   ErrorLevel is set to the word FAIL.
;
;
; Calls To Other Functions:
;
;   PopupXY (optional)
;
;
; Hotkey Support:
;
;   AutoHotkey is a very robust program and can accept hotkey definitions in an
;   multitude of formats.  Unfortunately, this function is not that robust and
;   there are a couple of important limitations:
;
;    1.  The p_Limit parameter restricts the type of keys that can be supported.
;       For this reason, the following keys are not supported:
;
;        * *Modifier keys* (as hotkeys). Example: Alt, Shift, LWin, etc.
;        * *Joystick keys*. Example: Joy1, Joy2, etc.
;        * *Custom combinations*. Example: Numpad0 & Numpad1.
;
;    2.  Shift-only keys (Ex: "~","!","@","#",etc.) are not supported.  See the
;       "Processing and Usage Notes" section for more information.
;
;
; Programming Notes:
;
;   To keep the code as friendly as possible, static variables (in lieu of
;   global variables) are used whenever a GUI object needs a variable. Object
;   variables are defined so that a single "gui Submit" command can be used to
;   collect the GUI values instead of having to execute a "GUIControlGet"
;   command on every GUI control. For the few GUI objects that are
;   programmatically updated, the ClassNN (class name and instance number of the
;   object  Ex: Static4) is used.
;
;   Important: Any changes to the GUI (additions, deletions, etc.) may change
;   the ClassNN of objects that are updated.  Use Window Spy (or similar
;   program) to identify any changes.
;
;-------------------------------------------------------------------------------
Hotkey_GUI(p_Owner="",p_Hotkey="",p_Limit="",p_OptionalAttrib=False,p_Title="")
    {
    ;[====================]
    ;[  Static variables  ]
    ;[====================]
    Static s_GUI:=0
                ;-- This variable stores the currently active GUI.  If not zero
                ;   when entering the function, the GUI is currently showing.

          ,s_StartGUI:=53
                ;-- Default starting GUI window number for HotkeyGUI window.
                ;   Change if desired.

          ,s_PopupXY_Function:="PopupXY"
                ;-- Name of the PopupXY function.  Defined as a variable so that
                ;   function will use if the "PopupXY" function is included but
                ;   will not fail if it's not.

    ;[===========================]
    ;[  Window already showing?  ]
    ;[===========================]
    if s_GUI
        {
        Errorlevel:="FAIL"
        outputdebug,
           (ltrim join`s
            End Func: %A_ThisFunc% -
            A %A_ThisFunc% window already exists.  Errorlevel=FAIL
           )

        Return
        }

    ;[==============]
    ;[  Initialize  ]
    ;[==============]
    SplitPath A_ScriptName,,,,l_ScriptName
    l_ErrorLevel  :=0

    ;-------------
    ;-- Key lists
    ;-------------
    ;-- Standard keys
    l_StandardKeysList=
       (ltrim join|
        A|B|C|D|E|F|G|H|I|J|K|L|M|N|O|P|Q|R|S|T|U|V|W|X|Y|Z
        0|1|2|3|4|5|6|7|8|9|0
        ``|-|=|[|]|`\|;
        '|,|.|/
        Space
        Tab
        Enter
        Escape
        Backspace
        Delete
        ScrollLock
        CapsLock
        NumLock
        PrintScreen
        CtrlBreak
        Pause
        Break
        Insert
        Home
        End
        PgUp
        PgDn
        Up
        Down
        Left
        Right
       )

    ;-- Function keys
    l_FunctionKeysList=
       (ltrim join|
        F1|F2|F3|F4|F5|F6|F7|F8|F9|F10
        F11|F12|F13|F14|F15|F16|F17|F18|F19|F20
        F21|F22|F23|F24
       )

    ;-- Numpad
    l_NumpadKeysList=
       (ltrim join|
        NumLock
        NumpadDiv
        NumpadMult
        NumpadAdd
        NumpadSub
        NumpadEnter
        NumpadDel
        NumpadIns
        NumpadClear
        NumpadUp
        NumpadDown
        NumpadLeft
        NumpadRight
        NumpadHome
        NumpadEnd
        NumpadPgUp
        NumpadPgDn
        Numpad0
        Numpad1
        Numpad2
        Numpad3
        Numpad4
        Numpad5
        Numpad6
        Numpad7
        Numpad8
        Numpad9
        NumpadDot
       )

    ;-- Mouse
    l_MouseKeysList=
       (ltrim join|
        LButton
        RButton
        MButton
        WheelDown
        WheelUp
        XButton1
        XButton2
       )

    ;-- Multimedia
    l_MultimediaKeysList=
       (ltrim join|
        Browser_Back
        Browser_Forward
        Browser_Refresh
        Browser_Stop
        Browser_Search
        Browser_Favorites
        Browser_Home
        Volume_Mute
        Volume_Down
        Volume_Up
        Media_Next
        Media_Prev
        Media_Stop
        Media_Play_Pause
        Launch_Mail
        Launch_Media
        Launch_App1
        Launch_App2
       )

    ;-- Special
    l_SpecialKeysList:="Help|Sleep"

    ;[==================]
    ;[    Parameters    ]
    ;[  (Set defaults)  ]
    ;[==================]
    ;-- Owner
    p_Owner=%p_Owner%  ;-- AutoTrim
    if p_Owner is not Integer
        p_Owner:=0
     else
        if p_Owner not Between 1 and 99
            p_Owner:=0

    ;-- Owner window exists?
    if p_Owner
        {
        gui %p_Owner%:+LastFoundExist
        IfWinNotExist
            {
            outputdebug,
               (ltrim join`s
                Function: %A_ThisFunc% -
                Owner window does not exist.  p_Owner=%p_Owner%
               )

            p_Owner:=0
            }
        }

    ;-- Default hotkey
    l_Hotkey=%p_Hotkey%  ;-- AutoTrim

    ;-- Limit
    p_Limit=%p_Limit%  ;-- AutoTrim
    if p_Limit is not Integer
        p_Limit:=0
     else
        if p_Limit not between 0 and 2047
            p_Limit:=0

    ;-- Title
    p_Title=%p_Title%  ;-- AutoTrim
    if p_Title is Space
        p_Title:=l_ScriptName . " - Select Hotkey"
     else
        {
        ;-- Append to script name if p_title begins with "++"?
        if SubStr(p_Title,1,2)="++"
            {
            StringTrimLeft p_Title,p_Title,2
            p_Title:=l_ScriptName . A_Space . p_Title
            }
        }

    ;[==============================]
    ;[     Find available window    ]
    ;[  (Starting with s_StartGUI)  ]
    ;[==============================]
    s_GUI:=s_StartGUI
    Loop
        {
        ;-- Window available?
        gui %s_GUI%:+LastFoundExist
        IfWinNotExist
            Break

        ;-- Nothing available?
        if (s_GUI=99)
            {
            MsgBox
                ,262160
                    ;-- 262160=0 (OK button) + 16 (Error icon) + 262144 (AOT)
                ,%A_ThisFunc% Error,
                   (ltrim join`s
                    Unable to create a %A_ThisFunc% window.  GUI windows
                    %s_StartGUI% to 99 are already in use.  %A_Space%
                   )

            s_GUI:=0
            ErrorLevel:="FAIL"
            Return
            }

        ;-- Increment window
        s_GUI++
        }

    ;[=============]
    ;[  Build GUI  ]
    ;[=============]
    ;-- Assign ownership
    if p_Owner
        {
        gui %p_Owner%:+Disabled      ;-- Disable Owner window
        gui %s_GUI%:+Owner%p_Owner%  ;-- Set ownership
        }
     else
        gui %s_GUI%:+Owner           ;-- Gives ownership to the script window

    ;-- GUI options
    gui %s_GUI%:Margin,6,6
    gui %s_GUI%:-MinimizeBox +LabelHotkeyGUI_

    if not p_Owner
        gui %s_GUI%:+AlwaysOnTop

    ;---------------
    ;-- GUI objects
    ;---------------
    ;-- Modifiers
    Static HG_ModifierGB
    gui %s_GUI%:Add
       ,GroupBox
       , xm y10 w170 h10 vHG_ModifierGB
       ,Modifier

    Static HG_CtrlModifier
    gui %s_GUI%:Add
       ,CheckBox
       ,xp+10 yp+20 Section vHG_CtrlModifier gHotkeyGUI_UpdateHotkey
       ,Ctrl

    Static HG_ShiftModifier
    gui %s_GUI%:Add
       ,CheckBox
       ,xs vHG_ShiftModifier gHotkeyGUI_UpdateHotkey
       ,Shift

    Static HG_WinModifier
    gui %s_GUI%:Add
       ,CheckBox
       ,xs vHG_WinModifier gHotkeyGUI_UpdateHotkey
       ,Win

    Static HG_AltModifier
    gui %s_GUI%:Add
       ,CheckBox
       ,xs vHG_AltModifier gHotkeyGUI_UpdateHotkey
       ,Alt

    ;-- Optional Attributes
    Static HG_OptionalAttributesGB
    gui %s_GUI%:Add
       ,GroupBox
       ,xs+160 y10 w170 h10 vHG_OptionalAttributesGB
       ,Optional Attributes

    Static HG_NativeOption
    gui %s_GUI%:Add
       ,CheckBox
       ,xp+10 yp+20 Disabled Section vHG_NativeOption gHotkeyGUI_UpdateHotkey
       ,~ (Native)

    Static HG_WildcardOption
    gui %s_GUI%:Add
       ,CheckBox
       ,xs Disabled vHG_WildcardOption gHotkeyGUI_UpdateHotkey
       ,*  (Wildcard)

    Static HG_LeftPairOption
    gui %s_GUI%:Add                                                 ;-- Button9
       ,CheckBox
       ,xs Disabled vHG_LeftPairOption gHotkeyGUI_LeftPair
       ,< (Left pair only)

    Static HG_RightPairOption
    gui %s_GUI%:Add                                                 ;-- Button10
       ,CheckBox
       ,xs Disabled vHG_RightPairOption gHotkeyGUI_RightPair
       ,> (Right pair only)

    ;-- Enable "Optional Attributes"?
    if p_OptionalAttrib
        {
        GUIControl %s_GUI%:Enable,HG_NativeOption
        GUIControl %s_GUI%:Enable,HG_WildcardOption
        GUIControl %s_GUI%:Enable,HG_LeftPairOption
        GUIControl %s_GUI%:Enable,HG_RightPairOption
        }

    ;-- Resize the Modifier and Optional Attributes group boxes
    GUIControlGet $Group1Pos,%s_GUI%:Pos,HG_OptionalAttributesGB
    GUIControlGet $Group2Pos,%s_GUI%:Pos,HG_RightPairOption
    GUIControl
        ,%s_GUI%:Move
        ,HG_ModifierGB
        ,% "h" . ($Group2PosY-$Group1PosY)+$Group2PosH+10

    GUIControl
        ,%s_GUI%:Move
        ,HG_OptionalAttributesGB
        ,% "h" . ($Group2PosY-$Group1PosY)+$Group2PosH+10

    ;-- Keys
    YPos:=($Group2PosY-$Group1PosY)+$Group2PosH+20
    gui %s_GUI%:Add
       ,GroupBox
       ,xm y%YPos% w340 h180
       ,Keys

    Static HG_StandardKeysView
    gui %s_GUI%:Add
       ,Radio
       ,xp+10 yp+20 Checked Section vHG_StandardKeysView gHotkeyGUI_UpdateKeyList
       ,Standard

    Static HG_FunctionKeysView
    gui %s_GUI%:Add
       ,Radio
       ,xs vHG_FunctionKeysView gHotkeyGUI_UpdateKeyList
       ,Function keys

    Static HG_NumpadKeysView
    gui %s_GUI%:Add
       ,Radio
       ,xs vHG_NumpadKeysView gHotkeyGUI_UpdateKeyList
       ,Numpad

    Static HG_MouseKeysView
    gui %s_GUI%:Add
       ,Radio
       ,xs vHG_MouseKeysView gHotkeyGUI_UpdateKeyList
       ,Mouse

    Static HG_MultimediaKeysView
    gui %s_GUI%:Add
       ,Radio
       ,xs vHG_MultimediaKeysView gHotkeyGUI_UpdateKeyList
       ,Multimedia

    Static HG_SpecialKeysView
    gui %s_GUI%:Add
       ,Radio
       ,xs vHG_SpecialKeysView gHotkeyGUI_UpdateKeyList
       ,Special

    Static HG_Key
    gui %s_GUI%:Add
       ,ListBox                                                     ;-- ListBox1
       ,xs+140 ys w180 h150 vHG_Key gHotkeyGUI_UpdateHotkey

    ;-- Set initial values
    gosub HotkeyGUI_UpdateKeyList

    ;-- Hotkey display
    YPos+=190
    gui %s_GUI%:Add
       ,Text
       ,xm y%YPos% w70
       ,Hotkey:

    gui %s_GUI%:Add
       ,Edit                                                        ;-- Edit1
       ,x+0 w270 +ReadOnly

    gui %s_GUI%:Add
       ,Text
       ,xm y+5 w70 r2 hp
       ,Desc:

    gui %s_GUI%:Add
       ,Text                                                        ;-- Static3
       ,x+0 w270 hp +ReadOnly
       ,None

    ;-- Buttons
    Static HG_AcceptButton
    gui %s_GUI%:Add                                                 ;-- Button18
       ,Button
       ,xm y+5 Default Disabled vHG_AcceptButton gHotkeyGUI_AcceptButton
       ,%A_Space% &Accept %A_Space%
            ;-- Note: All characters are used to determine the button's W+H

    gui %s_GUI%:Add
       ,Button
       ,x+5 wp hp gHotkeyGUI_Close
       ,Cancel

     gui %s_GUI%:Add ; TD: add reset button
       ,Button
       ,x+5 wp hp gHotkeyGUI_Reset
       ,Reset

    ;[================]
    ;[  Set defaults  ]
    ;[================]
    if l_Hotkey is not Space
        {
        ;-- Modifiers and optional attributes
        Loop
            {
            l_FirstChar:=SubStr(l_Hotkey,1,1)
            if l_FirstChar in ^,+,#,!,~,*,<,>
                {
                if (l_FirstChar="^")
                    GUIControl %s_GUI%:,HG_CtrlModifier,1
                else if (l_FirstChar="+")
                    GUIControl %s_GUI%:,HG_ShiftModifier,1
                else if (l_FirstChar="#")
                    GUIControl %s_GUI%:,HG_WinModifier,1
                else if (l_FirstChar="!")
                    GUIControl %s_GUI%:,HG_AltModifier,1
                else  if (l_FirstChar="~")
                    GUIControl %s_GUI%:,HG_NativeOption,1
                else if (l_FirstChar="*")
                    GUIControl %s_GUI%:,HG_WildcardOption,1
                else if (l_FirstChar="<")
                    GUIControl %s_GUI%:,HG_LeftPairOption,1
                else if (l_FirstChar=">")
                    GUIControl %s_GUI%:,HG_RightPairOption,1
    
                ;-- On to the next
                StringTrimLeft l_Hotkey,l_Hotkey,1
                Continue
                }
    
            ;-- We're done here
            Break
            }
    
        ;-- Find key in key lists
        if l_Hotkey is not Space
            {
            ;-- Standard keys
            if Instr("|" . l_StandardKeysList . "|","|" . l_Hotkey . "|")
                GUIControl %s_GUI%:,HG_StandardKeysView,1
            ;-- Function keys
            else if Instr("|" . l_FunctionKeysList . "|","|" . l_Hotkey . "|")
                GUIControl %s_GUI%:,HG_FunctionKeysView,1
            ;-- Numpad keys
            else if Instr("|" . l_NumpadKeysList . "|","|" . l_Hotkey . "|")
                GUIControl %s_GUI%:,HG_NumpadKeysView,1
            ;-- Mouse keys
            else if Instr("|" . l_MouseKeysList . "|","|" . l_Hotkey . "|")
                GUIControl %s_GUI%:,HG_MouseKeysView,1
            ;-- Multimedia keys
            else if Instr("|" . l_MultimediaKeysList . "|","|" . l_Hotkey . "|")
                GUIControl %s_GUI%:,HG_MultimediaKeysView,1
            ;-- Special keys
            else if Instr("|" . l_SpecialKeysList . "|","|" . l_Hotkey . "|")
                GUIControl %s_GUI%:,HG_SpecialKeysView,1
    
            ;-- Update keylist and select it
            gosub HotkeyGUI_UpdateKeyList
            GUIControl %s_GUI%:ChooseString,HG_Key,%l_Hotkey%
            }

        ;-- Update Hotkey field and description
        gosub HotkeyGUI_UpdateHotkey
        }

    ;[=============]
    ;[  Set focus  ]
    ;[=============]
    GUIControl %s_GUI%:Focus,HG_AcceptButton
        ;-- Note: This only works when the Accept button is enabled

    ;[================]
    ;[  Collect hWnd  ]
    ;[================]
    gui %s_GUI%:+LastFound
    WinGet l_HotkeyGUI_hWnd,ID

    ;[===============]
    ;[  Show window  ]
    ;[===============]
     if p_Owner and IsFunc(s_PopupXY_Function)
        {
        gui %s_GUI%:Show,Hide,%p_Title%   ;-- Render but don't show
        %s_PopupXY_Function%(p_Owner,"ahk_id " . l_HotkeyGUI_hWnd,PosX,PosY)
        gui %s_GUI%:Show,x%PosX% y%PosY%  ;-- Show in the correct location
        }
     else
        gui %s_GUI%:Show,,%p_Title%

    ;[=====================]
    ;[  Loop until window  ]
    ;[      is closed      ]
    ;[=====================]
    WinWaitClose ahk_id %l_HotkeyGUI_hWnd%

    ;[====================]
    ;[  Return to sender  ]
    ;[====================]
    ErrorLevel:=l_ErrorLevel
    Return HG_HotKey  ;-- End of function



    ;*****************************
    ;*                           *
    ;*                           *
    ;*        Subroutines        *
    ;*        (HotkeyGUI)        *
    ;*                           *
    ;*                           *
    ;*****************************
    ;***********************
    ;*                     *
    ;*    Update Hotkey    *
    ;*                     *
    ;***********************
    HotkeyGUI_UpdateHotkey:

    ;-- Collect form values
    gui %s_GUI%:Submit,NoHide

    ;-- Enable/Disable Accept button
    if HG_Key
        GUIControl %s_GUI%:Enable,Button18
     else
        GUIControl %s_GUI%:Disable,Button18

    ;-- Substitute Pause|Break for CtrlBreak?
    if HG_Key in Pause,Break
        if HG_CtrlModifier
            HG_Key:="CtrlBreak"

    ;-- Substitute CtrlBreak for Pause (Break would work OK too)
    if (HG_Key="CtrlBreak")
        if not HG_CtrlModifier
            HG_Key:="Pause"

    ;[================]
    ;[  Build Hotkey  ]
    ;[================]
    ;-- Initialize
    HG_Hotkey:=""
    HG_HKDesc:=""

    ;-- Options
    if HG_NativeOption
        HG_Hotkey.="~"

    if HG_WildcardOption
        HG_Hotkey.="*"

    if HG_LeftPairOption
        HG_Hotkey.="<"

    if HG_RightPairOption
        HG_Hotkey.=">"

    ;-- Modifiers
    if HG_CtrlModifier
        {
        HG_Hotkey.="^"
        HG_HKDesc.="Ctrl + "
        }

    if HG_ShiftModifier
        {
        HG_Hotkey.="+"
        HG_HKDesc.="Shift + "
        }

    if HG_WinModifier
        {
        HG_Hotkey.="#"
        HG_HKDesc.="Win + "
        }

    if HG_AltModifier
        {
        HG_Hotkey.="!"
        HG_HKDesc.="Alt + "
        }

    HG_Hotkey.=HG_Key
    HG_HKDesc.=HG_Key

    ;-- Update Hotkey and HKDescr fields
    GUIControl %s_GUI%:,Edit1,%HG_Hotkey%
    GUIControl %s_GUI%:,Static3,%HG_HKDesc%
    return


    ;**********************
    ;*                    *
    ;*    Pair options    *
    ;*                    *
    ;**********************
    HotkeyGUI_LeftPair:

    ;-- Deselect HG_RightPairOption
    GUIControl %s_GUI%:,Button10,0
    gosub HotkeyGUI_UpdateHotkey
    return


    HotkeyGUI_RightPair:

    ;-- Deselect HG_LeftPairOption
    GUIControl %s_GUI%:,Button9,0
    gosub HotkeyGUI_UpdateHotkey
    return


    ;*************************
    ;*                       *
    ;*    Update Key List    *
    ;*                       *
    ;*************************
    HotkeyGUI_UpdateKeyList:

    ;-- Collect form values
    gui %s_GUI%:Submit,NoHide

    ;-- Standard
    if HG_StandardKeysView
        l_KeysList:=l_StandardKeysList
     else
        ;-- Function keys
        if HG_FunctionKeysView
            l_KeysList:=l_FunctionKeysList
         else
            ;-- Numpad
            if HG_NumpadKeysView
                l_KeysList:=l_NumpadKeysList
             else
                ;-- Mouse
                if HG_MouseKeysView
                    l_KeysList:=l_MouseKeysList
                 else
                    ;-- Multimedia
                    if HG_MultimediaKeysView
                        l_KeysList:=l_MultimediaKeysList
                     else
                        ;-- Special
                        if HG_SpecialKeysView
                            l_KeysList:=l_SpecialKeysList

    ;-- Update l_KeysList
    GUIControl %s_GUI%:-Redraw,ListBox1
    GUIControl %s_GUI%:,ListBox1,|%l_KeysList%
    GUIControl %s_GUI%:+Redraw,ListBox1

    ;--- Reset HG_Hotkey and HG_HKDesc
    HG_Key:=""
    gosub HotkeyGUI_UpdateHotkey
    return


    ;***********************
    ;*                     *
    ;*    Accept Button    *
    ;*                     *
    ;***********************
    HotkeyGUI_AcceptButton:

    ;-- (The following test is now redundant but it is retained as a fail-safe)
    ;-- Any key?
    if HG_Key is Space
        {
        gui %s_GUI%:+OwnDialogs
        MsgBox
            ,16         ;-- Error icon
            ,%p_Title%
            ,A key must be selected.  %A_Space%

        return
        }

    ;[===============]
    ;[  Limit tests  ]
    ;[===============]
    l_Limit:=p_Limit
    l_LimitFailure:=False

    ;-- Loop until failure or until all tests have been performed
    Loop
        {
        ;-- Are we done here?
        if (l_limit<=0)
            Break

        ;-----------------
        ;-- Shift+Win+Alt
        ;-----------------
        if (l_limit>=1024)
            {
            if (HG_ShiftModifier and HG_WinModifier and HG_AltModifier)
                {
                l_Message:="SHIFT+WIN+ALT keys are not allowed."
                l_LimitFailure:=True
                Break
                }

            l_limit-=1024
            Continue
            }

        ;------------------
        ;-- Shift+Ctrl+Win
        ;------------------
        if (l_limit>=512)
            {
            if (HG_ShiftModifier and HG_CtrlModifier and HG_WinModifier)
                {
                l_Message:="SHIFT+CTRL+WIN keys are not allowed."
                l_LimitFailure:=True
                Break
                }

            l_limit-=512
            Continue
            }

        ;------------------
        ;-- Shift+Ctrl+Alt
        ;------------------
        if (l_limit>=256)
            {
            if (HG_ShiftModifier and HG_CtrlModifier and HG_AltModifier)
                {
                l_Message:="SHIFT+CTRL+ALT keys are not allowed."
                l_LimitFailure:=True
                Break
                }

            l_limit-=256
            Continue
            }

        ;-------------
        ;-- Shift+Win
        ;-------------
        if (l_limit>=128)
            {
            if (HG_ShiftModifier and HG_WinModifier)
                {
                l_Message:="SHIFT+WIN keys are not allowed."
                l_LimitFailure:=True
                Break
                }

            l_limit-=128
            Continue
            }

        ;-------------
        ;-- Shift+Alt
        ;-------------
        if (l_limit>=64)
            {
            if (HG_ShiftModifier and HG_AltModifier)
                {
                l_Message:="SHIFT+ALT keys are not allowed."
                l_LimitFailure:=True
                Break
                }

            l_limit-=64
            Continue
            }

        ;--------------
        ;-- Shift+Ctrl
        ;--------------
        if (l_limit>=32)
            {
            if (HG_ShiftModifier and HG_CtrlModifier)
                {
                l_Message:="SHIFT+CTRL keys are not allowed."
                l_LimitFailure:=True
                Break
                }

            l_limit-=32
            Continue
            }

        ;------------
        ;-- Win only
        ;------------
        if (l_limit>=16)
            {
            if (HG_WinModifier
            and not (HG_CtrlModifier or HG_ShiftModifier or HG_AltModifier))
                {
                l_Message:="WIN-only keys are not allowed."
                l_LimitFailure:=True
                Break
                }

            l_limit-=16
            Continue
            }

        ;------------
        ;-- Alt only
        ;------------
        if (l_limit>=8)
            {
            if (HG_AltModifier
            and not (HG_CtrlModifier or HG_ShiftModifier or HG_WinModifier))
                {
                l_Message:="ALT-only keys are not allowed."
                l_LimitFailure:=True
                Break
                }

            l_limit-=8
            Continue
            }

        ;-------------
        ;-- Ctrl only
        ;-------------
        if (l_limit>=4)
            {
            if (HG_CtrlModifier
            and not (HG_ShiftModifier or HG_WinModifier or HG_AltModifier))
                {
                l_Message:="CTRL-only keys are not allowed."
                l_LimitFailure:=True
                Break
                }

            l_limit-=4
            Continue
            }

        ;--------------
        ;-- Shift only
        ;--------------
        if (l_limit>=2)
            {
            if (HG_ShiftModifier
            and not (HG_CtrlModifier or HG_WinModifier or HG_AltModifier))
                {
                l_Message:="SHIFT-only keys are not allowed."
                l_LimitFailure:=True
                Break
                }

            l_limit-=2
            Continue
            }

        ;--------------
        ;-- Unmodified
        ;--------------
        if (l_limit>=1)
            {
            if not (HG_CtrlModifier
                or  HG_ShiftModifier
                or  HG_WinModifier
                or  HG_AltModifier)
                {
                l_Message=
                   (ltrim join`s
                    At least one modifier must be used.  Other restrictions
                    may apply.
                   )

                l_LimitFailure:=True
                Break
                }

            l_limit-=1
            Continue
            }
        }

    ;[====================]
    ;[  Display message?  ]
    ;[====================]
    if l_LimitFailure
        {
        ;-- Display message
        gui %s_GUI%:+OwnDialogs
        MsgBox
            ,16         ;-- Error icon
            ,%p_Title%
            ,%l_Message%  %A_Space%

        ;-- Send 'em back
        return
        }

    ;[==================]
    ;[  Ok, We're done  ]
    ;[   Shut it done   ]
    ;[==================]
    gosub HotkeyGUI_Exit
    return


    ;***********************
    ;*                     *
    ;*    Close up shop    *
    ;*                     *
    ;***********************
    HotkeyGUI_Escape:
    HotkeyGUI_Close:
    HG_Hotkey:=p_Hotkey
    l_ErrorLevel:=1
    goto, HotkeyGUI_Exit ;TD

    HotkeyGUI_Reset: ; TD
    HG_Hotkey:=

    HotkeyGUI_Exit:

    ;-- Enable Owner window
    if p_Owner
        gui %p_Owner%:-Disabled

    ;-- Destroy the HotkeyGUI window so that the window can be reused
    gui %s_GUI%:Destroy
    s_GUI:=0
    return  ;-- End of subroutines
    }






/*
Hotkey_Parse()
© Avi Aryan
https://github.com/aviaryan/autohotkey-scripts/blob/master/Functions/HotkeyParser.ahk

5th Revision - 8/7/14
=========================================================================
Extract Autohotkey hotkeys from user-friendly shortcuts reliably and V.V
=========================================================================
==========================================
EXAMPLES - Pre-Runs
==========================================

msgbox % Hparse("Cntrol + ass + S", false)		;returns <blank>   	As 'ass' is out of scope and RemoveInvaild := false
msgbox % Hparse("Contrl + At + S")		;returns ^!s
msgbox % Hparse("^!s")			;returns	^!s		as the function-feed is already in Autohotkey format.
msgbox % Hparse("LeftContrl + X")		;returns Lcontrol & X
msgbox % Hparse("Contrl + Pageup + S")		;returns <blank>  As the hotkey is invalid
msgbox % HParse("PagUp + Ctrl", true)		;returns  ^PgUp  	as  ManageOrder is true (by default)
msgbox % HParse("PagUp + Ctrl", true, false)		;returns  <blank>  	as ManageOrder is false and hotkey is invalid	
msgbox % Hparse("Ctrl + Alt + Ctrl + K")		;returns  <blank> 	as two Ctrls are wrong
msgbox % HParse("Control + Alt")		;returns  ^Alt and NOT ^!
msgbox % HParse("Ctrl + F1 + Nmpd1")		;returns <blank>	As the hotkey is invalid
msgbox % HParse("Prbitscreen + f1")		;returns	PrintScreen & F1
msgbox % HParse("Prbitscreen + yyy")		;returns	PrintScreen		As RemoveInvalid is enabled by default.
msgbox % HParse("f1+ browser_srch")		;returns	F1 & Browser_Search
msgbox % HParse("Ctrl + joy1")			;returns	Ctrl & Joy1
msgbox % Hparse("pagup & paegdown")		;returns	PgUp & PgDn
MsgBox % HParse("Ctrl + printskreen", 1, 1, 1) 	; SEND Mode - on returns ^{printscreen}
msgbox % Hparse_rev("^!s")		;returns Ctrl+Alt+S
msgbox % Hparse_rev("Pgup & PgDn")		;returns Pageup & PgDn
*/

;###################################################################
;PARAMETERS - Hotkey_Parse() 		[See also Hotkey_ParseRev() below]
;-------------------------------
;Hotkey_Parse(Hotkey, RemoveInvalid, ManageOrder, sendMd)
;###################################################################

;• Hotkey - The user shortcut such as (Control + Alt + X) to be converted

;• RemoveInvalid(true) - Remove Invalid entries such as the 'ass' from (Control + ass + S) so that the return is ^s. When false the function will return <blank> when an
;  invalid entry is found.
  
;• ManageOrder(true) - Change (X + Control) to ^x and not x^ so that you are free from errors. If false, a <blank> value is returned when the hotkey is found un-ordered.

;+ SendMd(true) - returns ^{printscreen} instead of ^printscreen so that the hotkey properly works with the Send command

Hotkey_Parse(Hotkey, RemoveInvaild = true, ManageOrder = true, sendMd=false)
{

firstkey := Substr(Hotkey, 1, 1)
if firstkey in ^,!,+,#
	return, Hotkey

loop,parse,Hotkey,+-&,%a_space%
{
	if (Strlen(A_LoopField) != 1)
	{
		parsed := Hparse_LiteRegexM(A_LoopField)
		if sendMd && (StrLen(parsed)>1) && (Instr(parsed, "vk") != 1)
			parsed := "{" parsed "}"
		If !(RemoveInvaild)
		{
			IfEqual,parsed
			{
				Combo = 
				break
			}
			else
				Combo .= " & " . parsed
		}
		else
			IfNotEqual,parsed
				Combo .= " & " . parsed
	}
	else
		Combo .= " & " . A_LoopField
}

non_hotkey := 0
IfNotEqual, Combo		;Convert the hotkey to perfect format
{
	StringTrimLeft,Combo,Combo,3
	loop,parse,Combo,&,%A_Space%
	{
		if A_Loopfield not in ^,!,+,#
			non_hotkey+=1
	}
;END OF LOOP
	if (non_hotkey == 0)
	{
		StringRight,rightest,Combo,1
		StringTrimRight,Combo,Combo,1
		IfEqual,rightest,^
			rightest = Ctrl
		else IfEqual,rightest,!
			rightest = Alt
		ELSE IfEqual,rightest,+
			rightest = Shift
		else rightest = LWin
		Combo := Combo . Rightest
	}
;Remove last non
	IfLess,non_hotkey,2
	{
	IfNotInString,Combo,Joy
	{
		StringReplace,Combo,Combo,%A_Space%&%A_Space%,,All
		temp := Combo
		loop,parse,temp
		{
			if A_loopfield in ^,!,+,#
			{
			StringReplace,Combo,Combo,%A_loopfield%
			_hotkey .= A_loopfield
			}
		}
		Combo := _hotkey . Combo
	
		If !(ManageOrder)				;ManageOrder
			IfNotEqual,Combo,%temp%
				Combo = 
	
		temp := "^!+#"		;just reusing the variable . Checking for Duplicates Actually.
		IfNotEqual,Combo
		{
			loop,parse,temp
			{
				StringGetPos,pos,Combo,%A_loopfield%,L2
				IF (pos != -1){
					Combo = 
					break
				}
			}
		}
	;End of Joy
	}
	else	;Managing Joy
	{
		StringReplace,Combo,Combo,^,Ctrl
		StringReplace,Combo,Combo,!,Alt
		StringReplace,Combo,Combo,+,Shift
		StringReplace,Combo,Combo,#,LWin
		StringGetPos,pos,Combo,&,L2
		if (pos != -1)
			Combo = 
	}
}
else
{
	StringGetPos,pos,Combo,&,L2
	if (pos != -1)
		Combo = 
}
}

return, Combo
}

;###########################################################################################
;Hparse_rev(Keycombo)
;	Returns the user displayable format of Ahk Hotkey
;###########################################################################################

; TD changed 2023-10-04: remove spaces between +
Hotkey_ParseRev(Keycombo){

	if Instr(Keycombo, "&")
	{
		loop,parse,Keycombo,&,%A_space%%A_tab%
			toreturn .= A_LoopField "+"
		return Substr(toreturn, 1, -1)
	}
	Else
	{
		StringReplace, Keycombo, Keycombo,^,Ctrl&
		StringReplace, Keycombo, Keycombo,#,Win&
		StringReplace, Keycombo, Keycombo,+,Shift&
		StringReplace, Keycombo, Keycombo,!,Alt&
		loop,parse,Keycombo,&,%A_space%%A_tab%
			toreturn .= ( Strlen(A_LoopField)=1 ? Hparse_StringUpper(A_LoopField) : A_LoopField ) "+"
		return Substr(toreturn, 1, -1)
	}
}

Hparse_StringUpper(str){
	StringUpper, o, str
	return o
}

;------------------------------------------------------
;SYSTEM FUNCTIONS : NOT FOR USER'S USE
;------------------------------------------------------

Hparse_LiteRegexM(matchitem, primary=1)
{

regX := Hparse_ListGen("RegX", primary)
keys := Hparse_Listgen("Keys", primary)
matchit := matchitem

loop,parse,Regx,`r`n,
{
	curX := A_LoopField
	matchitem := matchit
	exitfrombreak := false

	loop,parse,A_LoopField,*
	{
		if (A_index == 1)
			if (SubStr(matchitem, 1, 1) != A_LoopField){
				exitfrombreak := true
				break
			}

		if (Hparse_comparewith(matchitem, A_loopfield))
			matchitem := Hparse_Vanish(matchitem, A_LoopField)
		else{
			exitfrombreak := true
			break
		}
	}

	if !(exitfrombreak){
		linenumber := A_Index
		break
	}
}

IfNotEqual, linenumber
{
	StringGetPos,pos1,keys,`n,% "L" . (linenumber - 1)
	StringGetPos,pos2,keys,`n,% "L" . (linenumber)
	return, Substr(keys, (pos1 + 2), (pos2 - pos1 - 1))
}
else
	return Hparse_LiteRegexM(matchit, 2)
}
; Extra Functions -----------------------------------------------------------------------------------------------------------------

Hparse_Vanish(matchitem, character){
	StringGetPos,pos,matchitem,%character%,L
	StringTrimLeft,matchitem,matchitem,(pos + 1)
	return, matchitem
}

Hparse_comparewith(first, second)
{
if first is Integer
	IfEqual,first,%second%
		return, true
	else
		return, false

IfInString,first,%second%
	return, true
else
	return, false
}

;######################   DANGER    ################################
;SIMPLY DONT EDIT BELOW THIS . MORE OFTEN THAN NOT, YOU WILL MESS IT.
;###################################################################
Hparse_ListGen(what,primary=1){
if (primary == 1)
{
IfEqual,what,Regx
Rvar = 
(
L*c*t
r*c*t
l*s*i
r*s*i
l*a*t
r*a*t
S*p*c
C*t*r
A*t
S*t
W*N
t*b
E*r
E*s*c
B*K
D*l
I*S
H*m
E*d
P*u
p*d
l*b*t
r*b*t
m*b*t
up
d*n
l*f
r*t
F*1
F*2
F*3
F*4
F*5
F*6
F*7
F*8
F*9
F*10
F*11
F*12
N*p*Do
N*p*D*v
N*p*M*t
N*p*d*Ad
N*p*S*b
N*p*E*r
s*l*k
c*l
n*l*k
p*s
c*t*b
pa*s
b*r*k
x*b*1
x*b*2
z*z*z*z*callmelazybuthtisisaworkaround
)
;====================================================
;# Original return values below (in respect with their above positions, dont EDIT)
IfEqual,what,Keys
Rvar = 
(
LControl
RControl 
LShift
RShift
LAlt
RAlt
space
^
!
+
#
Tab
Enter
Escape
Backspace
Delete
Insert
Home
End
PgUp
PgDn
LButton
RButton
MButton
Up
Down
Left
Right
F1
F2
F3
F4
F5
F6
F7
F8
F9
F10
F11
F12
NumpadDot
NumpadDiv
NumpadMult
NumpadAdd
NumpadSub
NumpadEnter
ScrollLock
CapsLock
NumLock
PrintScreen
CtrlBreak
Pause
Break
XButton1
XButton2
A_lazys_workaround
)
}
else
{
;here starts the second preference list.
IfEqual,what,Regx
Rvar=
(
N*p*0
N*p*1
N*p*2
N*p*3
N*p*4
N*p*5
N*p*6
N*p*7
N*p*8
N*p*9
F*13
F*14
F*15
F*16
F*17
F*18
F*19
F*20
F*21
F*22
F*23
F*24
N*p*I*s
N*p*E*d
N*p*D*N
N*p*P*D
N*p*L*f
N*p*C*r
N*p*R*t
N*p*H*m
N*p*Up
N*p*P*U
N*p*D*l
J*y*1
J*y*2
J*y*3
J*y*4
J*y*5
J*y*6
J*y*7
J*y*8
J*y*9
J*y*10
J*y*11
J*y*12
J*y*13
J*y*14
J*y*15
J*y*16
J*y*17
J*y*18
J*y*19
J*y*20
J*y*21
J*y*22
J*y*23
J*y*24
J*y*25
J*y*26
J*y*27
J*y*28
J*y*29
J*y*30
J*y*31
J*y*32
B*_B*k
B*_F*r
B*_R*e*h
B*_S*p
B*_S*c
B*_F*t
B*_H*m
V*_M*e
V*_D*n
V*_U
M*_N*x
M*_P
M*_S*p
M*_P*_P
L*_M*l
L*_M*a
L*_A*1
L*_A*2

)
IfEqual,what,keys
Rvar=
(
Numpad0
Numpad1
Numpad2
Numpad3
Numpad4
Numpad5
Numpad6
Numpad7
Numpad8
Numpad9
F13
F14
F15
F16
F17
F18
F19
F20
F21
F22
F23
F24
NumpadIns
NumpadEnd
NumpadDown
NumpadPgDn
NumpadLeft
NumpadClear
NumpadRight
NumpadHome
NumpadUp
NumpadPgUp
NumpadDel
Joy1
Joy2
Joy3
Joy4
Joy5
Joy6
Joy7
Joy8
Joy9
Joy10
Joy11
Joy12
Joy13
Joy14
Joy15
Joy16
Joy17
Joy18
Joy19
Joy20
Joy21
Joy22
Joy23
Joy24
Joy25
Joy26
Joy27
Joy28
Joy29
Joy30
Joy31
Joy32
Browser_Back
Browser_Forward
Browser_Refresh
Browser_Stop
Browser_Search
Browser_Favorites
Browser_Home
Volume_Mute
Volume_Down
Volume_Up
Media_Next
Media_Prev
Media_Stop
Media_Play_Pause
Launch_Mail
Launch_Media
Launch_App1
Launch_App2

)
}
;<<<<<<<<<<<<<<<<END>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
return, Rvar
}