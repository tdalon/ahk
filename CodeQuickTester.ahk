; CodeQuickTester v2.8
; Copyright GeekDude 2018
; https://github.com/G33kDude/CodeQuickTester

#SingleInstance, Off
#NoEnv
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

global B_Params := []
Loop, %0%
	B_Params.Push(%A_Index%)

Menu, Tray, Icon, %A_AhkPath%, 2
FileEncoding, UTF-8

Settings :=
( LTrim Join Comments
{
	; File path for the starting contents
	"DefaultPath": "C:\Windows\ShellNew\Template.ahk",

	; When True, this setting may conflict with other instances of CQT
	"GlobalRun": False,

	; Script options
	"AhkPath": A_AhkPath,
	"Params": "",

	; Editor (colors are 0xBBGGRR)
	"FGColor": 0xEDEDCD,
	"BGColor": 0x3F3F3F,
	"TabSize": 4,
	"Font": {
		"Typeface": "Consolas",
		"Size": 11,
		"Bold": False
	},
	"Gutter": {
		; Width in pixels. Make this larger when using
		; larger fonts. Set to 0 to disable the gutter.
		"Width": 40,

		"FGColor": 0x9FAFAF,
		"BGColor": 0x262626
	},

	; Highlighter (colors are 0xRRGGBB)
	"UseHighlighter": True,
	"Highlighter": "HighlightAHK",
	"HighlightDelay": 200, ; Delay until the user is finished typing
	"Colors": {
		"Comments":     0x7F9F7F,
		"Functions":    0x7CC8CF,
		"Keywords":     0xE4EDED,
		"Multiline":    0x7F9F7F,
		"Numbers":      0xF79B57,
		"Punctuation":  0x97C0EB,
		"Strings":      0xCC9893,
		"A_Builtins":   0xF79B57,
		"Commands":     0xCDBFA3,
		"Directives":   0x7CC8CF,
		"Flow":         0xE4EDED,
		"KeyNames":     0xCB8DD9
	},

	; Auto-Indenter
	"Indent": "`t",

	; Pastebin
	"DefaultName": A_UserName,
	"DefaultDesc": "Pasted with CodeQuickTester",

	; AutoComplete
	"UseAutoComplete": True,
	"ACListRebuildDelay": 500 ; Delay until the user is finished typing
}
)

; Overlay any external settings onto the above defaults
if FileExist("Settings.ini")
{
	ExtSettings := Ini_Load(FileOpen("Settings.ini", "r").Read())
	for k, v in ExtSettings
		if IsObject(v)
			v.base := Settings[k]
	ExtSettings.base := Settings
	Settings := ExtSettings
}

Tester := new CodeQuickTester(Settings)
Tester.RegisterCloseCallback(Func("TesterClose"))
return

#If Tester.Exec.Status == 0 ; Running

~*Escape::Tester.Exec.Terminate()

#If (Tester.Settings.GlobalRun && Tester.Exec.Status == 0) ; Running

F5::
!r::
; Reloads
Tester.RunButton()
Tester.RunButton()
return

#If Tester.Settings.GlobalRun

F5::
!r::
Tester.RunButton()
return

#If

TesterClose(Tester)
{
	ExitApp
}

/*
	class RichCode({"TabSize": 4     ; Width of a tab in characters
	, "Indent": "`t"             ; What text to insert on indent
	, "FGColor": 0xRRGGBB        ; Foreground (text) color
	, "BGColor": 0xRRGGBB        ; Background color
	, "Font"                     ; Font to use
	: {"Typeface": "Courier New" ; Name of the typeface
	, "Size": 12             ; Font size in points
	, "Bold": False}         ; Bolded (True/False)
	
	
	; Whether to use the highlighter, or leave it as plain text
	, "UseHighlighter": True
	
	; Delay after typing before the highlighter is run
	, "HighlightDelay": 200
	
	; The highlighter function (FuncObj or name)
	; to generate the highlighted RTF. It will be passed
	; two parameters, the first being this settings array
	; and the second being the code to be highlighted
	, "Highlighter": Func("HighlightAHK")
	
	; The colors to be used by the highlighter function.
	; This is currently used only by the highlighter, not at all by the
	; RichCode class. As such, the RGB ordering is by convention only.
	; You can add as many colors to this array as you want.
	, "Colors"
	: [0xRRGGBB
	, 0xRRGGBB
	, 0xRRGGBB,
	, 0xRRGGBB]})
*/

class RichCode
{
	static Msftedit := DllCall("LoadLibrary", "Str", "Msftedit.dll")
	static IID_ITextDocument := "{8CC497C0-A1DF-11CE-8098-00AA0047BE5D}"
	static MenuItems := ["Cut", "Copy", "Paste", "Delete", "", "Select All", ""
	, "UPPERCASE", "lowercase", "TitleCase"]
	
	_Frozen := False
	
	; --- Static Methods ---
	
	BGRFromRGB(RGB)
	{
		return RGB>>16&0xFF | RGB&0xFF00 | RGB<<16&0xFF0000
	}
	
	; --- Properties ---
	
	Value[]
	{
		get {
			GuiControlGet, Code,, % this.hWnd
			return Code
		}
		
		set {
			this.Highlight(Value)
			return Value
		}
	}
	
	; TODO: reserve and reuse memory
	Selection[i:=0]
	{
		get {
			VarSetCapacity(CHARRANGE, 8, 0)
			this.SendMsg(0x434, 0, &CHARRANGE) ; EM_EXGETSEL
			Out := [NumGet(CHARRANGE, 0, "Int"), NumGet(CHARRANGE, 4, "Int")]
			return i ? Out[i] : Out
		}
		
		set {
			if i
				Temp := this.Selection, Temp[i] := Value, Value := Temp
			VarSetCapacity(CHARRANGE, 8, 0)
			NumPut(Value[1], &CHARRANGE, 0, "Int") ; cpMin
			NumPut(Value[2], &CHARRANGE, 4, "Int") ; cpMax
			this.SendMsg(0x437, 0, &CHARRANGE) ; EM_EXSETSEL
			return Value
		}
	}
	
	SelectedText[]
	{
		get {
			Selection := this.Selection, Length := Selection[2] - Selection[1]
			VarSetCapacity(Buffer, (Length + 1) * 2) ; +1 for null terminator
			if (this.SendMsg(0x43E, 0, &Buffer) > Length) ; EM_GETSELTEXT
				throw Exception("Text larger than selection! Buffer overflow!")
			Text := StrGet(&Buffer, Selection[2]-Selection[1], "UTF-16")
			return StrReplace(Text, "`r", "`n")
		}
		
		set {
			this.SendMsg(0xC2, 1, &Value) ; EM_REPLACESEL
			this.Selection[1] -= StrLen(Value)
			return Value
		}
	}
	
	EventMask[]
	{
		get {
			return this._EventMask
		}
		
		set {
			this._EventMask := Value
			this.SendMsg(0x445, 0, Value) ; EM_SETEVENTMASK
			return Value
		}
	}
	
	UndoSuspended[]
	{
		get {
			return this._UndoSuspended
		}
		
		set {
			try ; ITextDocument is not implemented in WINE
			{
				if Value
					this.ITextDocument.Undo(-9999995) ; tomSuspend
				else
					this.ITextDocument.Undo(-9999994) ; tomResume
			}
			return this._UndoSuspended := !!Value
		}
	}
	
	Frozen[]
	{
		get {
			return this._Frozen
		}
		
		set {
			if (Value && !this._Frozen)
			{
				try ; ITextDocument is not implemented in WINE
					this.ITextDocument.Freeze()
				catch
					GuiControl, -Redraw, % this.hWnd
			}
			else if (!Value && this._Frozen)
			{
				try ; ITextDocument is not implemented in WINE
					this.ITextDocument.Unfreeze()
				catch
					GuiControl, +Redraw, % this.hWnd
			}
			return this._Frozen := !!Value
		}
	}
	
	Modified[]
	{
		get {
			return this.SendMsg(0xB8, 0, 0) ; EM_GETMODIFY
		}
		
		set {
			this.SendMsg(0xB9, Value, 0) ; EM_SETMODIFY
			return Value
		}
	}
	
	; --- Construction, Destruction, Meta-Functions ---
	
	__New(Settings, Options:="")
	{
		static Test
		this.Settings := Settings
		FGColor := this.BGRFromRGB(Settings.FGColor)
		BGColor := this.BGRFromRGB(Settings.BGColor)
		
		Gui, Add, Custom, ClassRichEdit50W hWndhWnd +0x5031b1c4 +E0x20000 %Options%
		this.hWnd := hWnd
		
		; Enable WordWrap in RichEdit control ("WordWrap" : true)
		if this.Settings.WordWrap
			SendMessage, 0x0448, 0, 0, , % "ahk_id " . This.HWND
		
		; Register for WM_COMMAND and WM_NOTIFY events
		; NOTE: this prevents garbage collection of
		; the class until the control is destroyed
		this.EventMask := 1 ; ENM_CHANGE
		CtrlEvent := this.CtrlEvent.Bind(this)
		GuiControl, +g, %hWnd%, %CtrlEvent%
		
		; Set background color
		this.SendMsg(0x443, 0, BGColor) ; EM_SETBKGNDCOLOR
		
		; Set character format
		VarSetCapacity(CHARFORMAT2, 116, 0)
		NumPut(116,                    CHARFORMAT2, 0,  "UInt")       ; cbSize      = sizeof(CHARFORMAT2)
		NumPut(0xE0000000,             CHARFORMAT2, 4,  "UInt")       ; dwMask      = CFM_COLOR|CFM_FACE|CFM_SIZE
		NumPut(FGColor,                CHARFORMAT2, 20, "UInt")       ; crTextColor = 0xBBGGRR
		NumPut(Settings.Font.Size*20,  CHARFORMAT2, 12, "UInt")       ; yHeight     = twips
		StrPut(Settings.Font.Typeface, &CHARFORMAT2+26, 32, "UTF-16") ; szFaceName  = TCHAR
		this.SendMsg(0x444, 0, &CHARFORMAT2) ; EM_SETCHARFORMAT
		
		; Set tab size to 4 for non-highlighted code
		VarSetCapacity(TabStops, 4, 0), NumPut(Settings.TabSize*4, TabStops, "UInt")
		this.SendMsg(0x0CB, 1, &TabStops) ; EM_SETTABSTOPS
		
		; Change text limit from 32,767 to max
		this.SendMsg(0x435, 0, -1) ; EM_EXLIMITTEXT
		
		; Bind for keyboard events
		; Use a pointer to prevent reference loop
		this.OnMessageBound := this.OnMessage.Bind(&this)
		OnMessage(0x100, this.OnMessageBound) ; WM_KEYDOWN
		OnMessage(0x205, this.OnMessageBound) ; WM_RBUTTONUP
		
		; Bind the highlighter
		this.HighlightBound := this.Highlight.Bind(&this)
		
		; Create the right click menu
		this.MenuName := this.__Class . &this
		RCMBound := this.RightClickMenu.Bind(&this)
		for Index, Entry in this.MenuItems
			Menu, % this.MenuName, Add, %Entry%, %RCMBound%
		
		; Get the ITextDocument object
		VarSetCapacity(pIRichEditOle, A_PtrSize, 0)
		this.SendMsg(0x43C, 0, &pIRichEditOle) ; EM_GETOLEINTERFACE
		this.pIRichEditOle := NumGet(pIRichEditOle, 0, "UPtr")
		this.IRichEditOle := ComObject(9, this.pIRichEditOle, 1), ObjAddRef(this.pIRichEditOle)
		this.pITextDocument := ComObjQuery(this.IRichEditOle, this.IID_ITextDocument)
		this.ITextDocument := ComObject(9, this.pITextDocument, 1), ObjAddRef(this.pITextDocument)
	}
	
	RightClickMenu(ItemName, ItemPos, MenuName)
	{
		if !IsObject(this)
			this := Object(this)
		
		if (ItemName == "Cut")
			Clipboard := this.SelectedText, this.SelectedText := ""
		else if (ItemName == "Copy")
			Clipboard := this.SelectedText
		else if (ItemName == "Paste")
			this.SelectedText := Clipboard
		else if (ItemName == "Delete")
			this.SelectedText := ""
		else if (ItemName == "Select All")
			this.Selection := [0, -1]
		else if (ItemName == "UPPERCASE")
			this.SelectedText := Format("{:U}", this.SelectedText)
		else if (ItemName == "lowercase")
			this.SelectedText := Format("{:L}", this.SelectedText)
		else if (ItemName == "TitleCase")
			this.SelectedText := Format("{:T}", this.SelectedText)
	}
	
	__Delete()
	{
		; Release the ITextDocument object
		this.ITextDocument := "", ObjRelease(this.pITextDocument)
		this.IRichEditOle := "", ObjRelease(this.pIRichEditOle)
		
		; Release the OnMessage handlers
		OnMessage(0x100, this.OnMessageBound, 0) ; WM_KEYDOWN
		OnMessage(0x205, this.OnMessageBound, 0) ; WM_RBUTTONUP
		
		; Destroy the right click menu
		Menu, % this.MenuName, Delete
		
		HighlightBound := this.HighlightBound
		SetTimer, %HighlightBound%, Delete
	}
	
	; --- Event Handlers ---
	
	OnMessage(wParam, lParam, Msg, hWnd)
	{
		if !IsObject(this)
			this := Object(this)
		if (hWnd != this.hWnd)
			return
		
		if (Msg == 0x100) ; WM_KEYDOWN
		{
			if (wParam == GetKeyVK("Tab"))
			{
				; Indentation
				Selection := this.Selection
				if GetKeyState("Shift")
					this.IndentSelection(True) ; Reverse
				else if (Selection[2] - Selection[1]) ; Something is selected
					this.IndentSelection()
				else
				{
					; TODO: Trim to size needed to reach next TabSize
					this.SelectedText := this.Settings.Indent
					this.Selection[1] := this.Selection[2] ; Place cursor after
				}
				return False
			}
			else if (wParam == GetKeyVK("Escape")) ; Normally closes the window
				return False
			else if (wParam == GetKeyVK("v") && GetKeyState("Ctrl"))
			{
				this.SelectedText := Clipboard ; Strips formatting
				this.Selection[1] := this.Selection[2] ; Place cursor after
				return False
			}
		}
		else if (Msg == 0x205) ; WM_RBUTTONUP
		{
			Menu, % this.MenuName, Show
			return False
		}
	}
	
	CtrlEvent(CtrlHwnd, GuiEvent, EventInfo, _ErrorLevel:="")
	{
		if (GuiEvent == "Normal" && EventInfo == 0x300) ; EN_CHANGE
		{
			; Delay until the user is finished changing the document
			HighlightBound := this.HighlightBound
			SetTimer, %HighlightBound%, % -Abs(this.Settings.HighlightDelay)
		}
	}
	
	; --- Methods ---
	
	; First parameter is taken as a replacement value
	; Variadic form is used to detect when a parameter is given,
	; regardless of content
	Highlight(NewVal*)
	{
		if !IsObject(this)
			this := Object(this)
		if !(this.Settings.UseHighlighter && this.Settings.Highlighter)
		{
			if NewVal.Length()
				GuiControl,, % this.hWnd, % NewVal[1]
			return
		}
		
		; Freeze the control while it is being modified, stop change event
		; generation, suspend the undo buffer, buffer any input events
		PrevFrozen := this.Frozen, this.Frozen := True
		PrevEventMask := this.EventMask, this.EventMask := 0 ; ENM_NONE
		PrevUndoSuspended := this.UndoSuspended, this.UndoSuspended := True
		PrevCritical := A_IsCritical
		Critical, 1000
		
		; Run the highlighter
		Highlighter := this.Settings.Highlighter
		RTF := %Highlighter%(this.Settings, NewVal.Length() ? NewVal[1] : this.Value)
		
		; "TRichEdit suspend/resume undo function"
		; https://stackoverflow.com/a/21206620
		
		; Save the rich text to a UTF-8 buffer
		VarSetCapacity(Buf, StrPut(RTF, "UTF-8"), 0)
		StrPut(RTF, &Buf, "UTF-8")
		
		; Set up the necessary structs
		VarSetCapacity(ZOOM,      8, 0) ; Zoom Level
		VarSetCapacity(POINT,     8, 0) ; Scroll Pos
		VarSetCapacity(CHARRANGE, 8, 0) ; Selection
		VarSetCapacity(SETTEXTEX, 8, 0) ; SetText Settings
		NumPut(1, SETTEXTEX, 0, "UInt") ; flags = ST_KEEPUNDO
		
		; Save the scroll and cursor positions, update the text,
		; then restore the scroll and cursor positions
		MODIFY := this.SendMsg(0xB8, 0, 0)    ; EM_GETMODIFY
		this.SendMsg(0x4E0, &ZOOM, &ZOOM+4)   ; EM_GETZOOM
		this.SendMsg(0x4DD, 0, &POINT)        ; EM_GETSCROLLPOS
		this.SendMsg(0x434, 0, &CHARRANGE)    ; EM_EXGETSEL
		this.SendMsg(0x461, &SETTEXTEX, &Buf) ; EM_SETTEXTEX
		this.SendMsg(0x437, 0, &CHARRANGE)    ; EM_EXSETSEL
		this.SendMsg(0x4DE, 0, &POINT)        ; EM_SETSCROLLPOS
		this.SendMsg(0x4E1, NumGet(ZOOM, "UInt")
		, NumGet(ZOOM, 4, "UInt"))        ; EM_SETZOOM
		this.SendMsg(0xB9, MODIFY, 0)         ; EM_SETMODIFY
		
		; Restore previous settings
		Critical, %PrevCritical%
		this.UndoSuspended := PrevUndoSuspended
		this.EventMask := PrevEventMask
		this.Frozen := PrevFrozen
	}
	
	IndentSelection(Reverse:=False, Indent:="")
	{
		; Freeze the control while it is being modified, stop change event
		; generation, buffer any input events
		PrevFrozen := this.Frozen, this.Frozen := True
		PrevEventMask := this.EventMask, this.EventMask := 0 ; ENM_NONE
		PrevCritical := A_IsCritical
		Critical, 1000
		
		if (Indent == "")
			Indent := this.Settings.Indent
		IndentLen := StrLen(Indent)
		
		; Select back to the start of the first line
		Min := this.Selection[1]
		Top := this.SendMsg(0x436, 0, Min) ; EM_EXLINEFROMCHAR
		TopLineIndex := this.SendMsg(0xBB, Top, 0) ; EM_LINEINDEX
		this.Selection[1] := TopLineIndex
		
		; TODO: Insert newlines using SetSel/ReplaceSel to avoid having to call
		; the highlighter again
		Text := this.SelectedText
		if Reverse
		{
			; Remove indentation appropriately
			Loop, Parse, Text, `n, `r
			{
				if (InStr(A_LoopField, Indent) == 1)
				{
					Out .= "`n" SubStr(A_LoopField, 1+IndentLen)
					if (A_Index == 1)
						Min -= IndentLen
				}
				else
					Out .= "`n" A_LoopField
			}
			this.SelectedText := SubStr(Out, 2)
			
			; Move the selection start back, but never onto the previous line
			this.Selection[1] := Min < TopLineIndex ? TopLineIndex : Min
		}
		else
		{
			; Add indentation appropriately
			Trailing := (SubStr(Text, 0) == "`n")
			Temp := Trailing ? SubStr(Text, 1, -1) : Text
			Loop, Parse, Temp, `n, `r
				Out .= "`n" Indent . A_LoopField
			this.SelectedText := SubStr(Out, 2) . (Trailing ? "`n" : "")
			
			; Move the selection start forward
			this.Selection[1] := Min + IndentLen
		}
		
		this.Highlight()
		
		; Restore previous settings
		Critical, %PrevCritical%
		this.EventMask := PrevEventMask
		
		; When content changes cause the horizontal scrollbar to disappear,
		; unfreezing causes the scrollbar to jump. To solve this, jump back
		; after unfreezing. This will cause a flicker when that edge case
		; occurs, but it's better than the alternative.
		VarSetCapacity(POINT, 8, 0)
		this.SendMsg(0x4DD, 0, &POINT) ; EM_GETSCROLLPOS
		this.Frozen := PrevFrozen
		this.SendMsg(0x4DE, 0, &POINT) ; EM_SETSCROLLPOS
	}
	
	; --- Helper/Convenience Methods ---
	
	SendMsg(Msg, wParam, lParam)
	{
		SendMessage, Msg, wParam, lParam,, % "ahk_id" this.hWnd
		return ErrorLevel
	}
}
GenHighlighterCache(Settings)
{
	if Settings.HasKey("Cache")
		return
	Cache := Settings.Cache := {}
	
	
	; --- Process Colors ---
	Cache.Colors := Settings.Colors.Clone()
	
	; Inherit from the Settings array's base
	BaseSettings := Settings
	while (BaseSettings := BaseSettings.Base)
		for Name, Color in BaseSettings.Colors
			if !Cache.Colors.HasKey(Name)
				Cache.Colors[Name] := Color
	
	; Include the color of plain text
	if !Cache.Colors.HasKey("Plain")
		Cache.Colors.Plain := Settings.FGColor
	
	; Create a Name->Index map of the colors
	Cache.ColorMap := {}
	for Name, Color in Cache.Colors
		Cache.ColorMap[Name] := A_Index
	
	
	; --- Generate the RTF headers ---
	RTF := "{\urtf"
	
	; Color Table
	RTF .= "{\colortbl;"
	for Name, Color in Cache.Colors
	{
		RTF .= "\red"   Color>>16 & 0xFF
		RTF .= "\green" Color>>8  & 0xFF
		RTF .= "\blue"  Color     & 0xFF ";"
	}
	RTF .= "}"
	
	; Font Table
	if Settings.Font
	{
		FontTable .= "{\fonttbl{\f0\fmodern\fcharset0 "
		FontTable .= Settings.Font.Typeface
		FontTable .= ";}}"
		RTF .= "\fs" Settings.Font.Size * 2 ; Font size (half-points)
		if Settings.Font.Bold
			RTF .= "\b"
	}
	
	; Tab size (twips)
	RTF .= "\deftab" GetCharWidthTwips(Settings.Font) * Settings.TabSize
	
	Cache.RTFHeader := RTF
}

GetCharWidthTwips(Font)
{
	static Cache := {}
	
	if Cache.HasKey(Font.Typeface "_" Font.Size "_" Font.Bold)
		return Cache[Font.Typeface "_" font.Size "_" Font.Bold]
	
	; Calculate parameters of CreateFont
	Height := -Round(Font.Size*A_ScreenDPI/72)
	Weight := 400+300*(!!Font.Bold)
	Face := Font.Typeface
	
	; Get the width of "x"
	hDC := DllCall("GetDC", "UPtr", 0)
	hFont := DllCall("CreateFont"
	, "Int", Height ; _In_ int     nHeight,
	, "Int", 0      ; _In_ int     nWidth,
	, "Int", 0      ; _In_ int     nEscapement,
	, "Int", 0      ; _In_ int     nOrientation,
	, "Int", Weight ; _In_ int     fnWeight,
	, "UInt", 0     ; _In_ DWORD   fdwItalic,
	, "UInt", 0     ; _In_ DWORD   fdwUnderline,
	, "UInt", 0     ; _In_ DWORD   fdwStrikeOut,
	, "UInt", 0     ; _In_ DWORD   fdwCharSet, (ANSI_CHARSET)
	, "UInt", 0     ; _In_ DWORD   fdwOutputPrecision, (OUT_DEFAULT_PRECIS)
	, "UInt", 0     ; _In_ DWORD   fdwClipPrecision, (CLIP_DEFAULT_PRECIS)
	, "UInt", 0     ; _In_ DWORD   fdwQuality, (DEFAULT_QUALITY)
	, "UInt", 0     ; _In_ DWORD   fdwPitchAndFamily, (FF_DONTCARE|DEFAULT_PITCH)
	, "Str", Face   ; _In_ LPCTSTR lpszFace
	, "UPtr")
	hObj := DllCall("SelectObject", "UPtr", hDC, "UPtr", hFont, "UPtr")
	VarSetCapacity(SIZE, 8, 0)
	DllCall("GetTextExtentPoint32", "UPtr", hDC, "Str", "x", "Int", 1, "UPtr", &SIZE)
	DllCall("SelectObject", "UPtr", hDC, "UPtr", hObj, "UPtr")
	DllCall("DeleteObject", "UPtr", hFont)
	DllCall("ReleaseDC", "UPtr", 0, "UPtr", hDC)
	
	; Convert to twpis
	Twips := Round(NumGet(SIZE, 0, "UInt")*1440/A_ScreenDPI)
	Cache[Font.Typeface "_" Font.Size "_" Font.Bold] := Twips
	return Twips
}

EscapeRTF(Code)
{
	for each, Char in ["\", "{", "}", "`n"]
		Code := StrReplace(Code, Char, "\" Char)
	return StrReplace(StrReplace(Code, "`t", "\tab "), "`r")
}

HighlightAHK(Settings, ByRef Code)
{
	static Flow := "break|byref|catch|class|continue|else|exit|exitapp|finally|for|global|gosub|goto|if|ifequal|ifexist|ifgreater|ifgreaterorequal|ifinstring|ifless|iflessorequal|ifmsgbox|ifnotequal|ifnotexist|ifnotinstring|ifwinactive|ifwinexist|ifwinnotactive|ifwinnotexist|local|loop|onexit|pause|return|settimer|sleep|static|suspend|throw|try|until|var|while"
	, Commands := "autotrim|blockinput|clipwait|control|controlclick|controlfocus|controlget|controlgetfocus|controlgetpos|controlgettext|controlmove|controlsend|controlsendraw|controlsettext|coordmode|critical|detecthiddentext|detecthiddenwindows|drive|driveget|drivespacefree|edit|envadd|envdiv|envget|envmult|envset|envsub|envupdate|fileappend|filecopy|filecopydir|filecreatedir|filecreateshortcut|filedelete|fileencoding|filegetattrib|filegetshortcut|filegetsize|filegettime|filegetversion|fileinstall|filemove|filemovedir|fileread|filereadline|filerecycle|filerecycleempty|fileremovedir|fileselectfile|fileselectfolder|filesetattrib|filesettime|formattime|getkeystate|groupactivate|groupadd|groupclose|groupdeactivate|gui|guicontrol|guicontrolget|hotkey|imagesearch|inidelete|iniread|iniwrite|input|inputbox|keyhistory|keywait|listhotkeys|listlines|listvars|menu|mouseclick|mouseclickdrag|mousegetpos|mousemove|msgbox|outputdebug|pixelgetcolor|pixelsearch|postmessage|process|progress|random|regdelete|regread|regwrite|reload|run|runas|runwait|send|sendevent|sendinput|sendlevel|sendmessage|sendmode|sendplay|sendraw|setbatchlines|setcapslockstate|setcontroldelay|setdefaultmousespeed|setenv|setformat|setkeydelay|setmousedelay|setnumlockstate|setregview|setscrolllockstate|setstorecapslockmode|settitlematchmode|setwindelay|setworkingdir|shutdown|sort|soundbeep|soundget|soundgetwavevolume|soundplay|soundset|soundsetwavevolume|splashimage|splashtextoff|splashtexton|splitpath|statusbargettext|statusbarwait|stringcasesense|stringgetpos|stringleft|stringlen|stringlower|stringmid|stringreplace|stringright|stringsplit|stringtrimleft|stringtrimright|stringupper|sysget|thread|tooltip|transform|traytip|urldownloadtofile|winactivate|winactivatebottom|winclose|winget|wingetactivestats|wingetactivetitle|wingetclass|wingetpos|wingettext|wingettitle|winhide|winkill|winmaximize|winmenuselectitem|winminimize|winminimizeall|winminimizeallundo|winmove|winrestore|winset|winsettitle|winshow|winwait|winwaitactive|winwaitclose|winwaitnotactive"
	, Functions := "abs|acos|array|asc|asin|atan|ceil|chr|comobjactive|comobjarray|comobjconnect|comobjcreate|comobject|comobjenwrap|comobjerror|comobjflags|comobjget|comobjmissing|comobjparameter|comobjquery|comobjtype|comobjunwrap|comobjvalue|cos|dllcall|exception|exp|fileexist|fileopen|floor|func|getkeyname|getkeysc|getkeystate|getkeyvk|il_add|il_create|il_destroy|instr|isbyref|isfunc|islabel|isobject|isoptional|ln|log|ltrim|lv_add|lv_delete|lv_deletecol|lv_getcount|lv_getnext|lv_gettext|lv_insert|lv_insertcol|lv_modify|lv_modifycol|lv_setimagelist|mod|numget|numput|objaddref|objclone|object|objgetaddress|objgetcapacity|objhaskey|objinsert|objinsertat|objlength|objmaxindex|objminindex|objnewenum|objpop|objpush|objrawset|objrelease|objremove|objremoveat|objsetcapacity|onmessage|ord|regexmatch|regexreplace|registercallback|round|rtrim|sb_seticon|sb_setparts|sb_settext|sin|sqrt|strget|strlen|strput|strsplit|substr|tan|trim|tv_add|tv_delete|tv_get|tv_getchild|tv_getcount|tv_getnext|tv_getparent|tv_getprev|tv_getselection|tv_gettext|tv_modify|tv_setimagelist|varsetcapacity|winactive|winexist|_addref|_clone|_getaddress|_getcapacity|_haskey|_insert|_maxindex|_minindex|_newenum|_release|_remove|_setcapacity"
	, Keynames := "alt|altdown|altup|appskey|backspace|blind|browser_back|browser_favorites|browser_forward|browser_home|browser_refresh|browser_search|browser_stop|bs|capslock|click|control|ctrl|ctrlbreak|ctrldown|ctrlup|del|delete|down|end|enter|esc|escape|f1|f10|f11|f12|f13|f14|f15|f16|f17|f18|f19|f2|f20|f21|f22|f23|f24|f3|f4|f5|f6|f7|f8|f9|home|ins|insert|joy1|joy10|joy11|joy12|joy13|joy14|joy15|joy16|joy17|joy18|joy19|joy2|joy20|joy21|joy22|joy23|joy24|joy25|joy26|joy27|joy28|joy29|joy3|joy30|joy31|joy32|joy4|joy5|joy6|joy7|joy8|joy9|joyaxes|joybuttons|joyinfo|joyname|joypov|joyr|joyu|joyv|joyx|joyy|joyz|lalt|launch_app1|launch_app2|launch_mail|launch_media|lbutton|lcontrol|lctrl|left|lshift|lwin|lwindown|lwinup|mbutton|media_next|media_play_pause|media_prev|media_stop|numlock|numpad0|numpad1|numpad2|numpad3|numpad4|numpad5|numpad6|numpad7|numpad8|numpad9|numpadadd|numpadclear|numpaddel|numpaddiv|numpaddot|numpaddown|numpadend|numpadenter|numpadhome|numpadins|numpadleft|numpadmult|numpadpgdn|numpadpgup|numpadright|numpadsub|numpadup|pause|pgdn|pgup|printscreen|ralt|raw|rbutton|rcontrol|rctrl|right|rshift|rwin|rwindown|rwinup|scrolllock|shift|shiftdown|shiftup|space|tab|up|volume_down|volume_mute|volume_up|wheeldown|wheelleft|wheelright|wheelup|xbutton1|xbutton2"
	, Builtins := "base|clipboard|clipboardall|comspec|errorlevel|false|programfiles|true"
	, Keywords := "abort|abovenormal|activex|add|ahk_class|ahk_exe|ahk_group|ahk_id|ahk_pid|all|alnum|alpha|altsubmit|alttab|alttabandmenu|alttabmenu|alttabmenudismiss|alwaysontop|and|autosize|background|backgroundtrans|base|belownormal|between|bitand|bitnot|bitor|bitshiftleft|bitshiftright|bitxor|bold|border|bottom|button|buttons|cancel|capacity|caption|center|check|check3|checkbox|checked|checkedgray|choose|choosestring|click|clone|close|color|combobox|contains|controllist|controllisthwnd|count|custom|date|datetime|days|ddl|default|delete|deleteall|delimiter|deref|destroy|digit|disable|disabled|dpiscale|dropdownlist|edit|eject|enable|enabled|error|exit|expand|exstyle|extends|filesystem|first|flash|float|floatfast|focus|font|force|fromcodepage|getaddress|getcapacity|grid|group|groupbox|guiclose|guicontextmenu|guidropfiles|guiescape|guisize|haskey|hdr|hidden|hide|high|hkcc|hkcr|hkcu|hkey_classes_root|hkey_current_config|hkey_current_user|hkey_local_machine|hkey_users|hklm|hku|hotkey|hours|hscroll|hwnd|icon|iconsmall|id|idlast|ignore|imagelist|in|insert|integer|integerfast|interrupt|is|italic|join|label|lastfound|lastfoundexist|left|limit|lines|link|list|listbox|listview|localsameasglobal|lock|logoff|low|lower|lowercase|ltrim|mainwindow|margin|maximize|maximizebox|maxindex|menu|minimize|minimizebox|minmax|minutes|monitorcount|monitorname|monitorprimary|monitorworkarea|monthcal|mouse|mousemove|mousemoveoff|move|multi|na|new|no|noactivate|nodefault|nohide|noicon|nomainwindow|norm|normal|nosort|nosorthdr|nostandard|not|notab|notimers|number|off|ok|on|or|owndialogs|owner|parse|password|pic|picture|pid|pixel|pos|pow|priority|processname|processpath|progress|radio|range|rawread|rawwrite|read|readchar|readdouble|readfloat|readint|readint64|readline|readnum|readonly|readshort|readuchar|readuint|readushort|realtime|redraw|regex|region|reg_binary|reg_dword|reg_dword_big_endian|reg_expand_sz|reg_full_resource_descriptor|reg_link|reg_multi_sz|reg_qword|reg_resource_list|reg_resource_requirements_list|reg_sz|relative|reload|remove|rename|report|resize|restore|retry|rgb|right|rtrim|screen|seconds|section|seek|send|sendandmouse|serial|setcapacity|setlabel|shiftalttab|show|shutdown|single|slider|sortdesc|standard|status|statusbar|statuscd|strike|style|submit|sysmenu|tab|tab2|tabstop|tell|text|theme|this|tile|time|tip|tocodepage|togglecheck|toggleenable|toolwindow|top|topmost|transcolor|transparent|tray|treeview|type|uncheck|underline|unicode|unlock|updown|upper|uppercase|useenv|useerrorlevel|useunsetglobal|useunsetlocal|vis|visfirst|visible|vscroll|waitclose|wantctrla|wantf2|wantreturn|wanttab|wrap|write|writechar|writedouble|writefloat|writeint|writeint64|writeline|writenum|writeshort|writeuchar|writeuint|writeushort|xdigit|xm|xp|xs|yes|ym|yp|ys|__call|__delete|__get|__handle|__new|__set"
	, Needle := "
	( LTrim Join Comments
		ODims)
		((?:^|\s);[^\n]+)          ; Comments
		|(^\s*\/\*.+?\n\s*\*\/)    ; Multiline comments
		|((?:^|\s)#[^ \t\r\n,]+)   ; Directives
		|([+*!~&\/\\<>^|=?:
			,().```%{}\[\]\-]+)    ; Punctuation
		|\b(0x[0-9a-fA-F]+|[0-9]+) ; Numbers
		|(""[^""\r\n]*"")          ; Strings
		|\b(A_\w*|" Builtins ")\b  ; A_Builtins
		|\b(" Flow ")\b            ; Flow
		|\b(" Commands ")\b        ; Commands
		|\b(" Functions ")\b       ; Functions (builtin)
		|\b(" Keynames ")\b        ; Keynames
		|\b(" Keywords ")\b        ; Other keywords
		|(([a-zA-Z_$]+)(?=\())     ; Functions
	)"
	
	GenHighlighterCache(Settings)
	Map := Settings.Cache.ColorMap
	
	Pos := 1
	while (FoundPos := RegExMatch(Code, Needle, Match, Pos))
	{
		RTF .= "\cf" Map.Plain " "
		RTF .= EscapeRTF(SubStr(Code, Pos, FoundPos-Pos))
		
		; Flat block of if statements for performance
		if (Match.Value(1) != "")
			RTF .= "\cf" Map.Comments
		else if (Match.Value(2) != "")
			RTF .= "\cf" Map.Multiline
		else if (Match.Value(3) != "")
			RTF .= "\cf" Map.Directives
		else if (Match.Value(4) != "")
			RTF .= "\cf" Map.Punctuation
		else if (Match.Value(5) != "")
			RTF .= "\cf" Map.Numbers
		else if (Match.Value(6) != "")
			RTF .= "\cf" Map.Strings
		else if (Match.Value(7) != "")
			RTF .= "\cf" Map.A_Builtins
		else if (Match.Value(8) != "")
			RTF .= "\cf" Map.Flow
		else if (Match.Value(9) != "")
			RTF .= "\cf" Map.Commands
		else if (Match.Value(10) != "")
			RTF .= "\cf" Map.Functions
		else if (Match.Value(11) != "")
			RTF .= "\cf" Map.Keynames
		else if (Match.Value(12) != "")
			RTF .= "\cf" Map.Keywords
		else if (Match.Value(13) != "")
			RTF .= "\cf" Map.Functions
		else
			RTF .= "\cf" Map.Plain
		
		RTF .= " " EscapeRTF(Match.Value())
		Pos := FoundPos + Match.Len()
	}
	
	return Settings.Cache.RTFHeader . RTF
	. "\cf" Map.Plain " " EscapeRTF(SubStr(Code, Pos)) "\`n}"
}
class CodeQuickTester
{
	static Msftedit := DllCall("LoadLibrary", "Str", "Msftedit.dll")
	EditorString := """" A_AhkPath """ """ A_ScriptFullPath """ ""%1"""
	OrigEditorString := "notepad.exe %1"
	Title := "CodeQuickTester"
	
	__New(Settings)
	{
		this.Settings := Settings
		
		this.Shell := ComObjCreate("WScript.Shell")
		
		this.Bound := []
		this.Bound.RunButton := this.RunButton.Bind(this)
		this.Bound.GuiSize := this.GuiSize.Bind(this)
		this.Bound.OnMessage := this.OnMessage.Bind(this)
		this.Bound.UpdateStatusBar := this.UpdateStatusBar.Bind(this)
		this.Bound.UpdateAutoComplete := this.UpdateAutoComplete.Bind(this)
		this.Bound.CheckIfRunning := this.CheckIfRunning.Bind(this)
		this.Bound.Highlight := this.Highlight.Bind(this)
		this.Bound.SyncGutter := this.SyncGutter.Bind(this)
		
		Buttons := new this.MenuButtons(this)
		this.Bound.Indent := Buttons.Indent.Bind(Buttons)
		this.Bound.Unindent := Buttons.Unindent.Bind(Buttons)
		Menus :=
		( LTrim Join Comments
		[
			["&File", [
				["&Run`tF5", this.Bound.RunButton],
				[],
				["&New`tCtrl+N", Buttons.New.Bind(Buttons)],
				["&Open`tCtrl+O", Buttons.Open.Bind(Buttons)],
				["Open &Working Dir`tCtrl+Shift+O", Buttons.OpenFolder.Bind(Buttons)],
				["&Save`tCtrl+S", Buttons.Save.Bind(Buttons, False)],
				["&Save as`tCtrl+Shift+S", Buttons.Save.Bind(Buttons, True)],
				["Rename", Buttons.Rename.Bind(Buttons)],
				[],
				["&Publish", Buttons.Publish.Bind(Buttons)],
				["&Fetch", Buttons.Fetch.Bind(Buttons)],
				[],
				["E&xit`tCtrl+W", this.GuiClose.Bind(this)]
			]], ["&Edit", [
				["Find`tCtrl+F", Buttons.Find.Bind(Buttons)],
				[],
				["Comment Lines`tCtrl+K", Buttons.Comment.Bind(Buttons)],
				["Uncomment Lines`tCtrl+Shift+K", Buttons.Uncomment.Bind(Buttons)],
				[],
				["Indent Lines", this.Bound.Indent],
				["Unindent Lines", this.Bound.Unindent],
				[],
				["Include &Relative", Buttons.IncludeRel.Bind(Buttons)],
				["Include &Absolute", Buttons.IncludeAbs.Bind(Buttons)],
				[],
				["Script &Options", Buttons.ScriptOpts.Bind(Buttons)]
			]], ["&Tools", [
				["&Pastebin`tCtrl+P", Buttons.Paste.Bind(Buttons)],
				["Re&indent`tCtrl+I", Buttons.AutoIndent.Bind(Buttons)],
				[],
				["&AlwaysOnTop`tAlt+A", Buttons.ToggleOnTop.Bind(Buttons)],
				["Global Run Hotkeys", Buttons.GlobalRun.Bind(Buttons)],
				[],
				["Install Service Handler", Buttons.ServiceHandler.Bind(Buttons)],
				["Set as Default Editor", Buttons.DefaultEditor.Bind(Buttons)],
				[],
				["&Highlighter", Buttons.Highlighter.Bind(Buttons)],
				["AutoComplete", Buttons.AutoComplete.Bind(Buttons)]
			]], ["&Help", [
				["Open &Help File`tCtrl+H", Buttons.Help.Bind(Buttons)],
				["&About", Buttons.About.Bind(Buttons)]
			]]
		]
		)
		
		Gui, New, +Resize +hWndhMainWindow -AlwaysOnTop
		this.AlwaysOnTop := False
		this.hMainWindow := hMainWindow
		this.Menus := CreateMenus(Menus)
		Gui, Menu, % this.Menus[1]
		
		; If set as default, check the highlighter option
		if this.Settings.UseHighlighter
			Menu, % this.Menus[4], Check, &Highlighter
		
		; If set as default, check the global run hotkeys option
		if this.Settings.GlobalRun
			Menu, % this.Menus[4], Check, Global Run Hotkeys
		
		; If set as default, check the AutoComplete option
		if this.Settings.UseAutoComplete
			Menu, % this.Menus[4], Check, AutoComplete
		
		; If service handler is installed, check the menu option
		if ServiceHandler.Installed()
			Menu, % this.Menus[4], Check, Install Service Handler
		
		RegRead, Editor, HKCR, AutoHotkeyScript\Shell\Edit\Command
		if (Editor == this.EditorString)
			Menu, % this.Menus[4], Check, Set as Default Editor
		
		; Register for events
		WinEvents.Register(this.hMainWindow, this)
		for each, Msg in [0x111, 0x100, 0x101, 0x201, 0x202, 0x204] ; WM_COMMAND, WM_KEYDOWN, WM_KEYUP, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN
			OnMessage(Msg, this.Bound.OnMessage)
		
		; Add code editor and gutter for line numbers
		this.RichCode := new RichCode(this.Settings, "-E0x20000")
		RichEdit_AddMargins(this.RichCode.hWnd, 3, 3)
		if Settings.Gutter.Width
			this.AddGutter()
		
		if B_Params.HasKey(1)
			FilePath := RegExReplace(B_Params[1], "^ahk:") ; Remove leading service handler
		else
			FilePath := Settings.DefaultPath
		
		if (FilePath ~= "^https?://")
			this.RichCode.Value := UrlDownloadToVar(FilePath)
		else if (FilePath = "Clipboard")
			this.RichCode.Value := Clipboard
		else if InStr(FileExist(FilePath), "A")
		{
			this.RichCode.Value := FileOpen(FilePath, "r").Read()
			this.RichCode.Modified := False
			
			if (FilePath == Settings.DefaultPath)
			{
				; Place cursor after the default template text
				this.RichCode.Selection := [-1, -1]
			}
			else
			{
				; Keep track of the file currently being edited
				this.FilePath := GetFullPathName(FilePath)
				
				; Follow the directory of the most recently opened file
				SetWorkingDir, %FilePath%\..
			}
		}
		else
			this.RichCode.Value := ""
		
		if (this.FilePath == "")
			Menu, % this.Menus[2], Disable, Rename
		
		; Add run button
		Gui, Add, Button, hWndhRunButton, &Run
		this.hRunButton := hRunButton
		BoundFunc := this.Bound.RunButton
		GuiControl, +g, %hRunButton%, %BoundFunc%
		
		; Add status bar
		Gui, Add, StatusBar, hWndhStatusBar
		this.UpdateStatusBar()
		ControlGetPos,,,, StatusBarHeight,, ahk_id %hStatusBar%
		this.StatusBarHeight := StatusBarHeight
		
		; Initialize the AutoComplete
		this.AC := new this.AutoComplete(this, this.settings.UseAutoComplete)
		
		this.UpdateTitle()
		Gui, Show, w640 h480
	}
	
	AddGutter()
	{
		s := this.Settings, f := s.Font, g := s.Gutter
		
		; Add the RichEdit control for the gutter
		Gui, Add, Custom, ClassRichEdit50W hWndhGutter +0x5031b1c6 -HScroll -VScroll
		this.hGutter := hGutter
		
		; Set the background and font settings
		FGColor := RichCode.BGRFromRGB(g.FGColor)
		BGColor := RichCode.BGRFromRGB(g.BGColor)
		VarSetCapacity(CF2, 116, 0)
		NumPut(116,        &CF2+ 0, "UInt") ; cbSize      = sizeof(CF2)
		NumPut(0xE<<28,    &CF2+ 4, "UInt") ; dwMask      = CFM_COLOR|CFM_FACE|CFM_SIZE
		NumPut(f.Size*20,  &CF2+12, "UInt") ; yHeight     = twips
		NumPut(FGColor,    &CF2+20, "UInt") ; crTextColor = 0xBBGGRR
		StrPut(f.Typeface, &CF2+26, 32, "UTF-16") ; szFaceName = TCHAR
		SendMessage(0x444, 0, &CF2,    hGutter) ; EM_SETCHARFORMAT
		SendMessage(0x443, 0, BGColor, hGutter) ; EM_SETBKGNDCOLOR
		
		RichEdit_AddMargins(hGutter, 3, 3, -3, 0)
	}
	
	RunButton()
	{
		if (this.Exec.Status == 0) ; Running
			this.Exec.Terminate() ; CheckIfRunning updates the GUI
		else ; Not running or doesn't exist
		{
			this.Exec := ExecScript(this.RichCode.Value
			, this.Settings.Params
			, this.Settings.AhkPath)
			
			GuiControl,, % this.hRunButton, &Kill
			
			SetTimer(this.Bound.CheckIfRunning, 100)
		}
	}
	
	CheckIfRunning()
	{
		if (this.Exec.Status == 1)
		{
			SetTimer(this.Bound.CheckIfRunning, "Delete")
			GuiControl,, % this.hRunButton, &Run
		}
	}
	
	LoadCode(Code, FilePath:="")
	{
		; Do nothing if nothing is changing
		if (this.FilePath == FilePath && this.RichCode.Value == Code)
			return
		
		; Confirm the user really wants to load new code
		Gui, +OwnDialogs
		MsgBox, 308, % this.Title " - Confirm Overwrite"
		, Are you sure you want to overwrite your code?
		IfMsgBox, No
			return
		
		; If we're changing the open file mark as modified
		; If we're loading a new file mark as unmodified
		this.RichCode.Modified := this.FilePath == FilePath
		this.FilePath := FilePath
		if (this.FilePath == "")
			Menu, % this.Menus[2], Disable, Rename
		else
			Menu, % this.Menus[2], Enable, Rename
		
		; Update the GUI
		this.RichCode.Value := Code
		this.UpdateStatusBar()
	}
	
	OnMessage(wParam, lParam, Msg, hWnd)
	{
		if (hWnd == this.hMainWindow && Msg == 0x111 ; WM_COMMAND
			&& lParam == this.RichCode.hWnd)         ; for RichEdit
		{
			Command := wParam >> 16
			
			if (Command == 0x400) ; An event that fires on scroll
			{
				this.SyncGutter()
				
				; If the user is scrolling too fast it can cause some messages
				; to be dropped. Set a timer to make sure that when the user stops
				; scrolling that the line numbers will be in sync.
				SetTimer(this.Bound.SyncGutter, -50)
			}
			else if (Command == 0x200) ; EN_KILLFOCUS
				if this.Settings.UseAutoComplete
					this.AC.Fragment := ""
		}
		else if (hWnd == this.RichCode.hWnd)
		{
			; Call UpdateStatusBar after the edit handles the keystroke
			SetTimer(this.Bound.UpdateStatusBar, -0)
			
			if this.Settings.UseAutoComplete
			{
				SetTimer(this.Bound.UpdateAutoComplete
				, -Abs(this.Settings.ACListRebuildDelay))
				
				if (Msg == 0x100) ; WM_KEYDOWN
					return this.AC.WM_KEYDOWN(wParam, lParam)
				else if (Msg == 0x201) ; WM_LBUTTONDOWN
					this.AC.Fragment := ""
			}
		}
		else if (hWnd == this.hGutter
			&& {0x100:1,0x101:1,0x201:1,0x202:1,0x204:1}[Msg]) ; WM_KEYDOWN, WM_KEYUP, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN
		{
			; Disallow interaction with the gutter
			return True
		}
	}
	
	SyncGutter()
	{
		static BUFF, _ := VarSetCapacity(BUFF, 16, 0)
		
		if !this.Settings.Gutter.Width
			return
		
		SendMessage(0x4E0, &BUFF, &BUFF+4, this.RichCode.hwnd) ; EM_GETZOOM
		SendMessage(0x4DD, 0, &BUFF+8, this.RichCode.hwnd)     ; EM_GETSCROLLPOS
		
		; Don't update the gutter unnecessarily
		State := NumGet(BUFF, 0, "UInt") . NumGet(BUFF, 4, "UInt")
		. NumGet(BUFF, 8, "UInt") . NumGet(BUFF, 12, "UInt")
		if (State == this.GutterState)
			return
		
		NumPut(-1, BUFF, 8, "UInt") ; Don't sync horizontal position
		Zoom := [NumGet(BUFF, "UInt"), NumGet(BUFF, 4, "UInt")]
		PostMessage(0x4E1, Zoom[1], Zoom[2], this.hGutter)     ; EM_SETZOOM
		PostMessage(0x4DE, 0, &BUFF+8, this.hGutter)           ; EM_SETSCROLLPOS
		this.ZoomLevel := Zoom[1] / Zoom[2]
		if (this.ZoomLevel != this.LastZoomLevel)
			SetTimer(this.Bound.GuiSize, -0), this.LastZoomLevel := this.ZoomLevel
		
		this.GutterState := State
	}
	
	GetKeywordFromCaret()
	{
		; https://autohotkey.com/boards/viewtopic.php?p=180369#p180369
		static Buffer
		IsUnicode := !!A_IsUnicode
		
		rc := this.RichCode
		sel := rc.Selection
		
		; Get the currently selected line
		LineNum := rc.SendMsg(0x436, 0, sel[1]) ; EM_EXLINEFROMCHAR
		
		; Size a buffer according to the line's length
		Length := rc.SendMsg(0xC1, sel[1], 0) ; EM_LINELENGTH
		VarSetCapacity(Buffer, Length << !!A_IsUnicode, 0)
		NumPut(Length, Buffer, "UShort")
		
		; Get the text from the line
		rc.SendMsg(0xC4, LineNum, &Buffer) ; EM_GETLINE
		lineText := StrGet(&Buffer, Length)
		
		; Parse the line to find the word
		LineIndex := rc.SendMsg(0xBB, LineNum, 0) ; EM_LINEINDEX
		RegExMatch(SubStr(lineText, 1, sel[1]-LineIndex), "[#\w]+$", Start)
		RegExMatch(SubStr(lineText, sel[1]-LineIndex+1), "^[#\w]+", End)
		
		return Start . End
	}
	
	UpdateStatusBar()
	{
		; Delete the timer if it was called by one
		SetTimer(this.Bound.UpdateStatusBar, "Delete")
		
		; Get the document length and cursor position
		VarSetCapacity(GTL, 8, 0), NumPut(1200, GTL, 4, "UInt")
		Len := this.RichCode.SendMsg(0x45F, &GTL, 0) ; EM_GETTEXTLENGTHEX (Handles newlines better than GuiControlGet on RE)
		ControlGet, Row, CurrentLine,,, % "ahk_id" this.RichCode.hWnd
		ControlGet, Col, CurrentCol,,, % "ahk_id" this.RichCode.hWnd
		
		; Get Selected Text Length
		; If the user has selected 1 char further than the end of the document,
		; which is allowed in a RichEdit control, subtract 1 from the length
		Sel := this.RichCode.Selection
		Sel := Sel[2] - Sel[1] - (Sel[2] > Len)
		
		; Get the syntax tip, if any
		if (SyntaxTip := HelpFile.GetSyntax(this.GetKeywordFromCaret()))
			this.SyntaxTip := SyntaxTip
		
		; Update the Status Bar text
		Gui, % this.hMainWindow ":Default"
		SB_SetText("Len " Len ", Line " Row ", Col " Col
		. (Sel > 0 ? ", Sel " Sel : "") "     " this.SyntaxTip)
		
		; Update the title Bar
		this.UpdateTitle()
		
		; Update the gutter to match the document
		if this.Settings.Gutter.Width
		{
			ControlGet, Lines, LineCount,,, % "ahk_id" this.RichCode.hWnd
			if (Lines != this.LineCount)
			{
				Loop, %Lines%
					Text .= A_Index "`n"
				GuiControl,, % this.hGutter, %Text%
				this.SyncGutter()
				this.LineCount := Lines
			}
		}
	}
	
	UpdateTitle()
	{
		Title := this.Title
		
		; Show the current file name
		if this.FilePath
		{
			SplitPath, % this.FilePath, FileName
			Title .= " - " FileName
		}
		
		; Show the curernt modification status
		if this.RichCode.Modified
			Title .= "*"
		
		; Return if the title doesn't need to be updated
		if (Title == this.VisibleTitle)
			return
		this.VisibleTitle := Title
		
		HiddenWindows := A_DetectHiddenWindows
		DetectHiddenWindows, On
		WinSetTitle, % "ahk_id" this.hMainWindow,, %Title%
		DetectHiddenWindows, %HiddenWindows%
	}
	
	UpdateAutoComplete()
	{
		; Delete the timer if it was called by one
		SetTimer(this.Bound.UpdateAutoComplete, "Delete")
		
		this.AC.BuildWordList()
	}
	
	RegisterCloseCallback(CloseCallback)
	{
		this.CloseCallback := CloseCallback
	}
	
	GuiSize()
	{
		static RECT, _ := VarSetCapacity(RECT, 16, 0)
		if A_Gui
			gw := A_GuiWidth, gh := A_GuiHeight
		else
		{
			DllCall("GetClientRect", "UPtr", this.hMainWindow, "Ptr", &RECT, "UInt")
			gw := NumGet(RECT, 8, "Int"), gh := NumGet(RECT, 12, "Int")
		}
		gtw := 3 + Round(this.Settings.Gutter.Width) * (this.ZoomLevel ? this.ZoomLevel : 1), sbh := this.StatusBarHeight
		GuiControl, Move, % this.RichCode.hWnd, % "x" 0+gtw "y" 0         "w" gw-gtw "h" gh-28-sbh
		if this.Settings.Gutter.Width
			GuiControl, Move, % this.hGutter  , % "x" 0     "y" 0         "w" gtw    "h" gh-28-sbh
		GuiControl, Move, % this.hRunButton   , % "x" 0     "y" gh-28-sbh "w" gw     "h" 28
	}
	
	GuiDropFiles(hWnd, Files)
	{
		; TODO: support multiple file drop
		this.LoadCode(FileOpen(Files[1], "r").Read(), Files[1])
	}
	
	GuiClose()
	{
		if this.RichCode.Modified
		{
			Gui, +OwnDialogs
			MsgBox, 308, % this.Title " - Confirm Exit", There are unsaved changes. Are you sure you want to exit?
			IfMsgBox, No
				return true
		}
		
		if (this.Exec.Status == 0) ; Running
		{
			SetTimer(this.Bound.CheckIfRunning, "Delete")
			this.Exec.Terminate()
		}
		
		; Free up the AC class
		this.AC := ""
		
		; Release wm_message hooks
		for each, Msg in [0x100, 0x201, 0x202, 0x204] ; WM_KEYDOWN, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONDOWN
			OnMessage(Msg, this.Bound.OnMessage, 0)
		
		; Delete timers
		SetTimer(this.Bound.SyncGutter, "Delete")
		SetTimer(this.Bound.GuiSize, "Delete")
		
		; Break all the BoundFunc circular references
		this.Delete("Bound")
		
		; Release WinEvents handler
		WinEvents.Unregister(this.hMainWindow)
		
		; Release GUI window and control glabels
		Gui, Destroy
		
		; Release menu bar (Has to be done after Gui, Destroy)
		for each, MenuName in this.Menus
			Menu, %MenuName%, DeleteAll
		
		this.CloseCallback()
	}
	
	class Paste
	{
		static Targets := {"IRC": "#ahk", "Discord": "discord"}
		
		__New(Parent)
		{
			this.Parent := Parent
			
			ParentWnd := this.Parent.hMainWindow
			Gui, New, +Owner%ParentWnd% +ToolWindow +hWndhWnd
			this.hWnd := hWnd
			Gui, Margin, 5, 5
			
			Gui, Add, Text, xm ym w30 h22 +0x200, Desc: ; 0x200 for vcenter
			Gui, Add, Edit, x+5 yp w125 h22 hWndhPasteDesc, % this.Parent.Settings.DefaultDesc
			this.hPasteDesc := hPasteDesc
			
			Gui, Add, Button, x+4 yp-1 w52 h24 Default hWndhPasteButton, Paste
			this.hPasteButton := hPasteButton
			BoundPaste := this.Paste.Bind(this)
			GuiControl, +g, %hPasteButton%, %BoundPaste%
			
			Gui, Add, Text, xm y+5 w30 h22 +0x200, Name: ; 0x200 for vcenter
			Gui, Add, Edit, x+5 yp w100 h22 hWndhPasteName, % this.Parent.Settings.DefaultName
			this.hPasteName := hPasteName
			
			Gui, Add, DropDownList, x+5 yp w75 hWndhPasteChan, Announce||IRC|Discord
			this.hPasteChan := hPasteChan
			
			PostMessage, 0x153, -1, 22-6,, ahk_id %hPasteChan% ; Set height of ComboBox
			Gui, Show,, % this.Parent.Title " - Pastebin"
			
			WinEvents.Register(this.hWnd, this)
		}
		
		GuiClose()
		{
			GuiControl, -g, % this.hPasteButton
			WinEvents.Unregister(this.hWnd)
			Gui, Destroy
		}
		
		GuiEscape()
		{
			this.GuiClose()
		}
		
		Paste()
		{
			GuiControlGet, PasteDesc,, % this.hPasteDesc
			GuiControlGet, PasteName,, % this.hPasteName
			GuiControlGet, PasteChan,, % this.hPasteChan
			this.GuiClose()
			
			Link := Ahkbin(this.Parent.RichCode.Value, PasteName, PasteDesc, this.Targets[PasteChan])
			
			MsgBox, 292, % this.Parent.Title " - Pasted", Link received:`n%Link%`n`nCopy to clipboard?
			IfMsgBox, Yes
				Clipboard := Link
		}
	}
	class Publish
	{
		__New(Parent)
		{
			this.Parent := Parent
			
			ParentWnd := this.Parent.hMainWindow
			Gui, New, +Owner%ParentWnd% +ToolWindow +hWndhWnd
			this.hWnd := hWnd
			Gui, Margin, 5, 5
			
			; 0x200 for vcenter
			Gui, Add, Text, w245 h22 Center +0x200, Gather all includes and save to file.
			
			Gui, Add, Checkbox, hWndhWnd w120 h22 Checked Section, Keep Comments
			this.hComments := hWnd
			Gui, Add, Checkbox, hWndhWnd w120 h22 Checked, Keep Indentation
			this.hIndent := hWnd
			Gui, Add, Checkbox, hWndhWnd w120 h22 Checked, Keep Empty Lines
			this.hEmpties := hWnd
			
			Gui, Add, Button, hWndhWnd w120 h81 ys-1 Default, Export
			this.hButton := hWnd
			BoundPublish := this.Publish.Bind(this)
			GuiControl, +g, %hWnd%, %BoundPublish%
			
			Gui, Show,, % this.Parent.Title " - Publish"
			
			WinEvents.Register(this.hWnd, this)
		}
		
		GuiClose()
		{
			GuiControl, -g, % this.hButton
			WinEvents.Unregister(this.hWnd)
			Gui, Destroy
		}
		
		GuiEscape()
		{
			this.GuiClose()
		}
		
		Publish()
		{
			GuiControlGet, KeepComments,, % this.hComments
			GuiControlGet, KeepIndent,, % this.hIndent
			GuiControlGet, KeepEmpties,, % this.hEmpties
			this.GuiClose()
			
			Gui, % this.Parent.hMainWindow ":+OwnDialogs"
			FileSelectFile, FilePath, S18,, % this.Parent.Title " - Publish Code"
			if ErrorLevel
				return
			
			FileOpen(FilePath, "w").Write(this.Parent.RichCode.Value)
			PreprocessScript(Text, FilePath, KeepComments, KeepIndent, KeepEmpties)
			FileOpen(FilePath, "w").Write(Text)
		}
	}
	class Find
	{
		__New(Parent)
		{
			this.Parent := Parent
			
			ParentWnd := this.Parent.hMainWindow
			Gui, New, +Owner%ParentWnd% +ToolWindow +hWndhWnd
			this.hWnd := hWnd
			Gui, Margin, 5, 5
			
			
			; Search
			Gui, Add, Edit, hWndhWnd w200
			SendMessage, 0x1501, True, &cue := "Search Text",, ahk_id %hWnd% ; EM_SETCUEBANNER
			this.hNeedle := hWnd
			
			Gui, Add, Button, yp-1 x+m w75 Default hWndhWnd, Find Next
			Bound := this.BtnFind.Bind(this)
			GuiControl, +g, %hWnd%, %Bound%
			
			Gui, Add, Button, yp x+m w75 hWndhWnd, Coun&t All
			Bound := this.BtnCount.Bind(this)
			GuiControl, +g, %hWnd%, %Bound%
			
			
			; Replace
			Gui, Add, Edit, hWndhWnd w200 xm Section
			SendMessage, 0x1501, True, &cue := "Replacement",, ahk_id %hWnd% ; EM_SETCUEBANNER
			this.hReplace := hWnd
			
			Gui, Add, Button, yp-1 x+m w75 hWndhWnd, &Replace
			Bound := this.Replace.Bind(this)
			GuiControl, +g, %hWnd%, %Bound%
			
			Gui, Add, Button, yp x+m w75 hWndhWnd, Replace &All
			Bound := this.ReplaceAll.Bind(this)
			GuiControl, +g, %hWnd%, %Bound%
			
			
			; Options
			Gui, Add, Checkbox, hWndhWnd xm, &Case Sensitive
			this.hOptCase := hWnd
			Gui, Add, Checkbox, hWndhWnd, Re&gular Expressions
			this.hOptRegEx := hWnd
			Gui, Add, Checkbox, hWndhWnd, Transform`, &Deref
			this.hOptDeref := hWnd
			
			
			Gui, Show,, % this.Parent.Title " - Find"
			
			WinEvents.Register(this.hWnd, this)
		}
		
		GuiClose()
		{
			GuiControl, -g, % this.hButton
			WinEvents.Unregister(this.hWnd)
			Gui, Destroy
		}
		
		GuiEscape()
		{
			this.GuiClose()
		}
		
		GetNeedle()
		{
			Opts := this.Case ? "`n" : "i`n"
			Opts .= this.Needle ~= "^[^\(]\)" ? "" : ")"
			if this.RegEx
				return Opts . this.Needle
			else
				return Opts "\Q" StrReplace(this.Needle, "\E", "\E\\E\Q") "\E"
		}
		
		Find(StartingPos:=1, WrapAround:=True)
		{
			Needle := this.GetNeedle()
			
			; Search from StartingPos
			NextPos := RegExMatch(this.Haystack, Needle, Match, StartingPos)
			
			; Search from the top
			if (!NextPos && WrapAround)
				NextPos := RegExMatch(this.Haystack, Needle, Match)
			
			return NextPos ? [NextPos, NextPos+StrLen(Match)] : False
		}
		
		Submit()
		{
			; Options
			GuiControlGet, Deref,, % this.hOptDeref
			GuiControlGet, Case,, % this.hOptCase
			this.Case := Case
			GuiControlGet, RegEx,, % this.hOptRegEx
			this.RegEx := RegEx
			
			; Search Text/Needle
			GuiControlGet, Needle,, % this.hNeedle
			if Deref
				Transform, Needle, Deref, %Needle%
			this.Needle := Needle
			
			; Replacement
			GuiControlGet, Replace,, % this.hReplace
			if Deref
				Transform, Replace, Deref, %Replace%
			this.Replace := Replace
			
			; Haystack
			this.Haystack := StrReplace(this.Parent.RichCode.Value, "`r")
		}
		
		BtnFind()
		{
			Gui, +OwnDialogs
			this.Submit()
			
			; Find and select the item or error out
			if (Pos := this.Find(this.Parent.RichCode.Selection[1]+2))
				this.Parent.RichCode.Selection := [Pos[1] - 1, Pos[2] - 1]
			else
				MsgBox, 0x30, % this.Parent.Title " - Find", Search text not found
		}
		
		BtnCount()
		{
			Gui, +OwnDialogs
			this.Submit()
			
			; Find and count all instances
			Count := 0, Start := 1
			while (Pos := this.Find(Start, False))
				Start := Pos[1]+1, Count += 1
			
			MsgBox, 0x40, % this.Parent.Title " - Find", %Count% instances found
		}
		
		Replace()
		{
			this.Submit()
			
			; Get the current selection
			Sel := this.Parent.RichCode.Selection
			
			; Find the next occurrence including the current selection
			Pos := this.Find(Sel[1]+1)
			
			; If the found item is already selected
			if (Sel[1]+1 == Pos[1] && Sel[2]+1 == Pos[2])
			{
				; Replace it
				this.Parent.RichCode.SelectedText := this.Replace
				
				; Update the haystack to include the replacement
				this.Haystack := StrReplace(this.Parent.RichCode.Value, "`r")
				
				; Find the next item *not* including the current selection
				Pos := this.Find(Sel[1]+StrLen(this.Replace)+1)
			}
			
			; Select the next found item or error out
			if Pos
				this.Parent.RichCode.Selection := [Pos[1] - 1, Pos[2] - 1]
			else
				MsgBox, 0x30, % this.Parent.Title " - Find", No more instances found
		}
		
		ReplaceAll()
		{
			rc := this.Parent.RichCode
			this.Submit()
			
			Needle := this.GetNeedle()
			
			; Replace the text in a way that pushes to the undo buffer
			rc.Frozen := True
			Sel := rc.Selection
			rc.Selection := [0, -1]
			rc.SelectedText := RegExReplace(this.Haystack, Needle, this.Replace, Count)
			rc.Selection := Sel
			rc.Frozen := False
			
			MsgBox, 0x40, % this.Parent.Title " - Find", %Count% instances replaced
		}
	}
	class ScriptOpts
	{
		__New(Parent)
		{
			this.Parent := Parent
			
			; Create a GUI
			ParentWnd := this.Parent.hMainWindow
			Gui, New, +Owner%ParentWnd% +ToolWindow +hWndhWnd
			this.hWnd := hWnd
			WinEvents.Register(this.hWnd, this)
			
			; Add path picker button
			Gui, Add, Button, xm ym w95 hWndhButton, Pick AHK Path
			BoundSelectFile := this.SelectFile.Bind(this)
			GuiControl, +g, %hButton%, %BoundSelectFile%
			
			; Add path visualization field
			Gui, Add, Edit, ym w250 ReadOnly hWndhAhkPath, % this.Parent.Settings.AhkPath
			this.hAhkPath := hAhkPath
			
			; Add parameters field
			Gui, Add, Text, xm w95 h22 +0x200, Parameters:
			Gui, Add, Edit, yp x+m w250 hWndhParamEdit
			this.hParamEdit := hParamEdit
			ParamEditBound := this.ParamEdit.Bind(this)
			GuiControl, +g, %hParamEdit%, %ParamEditBound%
			
			; Add Working Directory field
			Gui, Add, Button, xm w95 hWndhWDButton, Pick Working Dir
			BoundSelectPath := this.SelectPath.Bind(this)
			GuiControl, +g, %hWDButton%, %BoundSelectPath%
			
			; Add Working Dir visualization field
			Gui, Add, Edit, x+m w250 ReadOnly hWndhWorkingDir, %A_WorkingDir%
			this.hWorkingDir := hWorkingDir
			
			; Show the GUI
			Gui, Show,, % this.Parent.Title " - Script Options"
		}
		
		ParamEdit()
		{
			GuiControlGet, ParamEdit,, % this.hParamEdit
			this.Parent.Settings.Params := ParamEdit
		}
		
		SelectFile()
		{
			GuiControlGet, AhkPath,, % this.hAhkPath
			FileSelectFile, AhkPath, 1, %AhkPath%, Pick an AHK EXE, Executables (*.exe)
			if !AhkPath
				return
			this.Parent.Settings.AhkPath := AhkPath
			GuiControl,, % this.hAhkPath, %AhkPath%
		}
		
		SelectPath()
		{
			FileSelectFolder, WorkingDir, *%A_WorkingDir%, 0, Choose the Working Directory
			if !WorkingDir
				return
			SetWorkingDir, %WorkingDir%
			this.UpdateFields()
		}
		
		UpdateFields()
		{
			GuiControl,, % this.hWorkingDir, %A_WorkingDir%
		}
		
		GuiClose()
		{
			WinEvents.Unregister(this.hWnd)
			Gui, Destroy
		}
		
		GuiEscape()
		{
			this.GuiClose()
		}
	}
	class MenuButtons
	{
		__New(Parent)
		{
			this.Parent := Parent
		}
		
		Save(SaveAs)
		{
			if (SaveAs || !this.Parent.FilePath)
			{
				Gui, +OwnDialogs
				FileSelectFile, FilePath, S18,, % this.Parent.Title " - Save Code"
				if ErrorLevel
					return
				this.Parent.FilePath := FilePath
			}
			
			FileOpen(this.Parent.FilePath, "w").Write(this.Parent.RichCode.Value)
			
			this.Parent.RichCode.Modified := False
			this.Parent.UpdateStatusBar()
		}
		
		Rename()
		{
			; Make sure the opened file still exists
			if !InStr(FileExist(this.Parent.FilePath), "A")
				throw Exception("Opened file no longer exists")
			
			; Ask where to move it to
			FileSelectFile, FilePath, S10, % this.Parent.FilePath
			, Rename As, AutoHotkey Scripts (*.ahk)
			if InStr(FileExist(FilePath), "A")
				throw Exception("Destination file already exists")
			
			; Attempt to move it
			FileMove, % this.Parent.FilePath, % FilePath
			if ErrorLevel
				throw Exception("Failed to rename file")
			this.Parent.FilePath := FilePath
		}
		
		Open()
		{
			Gui, +OwnDialogs
			FileSelectFile, FilePath, 3,, % this.Parent.Title " - Open Code"
			if ErrorLevel
				return
			this.Parent.LoadCode(FileOpen(FilePath, "r").Read(), FilePath)
			
			; Follow the directory of the most recently opened file
			SetWorkingDir, %FilePath%\..
			this.Parent.ScriptOpts.UpdateFields()
		}
		
		OpenFolder()
		{
			Run, explorer.exe "%A_WorkingDir%"
		}
		
		New()
		{
			Run, "%A_AhkPath%" "%A_ScriptFullPath%"
		}
		
		Publish()
		{ ; TODO: Recycle PubInstance
			if WinExist("ahk_id" this.PubInstance.hWnd)
				WinActivate, % "ahk_id" this.PubInstance.hWnd
			else
				this.PubInstance := new this.Parent.Publish(this.Parent)
		}
		
		Fetch()
		{
			Gui, +OwnDialogs
			InputBox, Url, % this.Parent.Title " - Fetch Code", Enter a URL to fetch code from.
			if (Url := Trim(Url))
				this.Parent.LoadCode(UrlDownloadToVar(Url))
		}
		
		Find()
		{ ; TODO: Recycle FindInstance
			if WinExist("ahk_id" this.FindInstance.hWnd)
				WinActivate, % "ahk_id" this.FindInstance.hWnd
			else
				this.FindInstance := new this.Parent.Find(this.Parent)
		}
		
		Paste()
		{ ; TODO: Recycle PasteInstance
			if WinExist("ahk_id" this.PasteInstance.hWnd)
				WinActivate, % "ahk_id" this.PasteInstance.hWnd
			else
				this.PasteInstance := new this.Parent.Paste(this.Parent)
		}
		
		ScriptOpts()
		{
			if WinExist("ahk_id" this.Parent.ScriptOptsInstance.hWnd)
				WinActivate, % "ahk_id" this.Parent.ScriptOptsInstance.hWnd
			else
				this.Parent.ScriptOptsInstance := new this.Parent.ScriptOpts(this.Parent)
		}
		
		ToggleOnTop()
		{
			if (this.Parent.AlwaysOnTop := !this.Parent.AlwaysOnTop)
			{
				Menu, % this.Parent.Menus[4], Check, &AlwaysOnTop`tAlt+A
				Gui, +AlwaysOnTop
			}
			else
			{
				Menu, % this.Parent.Menus[4], Uncheck, &AlwaysOnTop`tAlt+A
				Gui, -AlwaysOnTop
			}
		}
		
		Highlighter()
		{
			if (this.Parent.Settings.UseHighlighter := !this.Parent.Settings.UseHighlighter)
				Menu, % this.Parent.Menus[4], Check, &Highlighter
			else
				Menu, % this.Parent.Menus[4], Uncheck, &Highlighter
			
			; Force refresh the code, adding/removing any highlighting
			this.Parent.RichCode.Value := this.Parent.RichCode.Value
		}
		
		GlobalRun()
		{
			if (this.Parent.Settings.GlobalRun := !this.Parent.Settings.GlobalRun)
				Menu, % this.Parent.Menus[4], Check, Global Run Hotkeys
			else
				Menu, % this.Parent.Menus[4], Uncheck, Global Run Hotkeys
		}
		
		AutoIndent()
		{
			this.Parent.LoadCode(AutoIndent(this.Parent.RichCode.Value
			, this.Parent.Settings.Indent), this.Parent.FilePath)
		}
		
		Help()
		{
			HelpFile.Open(this.Parent.GetKeywordFromCaret())
		}
		
		About()
		{
			Gui, +OwnDialogs
			MsgBox,, % this.Parent.Title " - About", CodeQuickTester written by GeekDude
		}
		
		ServiceHandler()
		{
			Gui, +OwnDialogs
			if ServiceHandler.Installed()
			{
				MsgBox, 36, % this.Parent.Title " - Uninstall Service Handler"
				, Are you sure you want to remove CodeQuickTester from being the default service handler for "ahk:" links?
				IfMsgBox, Yes
				{
					ServiceHandler.Remove()
					Menu, % this.Parent.Menus[4], Uncheck, Install Service Handler
				}
			}
			else
			{
				MsgBox, 36, % this.Parent.Title " - Install Service Handler"
				, Are you sure you want to install CodeQuickTester as the default service handler for "ahk:" links?
				IfMsgBox, Yes
				{
					ServiceHandler.Install()
					Menu, % this.Parent.Menus[4], Check, Install Service Handler
				}
			}
		}
		
		DefaultEditor()
		{
			Gui, +OwnDialogs
			
			if !A_IsAdmin
			{
				MsgBox, 48, % this.Parent.Title " - Change Editor", You must be running as administrator to use this feature.
				return
			}
			
			RegRead, Editor, HKCR, AutoHotkeyScript\Shell\Edit\Command
			if (Editor == this.Parent.EditorString)
			{
				MsgBox, 36, % this.Parent.Title " - Remove as Default Editor"
				, % "Are you sure you want to restore the original default editor for .ahk files?"
				. "`n`n" this.Parent.OrigEditorString
				IfMsgBox, Yes
				{
					RegWrite REG_SZ, HKCR, AutoHotkeyScript\Shell\Edit\Command,, % this.Parent.OrigEditorString
					Menu, % this.Parent.Menus[4], Uncheck, Set as Default Editor
				}
			}
			else
			{
				MsgBox, 36, % this.Parent.Title " - Set as Default Editor"
				, % "Are you sure you want to install CodeQuickTester as the default editor for .ahk files?"
				. "`n`n" this.Parent.EditorString
				IfMsgBox, Yes
				{
					RegWrite REG_SZ, HKCR, AutoHotkeyScript\Shell\Edit\Command,, % this.Parent.EditorString
					MsgBox, %ErrorLevel%
					Menu, % this.Parent.Menus[4], Check, Set as Default Editor
				}
			}
			
			
		}
		
		Comment()
		{
			this.Parent.RichCode.IndentSelection(False, ";")
		}
		
		Uncomment()
		{
			this.Parent.RichCode.IndentSelection(True, ";")
		}
		
		Indent()
		{
			this.Parent.RichCode.IndentSelection()
		}
		
		Unindent()
		{
			this.Parent.RichCode.IndentSelection(True)
		}
		
		IncludeRel()
		{
			FileSelectFile, AbsPath, 1,, Pick a script to include, AutoHotkey Scripts (*.ahk)
			if ErrorLevel
				return
			
			; Get the relative path
			VarSetCapacity(RelPath, A_IsUnicode?520:260) ; MAX_PATH
			if !DllCall("Shlwapi.dll\PathRelativePathTo"
				, "Str", RelPath                    ; Out Directory
				, "Str", A_WorkingDir, "UInt", 0x10 ; From Directory
				, "Str", AbsPath, "UInt", 0x10)     ; To Directory
				throw Exception("Relative path could not be found")
			
			; Select the start of the line
			RC := this.Parent.RichCode
			Top := RC.SendMsg(0x436, 0, RC.Selection[1]) ; EM_EXLINEFROMCHAR
			TopLineIndex := RC.SendMsg(0xBB, Top, 0) ; EM_LINEINDEX
			RC.Selection := [TopLineIndex, TopLineIndex]
			
			; Insert the include
			RC.SelectedText := "#Include " RelPath "`n"
			RC.Selection[1] := RC.Selection[2]
		}
		
		IncludeAbs()
		{
			FileSelectFile, AbsPath, 1,, Pick a script to include, AutoHotkey Scripts (*.ahk)
			if ErrorLevel
				return
			
			; Select the start of the line
			RC := this.Parent.RichCode
			Top := RC.SendMsg(0x436, 0, RC.Selection[1]) ; EM_EXLINEFROMCHAR
			TopLineIndex := RC.SendMsg(0xBB, Top, 0) ; EM_LINEINDEX
			RC.Selection := [TopLineIndex, TopLineIndex]
			
			; Insert the include
			RC.SelectedText := "#Include " AbsPath "`n"
			RC.Selection[1] := RC.Selection[2]
		}
		
		AutoComplete()
		{
			if (this.Parent.Settings.UseAutoComplete := !this.Parent.Settings.UseAutoComplete)
				Menu, % this.Parent.Menus[4], Check, AutoComplete
			else
				Menu, % this.Parent.Menus[4], Uncheck, AutoComplete
			
			this.Parent.AC.Enabled := this.Parent.Settings.UseAutoComplete
		}
	}
	/*
		Implements functionality necessary for AutoCompletion of keywords in the
		RichCode control. Currently works off of values stored in the provided
		Parent object, but could be modified to work off a provided RichCode
		instance directly.
		
		The class is mostly self contained and could be easily extended to other
		projects, and even other types of controls. The main method of interacting
		with the class is by passing it WM_KEYDOWN messages. Another way to interact
		is by modifying the Fragment property, especially to clear it when you want
		to cancel autocompletion.
		
		Depends on CQT.ahk, and optionally on HelpFile.ahk
	*/
	
	class AutoComplete
	{
		; Maximum number of suggestions to be displayed in the dialog
		static MaxSuggestions := 9
		
		; Minimum length for a word to be entered into the word list
		static MinWordLen := 4
		
		; Minimum length of fragment before suggestions should be displayed
		static MinSuggestLen := 3
		
		; Stores the initial caret position for newly typed fragments
		static CaretX := 0, CaretY := 0
		
		
		; --- Properties ---
		
		Fragment[]
		{
			get
			{
				return this._Fragment
			}
			
			set
			{
				this._Fragment := Value
				
				; Give suggestions when a fragment of sufficient
				; length has been provided
				if (StrLen(this._Fragment) >= 3)
					this._Suggest()
				else
					this._Hide()
				
				return this._Fragment
			}
		}
		
		Enabled[]
		{
			get
			{
				return this._Enabled
			}
			
			set
			{
				this._Enabled := Value
				if (Value)
					this.BuildWordList()
				else
					this.Fragment := ""
				return Value
			}
		}
		
		
		; --- Constructor, Destructor ---
		
		__New(Parent, Enabled:=True)
		{
			this.Parent := Parent
			this.Enabled := Enabled
			this.WineVer := DllCall("ntdll.dll\wine_get_version", "AStr")
			
			; Create the tool GUI for the floating list
			hParentWnd := this.Parent.hMainWindow
			Gui, +hWndhDefaultWnd
			Relation := this.WineVer ? "Parent" Parent.RichCode.hWnd : "Owner" Parent.hMainWindow
			Gui, New, +%Relation% -Caption +ToolWindow +hWndhWnd
			this.hWnd := hWnd
			Gui, Margin, 0, 0
			
			; Create the ListBox control withe appropriate font and styling
			Font := this.Parent.Settings.Font
			Gui, Font, % "s" Font.Size, % Font.Typeface
			Gui, Add, ListBox, x0 y0 r1 0x100 AltSubmit hWndhListBox, Item
			this.hListBox := hListBox
			
			; Finish GUI creation and restore the default GUI
			Gui, Show, Hide, % this.Parent.Title " - AutoComplete"
			Gui, %hDefaultWnd%:Default
			
			; Get relevant dimensions of the ListBox for later resizing
			SendMessage, 0x1A1, 0, 0,, % "ahk_id" this.hListBox ; LB_GETITEMHEIGHT
			this.ListBoxItemHeight := ErrorLevel
			VarSetCapacity(ListBoxRect, 16, 0)
			DllCall("User32.dll\GetClientRect", "Ptr", this.hListBox, "Ptr", &ListBoxRect)
			this.ListBoxMargins := NumGet(ListBoxRect, 12, "Int") - this.ListBoxItemHeight
			
			; Set up the GDI Device Context for later text measurement in _GetWidth
			this.hDC := DllCall("GetDC", "UPtr", this.hListBox, "UPtr")
			SendMessage, 0x31, 0, 0,, % "ahk_id" this.hListBox ; WM_GETFONT
			this.hFont := DllCall("SelectObject", "UPtr", this.hDC, "UPtr", ErrorLevel, "UPtr")
			
			; Record the total screen width for later user. If the monitors get
			; rearranged while the script is still running this value will be
			; inaccurate. However, this will not likely be a significant issue,
			; and the issues caused by it would be minimal.
			SysGet, ScreenWidth, 78
			this.ScreenWidth := ScreenWidth
			
			; Pull a list of default words from the help file.
			; TODO: Include some kind of hard-coded list for when the help file is
			;       not present, or to supplement the help file.
			for Key in HelpFile.GetLookup()
				this.DefaultWordList .= "|" LTrim(Key, "#")
			
			; Build the initial word list based on the default words and the
			; RichCode's contents at the time of AutoComplete's initialization
			this.BuildWordList()
		}
		
		__Delete()
		{
			Gui, % this.hWnd ":Destroy"
			this.Visible := False
			DllCall("SelectObject", "UPtr", this.hDC, "UPtr", this.hFont, "UPtr")
			DllCall("ReleaseDC", "UPtr", this.hListBox, "UPtr", this.hDC)
		}
		
		
		; --- Private Methods ---
		
		; Gets the pixel-based width of a provided text snippet using the GDI font
		; selected into the ListBox control
		_GetWidth(Text)
		{
			MaxWidth := 0
			Loop, Parse, Text, |
			{
				DllCall("GetTextExtentPoint32", "UPtr", this.hDC, "Str", A_LoopField
				, "Int", StrLen(A_LoopField), "Int64*", Size), Size &= 0xFFFFFFFF
				
				if (Size > MaxWidth)
					MaxWidth := Size
			}
			
			return MaxWidth
		}
		
		; Shows the suggestion dialog with contents of the provided DisplayList
		_Show(DisplayList)
		{
			; Insert the new list
			GuiControl,, % this.hListBox, %DisplayList%
			GuiControl, Choose, % this.hListBox, 1
			
			; Resize to fit contents
			StrReplace(DisplayList, "|",, Rows)
			Height := Rows * this.ListBoxItemHeight + this.ListBoxMargins
			Width := this._GetWidth(DisplayList) + 10
			GuiControl, Move, % this.hListBox, w%Width% h%Height%
			
			; Keep the dialog from running off the screen
			X := this.CaretX, Y := this.CaretY + 20
			if ((X + Width) > this.ScreenWidth)
				X := this.ScreenWidth - Width
			
			; Make the dialog visible
			Gui, % this.hWnd ":Show", x%X% y%Y% AutoSize NoActivate
			this.Visible := True
		}
		
		; Hides the dialog if it is visible
		_Hide()
		{
			if !this.Visible
				return
			
			Gui, % this.hWnd ":Hide"
			this.Visible := False
		}
		
		; Filters the word list for entries starting with the fragment, then
		; shows the dialog with the filtered list as suggestions
		_Suggest()
		{
			; Filter the list for words beginning with the fragment
			Suggestions := LTrim(RegExReplace(this.WordList
			, "i)\|(?!" this.Fragment ")[^\|]+"), "|")
			
			; Fail out if there were no matches
			if !Suggestions
				return true, this._Hide()
			
			; Pull the first MaxSuggestions suggestions
			if (Pos := InStr(Suggestions, "|",,, this.MaxSuggestions))
				Suggestions := SubStr(Suggestions, 1, Pos-1)
			this.Suggestions := Suggestions
			
			this._Show("|" Suggestions)
		}
		
		; Finishes the fragment with the selected suggestion
		_Complete()
		{
			; Get the text of the selected item
			GuiControlGet, Selected,, % this.hListBox
			Suggestion := StrSplit(this.Suggestions, "|")[Selected]
			
			; Replace fragment preceding cursor with selected suggestion
			RC := this.Parent.RichCode
			RC.Selection[1] -= StrLen(this.Fragment)
			RC.SelectedText := Suggestion
			RC.Selection[1] := RC.Selection[2]
			
			; Clear out the fragment in preparation for further typing
			this.Fragment := ""
		}
		
		
		; --- Public Methods ---
		
		; Interpret WM_KEYDOWN messages, the primary means of interfacing with the
		; class. These messages can be provided by registering an appropriate
		; handler with OnMessage, or by forwarding the events from another handler
		; for the control.
		WM_KEYDOWN(wParam, lParam)
		{
			if (!this._Enabled)
				return
			
			; Get the name of the key using the virtual key code. The key's scan
			; code is not used here, but is available in bits 16-23 of lParam and
			; could be used in future versions for greater reliability.
			Key := GetKeyName(Format("vk{:02x}", wParam))
			
			; Treat Numpad variants the same as the equivalent standard keys
			Key := StrReplace(Key, "Numpad")
			
			; Handle presses meant to interact with the dialog, such as
			; navigational, confirmational, or dismissive commands.
			if (this.Visible)
			{
				if (Key == "Tab" || Key == "Enter")
					return False, this._Complete()
				else if (Key == "Up")
					return False, this.SelectUp()
				else if (Key == "Down")
					return False, this.SelectDown()
			}
			
			; Ignore standalone modifier presses, and some modified regular presses
			if Key in Shift,Control,Alt
				return
			
			; Reset on presses with the control modifier
			if GetKeyState("Control")
				return "", this.Fragment := ""
			
			; Subtract from the end of fragment on backspace
			if (Key == "Backspace")
				return "", this.Fragment := SubStr(this.Fragment, 1, -1)
			
			; Apply Shift and CapsLock
			if GetKeyState("Shift")
				Key := StrReplace(Key, "-", "_")
			if (GetKeyState("Shift") ^ GetKeyState("CapsLock", "T"))
				Key := Format("{:U}", Key)
			
			; Reset on unwanted presses -- Allow numbers but not at beginning
			if !(Key ~= "^[A-Za-z_]$" || (this.Fragment != "" && Key ~= "^[0-9]$"))
				return "", this.Fragment := ""
			
			; Record the starting position of new fragments
			if (this.Fragment == "")
			{
				CoordMode, Caret, % this.WineVer ? "Client" : "Screen"
				
				; Round "" to 0, which can prevent errors in the unlikely case that
				; input is received while the control is not focused.
				this.CaretX := Round(A_CaretX), this.CaretY := Round(A_CaretY)
			}
			
			; Update fragment with the press
			this.Fragment .= Key
		}
		
		; Triggers a rebuild of the word list from the RichCode control's contents
		BuildWordList()
		{
			if (!this._Enabled)
				return
			
			; Replace non-word chunks with delimiters
			List := RegExReplace(this.Parent.RichCode.Value, "\W+", "|")
			
			; Ignore numbers at the beginning of words
			List := RegExReplace(List, "\b[0-9]+")
			
			; Ignore words that are too small
			List := RegExReplace(List, "\b\w{1," this.MinWordLen-1 "}\b")
			
			; Append default entries, remove duplicates, and save the list
			List .= this.DefaultWordList
			Sort, List, U D| Z
			this.WordList := "|" Trim(List, "|")
		}
		
		; Moves the selected item in the dialog up one position
		SelectUp()
		{
			GuiControlGet, Selected,, % this.hListBox
			if (--Selected < 1)
				Selected := this.MaxSuggestions
			GuiControl, Choose, % this.hListBox, %Selected%
		}
		
		; Moves the selected item in the dialog down one position
		SelectDown()
		{
			GuiControlGet, Selected,, % this.hListBox
			if (++Selected > this.MaxSuggestions)
				Selected := 1
			GuiControl, Choose, % this.hListBox, %Selected%
		}
	}
}
class ServiceHandler ; static class
{
	static Protocol := "ahk"
	
	Install()
	{
		Protocol := this.Protocol
		RegWrite, REG_SZ, HKCU, Software\Classes\%Protocol%,, URL:AHK Script Protocol
		RegWrite, REG_SZ, HKCU, Software\Classes\%Protocol%, URL Protocol
		RegWrite, REG_SZ, HKCU, Software\Classes\%Protocol%\shell\open\command,, "%A_AhkPath%" "%A_ScriptFullPath%" "`%1"
	}
	
	Remove()
	{
		Protocol := this.Protocol
		RegDelete, HKCU, Software\Classes\%Protocol%
	}
	
	Installed()
	{
		Protocol := this.Protocol
		RegRead, Out, HKCU, Software\Classes\%Protocol%
		return !ErrorLevel
	}
}
class WinEvents ; static class
{
	static _ := WinEvents.AutoInit()
	
	AutoInit()
	{
		this.Table := []
		OnMessage(2, this.Destroy.bind(this))
	}
	
	Register(ID, HandlerClass, Prefix="Gui")
	{
		Gui, %ID%: +hWndhWnd +LabelWinEvents_
		this.Table[hWnd] := {Class: HandlerClass, Prefix: Prefix}
	}
	
	Unregister(ID)
	{
		Gui, %ID%: +hWndhWnd
		this.Table.Delete(hWnd)
	}
	
	Dispatch(Type, Params*)
	{
		Info := this.Table[Params[1]]
		return (Info.Class)[Info.Prefix . Type](Params*)
	}
	
	Destroy(wParam, lParam, Msg, hWnd)
	{
		this.Table.Delete(hWnd)
	}
}

WinEvents_Close(Params*) {
	return WinEvents.Dispatch("Close", Params*)
} WinEvents_Escape(Params*) {
	return WinEvents.Dispatch("Escape", Params*)
} WinEvents_Size(Params*) {
	return WinEvents.Dispatch("Size", Params*)
} WinEvents_ContextMenu(Params*) {
	return WinEvents.Dispatch("ContextMenu", Params*)
} WinEvents_DropFiles(Params*) {
	return WinEvents.Dispatch("DropFiles", Params*)
}
AutoIndent(Code, Indent = "`t", Newline = "`r`n")
{
	IndentRegEx =
	( LTrim Join
	Catch|else|for|Finally|if|IfEqual|IfExist|
	IfGreater|IfGreaterOrEqual|IfInString|
	IfLess|IfLessOrEqual|IfMsgBox|IfNotEqual|
	IfNotExist|IfNotInString|IfWinActive|IfWinExist|
	IfWinNotActive|IfWinNotExist|Loop|Try|while
	)
	
	; Lock and Block are modified ByRef by Current
	Lock := [], Block := []
	ParentIndent := Braces := 0
	ParentIndentObj := []
	
	for each, Line in StrSplit(Code, "`n", "`r")
	{
		Text := Trim(RegExReplace(Line, "\s;.*")) ; Comment removal
		First := SubStr(Text, 1, 1), Last := SubStr(Text, 0, 1)
		FirstTwo := SubStr(Text, 1, 2)
		
		IsExpCont := (Text ~= "i)^\s*(&&|OR|AND|\.|\,|\|\||:|\?)")
		IndentCheck := (Text ~= "iA)}?\s*\b(" IndentRegEx ")\b")
		
		if (First == "(" && Last != ")")
			Skip := True
		if (Skip)
		{
			if (First == ")")
				Skip := False
			Out .= Newline . RTrim(Line)
			continue
		}
		
		if (FirstTwo == "*/")
			Block := [], ParentIndent := 0
		
		if Block.MinIndex()
			Current := Block, Cur := 1
		else
			Current := Lock, Cur := 0
		
		; Round converts "" to 0
		Braces := Round(Current[Current.MaxIndex()].Braces)
		ParentIndent := Round(ParentIndentObj[Cur])
		
		if (First == "}")
		{
			while ((Found := SubStr(Text, A_Index, 1)) ~= "}|\s")
			{
				if (Found ~= "\s")
					continue
				if (Cur && Current.MaxIndex() <= 1)
					break
				Special := Current.Pop().Ind, Braces--
			}
		}
		
		if (First == "{" && ParentIndent)
			ParentIndent--
		
		Out .= Newline
		Loop, % Special ? Special-1 : Round(Current[Current.MaxIndex()].Ind) + Round(ParentIndent)
			Out .= Indent
		Out .= Trim(Line)
		
		if (FirstTwo == "/*")
		{
			if (!Block.MinIndex())
			{
				Block.Push({ParentIndent: ParentIndent
				, Ind: Round(Lock[Lock.MaxIndex()].Ind) + 1
				, Braces: Round(Lock[Lock.MaxIndex()].Braces) + 1})
			}
			Current := Block, ParentIndent := 0
		}
		
		if (Last == "{")
		{
			Braces++, ParentIndent := (IsExpCont && Last == "{") ? ParentIndent-1 : ParentIndent
			Current.Push({Braces: Braces
			, Ind: ParentIndent + Round(Current[Current.MaxIndex()].ParentIndent) + Braces
			, ParentIndent: ParentIndent + Round(Current[Current.MaxIndex()].ParentIndent)})
			ParentIndent := 0
		}
		
		if ((ParentIndent || IsExpCont || IndentCheck) && (IndentCheck && Last != "{"))
			ParentIndent++
		if (ParentIndent > 0 && !(IsExpCont || IndentCheck))
			ParentIndent := 0
		
		ParentIndentObj[Cur] := ParentIndent
		Special := 0
	}
	
	if Braces
		throw Exception("Segment Open!")
	
	return SubStr(Out, StrLen(Newline)+1)
}
; Modified from https://github.com/cocobelgica/AutoHotkey-Util/blob/master/ExecScript.ahk
ExecScript(Script, Params="", AhkPath="")
{
	static Shell := ComObjCreate("WScript.Shell")
	Name := "\\.\pipe\AHK_CQT_" A_TickCount
	Pipe := []
	Loop, 3
	{
		Pipe[A_Index] := DllCall("CreateNamedPipe"
		, "Str", Name
		, "UInt", 2, "UInt", 0
		, "UInt", 255, "UInt", 0
		, "UInt", 0, "UPtr", 0
		, "UPtr", 0, "UPtr")
	}
	if !FileExist(AhkPath)
		throw Exception("AutoHotkey runtime not found: " AhkPath)
	if (A_IsCompiled && AhkPath == A_ScriptFullPath)
		AhkPath .= " /E"
	if FileExist(Name)
	{
		Exec := Shell.Exec(AhkPath " /CP65001 " Name " " Params)
		DllCall("ConnectNamedPipe", "UPtr", Pipe[2], "UPtr", 0)
		DllCall("ConnectNamedPipe", "UPtr", Pipe[3], "UPtr", 0)
		FileOpen(Pipe[3], "h", "UTF-8").Write(Script)
	}
	else ; Running under WINE with improperly implemented pipes
	{
		FileOpen(Name := "AHK_CQT_TMP.ahk", "w").Write(Script)
		Exec := Shell.Exec(AhkPath " /CP65001 " Name " " Params)
	}
	Loop, 3
		DllCall("CloseHandle", "UPtr", Pipe[A_Index])
	return Exec
}

DeHashBang(Script)
{
	AhkPath := A_AhkPath
	if RegExMatch(Script, "`a)^\s*`;#!\s*(.+)", Match)
	{
		AhkPath := Trim(Match1)
		Vars := {"%A_ScriptDir%": A_WorkingDir
		, "%A_WorkingDir%": A_WorkingDir
		, "%A_AppData%": A_AppData
		, "%A_AppDataCommon%": A_AppDataCommon
		, "%A_LineFile%": A_ScriptFullPath
		, "%A_AhkPath%": A_AhkPath
		, "%A_AhkDir%": A_AhkPath "\.."}
		for SearchText, Replacement in Vars
			StringReplace, AhkPath, AhkPath, %SearchText%, %Replacement%, All
	}
	return AhkPath
}

UrlDownloadToVar(Url)
{
	xhr := ComObjCreate("MSXML2.XMLHTTP")
	xhr.Open("GET", url, false), xhr.Send()
	return xhr.ResponseText
}

; Helper function, to make passing in expressions resulting in function objects easier
SetTimer(Label, Period)
{
	SetTimer, %Label%, %Period%
}

SendMessage(Msg, wParam, lParam, hWnd)
{
	; DllCall("SendMessage", "UPtr", hWnd, "UInt", Msg, "UPtr", wParam, "Ptr", lParam, "UPtr")
	SendMessage, Msg, wParam, lParam,, ahk_id %hWnd%
	return ErrorLevel
}

PostMessage(Msg, wParam, lParam, hWnd)
{
	PostMessage, Msg, wParam, lParam,, ahk_id %hWnd%
	return ErrorLevel
}

Ahkbin(Content, Name="", Desc="", Channel="")
{
	static URL := "https://p.ahkscript.org/"
	Form := "code=" UriEncode(Content)
	if Name
		Form .= "&name=" UriEncode(Name)
	if Desc
		Form .= "&desc=" UriEncode(Desc)
	if Channel
		Form .= "&announce=on&channel=" UriEncode(Channel)
	
	Pbin := ComObjCreate("MSXML2.XMLHTTP")
	Pbin.Open("POST", URL, False)
	Pbin.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
	Pbin.Send(Form)
	return Pbin.getResponseHeader("ahk-location")
}

; Modified by GeekDude from http://goo.gl/0a0iJq
UriEncode(Uri, RE="[0-9A-Za-z]") {
	VarSetCapacity(Var, StrPut(Uri, "UTF-8"), 0), StrPut(Uri, &Var, "UTF-8")
	While Code := NumGet(Var, A_Index - 1, "UChar")
		Res .= (Chr:=Chr(Code)) ~= RE ? Chr : Format("%{:02X}", Code)
	Return, Res
}

CreateMenus(Menu)
{
	static MenuName := 0
	Menus := ["Menu_" MenuName++]
	for each, Item in Menu
	{
		Ref := Item[2]
		if IsObject(Ref) && Ref._NewEnum()
		{
			SubMenus := CreateMenus(Ref)
			Menus.Push(SubMenus*), Ref := ":" SubMenus[1]
		}
		Menu, % Menus[1], Add, % Item[1], %Ref%
	}
	return Menus
}

Ini_Load(Contents)
{
	Section := Out := []
	loop, Parse, Contents, `n, `r
	{
		if ((Line := Trim(A_LoopField)) ~= "^;|^$")
			continue
		else if RegExMatch(Line, "^\[(.+)\]$", Match)
			Out[Match1] := (Section := [])
		else if RegExMatch(Line, "^(.+?)=(.*)$", Match)
			Section[Trim(Match1)] := Trim(Match2)
	}
	return Out
}

GetFullPathName(FilePath)
{
	VarSetCapacity(Path, A_IsUnicode ? 520 : 260, 0)
	DllCall("GetFullPathName", "Str", FilePath
	, "UInt", 260, "Str", Path, "Ptr", 0, "UInt")
	return Path
}

RichEdit_AddMargins(hRichEdit, x:=0, y:=0, w:=0, h:=0)
{
	static WineVer := DllCall("ntdll.dll\wine_get_version", "AStr")
	VarSetCapacity(RECT, 16, 0)
	if (x | y | w | h)
	{
		if WineVer
		{
			; Workaround for bug in Wine 3.0.2.
			; This code will need to be updated this code
			; after future Wine releases that fix it.
			NumPut(x, RECT,  0, "Int"), NumPut(y, RECT,  4, "Int")
			NumPut(w, RECT,  8, "Int"), NumPut(h, RECT, 12, "Int")
		}
		else
		{
			if !DllCall("GetClientRect", "UPtr", hRichEdit, "UPtr", &RECT, "UInt")
				throw Exception("Couldn't get RichEdit Client RECT")
			NumPut(x + NumGet(RECT,  0, "Int"), RECT,  0, "Int")
			NumPut(y + NumGet(RECT,  4, "Int"), RECT,  4, "Int")
			NumPut(w + NumGet(RECT,  8, "Int"), RECT,  8, "Int")
			NumPut(h + NumGet(RECT, 12, "Int"), RECT, 12, "Int")
		}
	}
	SendMessage(0xB3, 0, &RECT, hRichEdit)
}

;
; Based on code from fincs' Ahk2Exe - https://github.com/fincs/ahk2exe
;

PreprocessScript(ByRef ScriptText, AhkScript, KeepComments=1, KeepIndent=1, KeepEmpties=0, FileList="", FirstScriptDir="", Options="", iOption=0)
{
	SplitPath, AhkScript, ScriptName, ScriptDir
	if !IsObject(FileList)
	{
		FileList := [AhkScript]
		; ScriptText := "; <COMPILER: v" A_AhkVersion ">`n"
		FirstScriptDir := A_WorkingDir
		IsFirstScript := true
		Options := { comm: ";", esc: "``" }
		
		OldWorkingDir := A_WorkingDir
		SetWorkingDir, %FirstScriptDir%
	}
	
	IfNotExist, %AhkScript%
		if !iOption
			Util_Error((IsFirstScript ? "Script" : "#include") " file """ AhkScript """ cannot be opened.")
	else return
		
	cmtBlock := false, contSection := false
	Loop, Read, %AhkScript%
	{
		tline := Trim(A_LoopReadLine)
		RegExMatch(A_LoopReadLine, "^[ \t]+", indent)
		if !cmtBlock
		{
			if !contSection
			{
				if StrStartsWith(tline, Options.comm) && !KeepComments
					continue
				else if (tline = "" && !KeepEmpties)
					continue
				else if StrStartsWith(tline, "/*")
				{
					if KeepComments
						ScriptText .= A_LoopReadLine "`n"
					cmtBlock := true
					continue
				}
			}
			if StrStartsWith(tline, "(") && !IsFakeCSOpening(tline)
				contSection := true
			else if StrStartsWith(tline, ")")
				contSection := false
			
			ttline := RegExReplace(tline, "\s+" RegExEscape(Options.comm) ".*$", "")
			if !contSection && RegExMatch(ttline, "i)^#Include(Again)?[ \t]*[, \t]?\s+(.*)$", o)
			{
				IsIncludeAgain := (o1 = "Again")
				IgnoreErrors := false
				IncludeFile := o2
				if RegExMatch(IncludeFile, "\*[iI]\s+?(.*)", o)
					IgnoreErrors := true, IncludeFile := Trim(o1)
				
				if RegExMatch(IncludeFile, "^<(.+)>$", o)
				{
					if IncFile2 := FindLibraryFile(o1, FirstScriptDir)
					{
						IncludeFile := IncFile2
						goto _skip_findfile
					}
				}
				
				StringReplace, IncludeFile, IncludeFile, `%A_ScriptDir`%, %FirstScriptDir%, All
				StringReplace, IncludeFile, IncludeFile, `%A_AppData`%, %A_AppData%, All
				StringReplace, IncludeFile, IncludeFile, `%A_AppDataCommon`%, %A_AppDataCommon%, All
				StringReplace, IncludeFile, IncludeFile, `%A_LineFile`%, %AhkScript%, All
				
				if InStr(FileExist(IncludeFile), "D")
				{
					SetWorkingDir, %IncludeFile%
					continue
				}
				
				_skip_findfile:
				
				IncludeFile := Util_GetFullPath(IncludeFile)
				
				AlreadyIncluded := false
				for k,v in FileList
					if (v = IncludeFile)
					{
						AlreadyIncluded := true
						break
					}
				if(IsIncludeAgain || !AlreadyIncluded)
				{
					if !AlreadyIncluded
						FileList.Insert(IncludeFile)
					PreprocessScript(ScriptText, IncludeFile, KeepComments, KeepIndent, KeepEmpties, FileList, FirstScriptDir, Options, IgnoreErrors)
				}
			}else if !contSection && ttline ~= "i)^FileInstall[, \t]"
			{
				if ttline ~= "^\w+\s+(:=|\+=|-=|\*=|/=|//=|\.=|\|=|&=|\^=|>>=|<<=)"
					continue ; This is an assignment!
				
				; workaround for `, detection
				EscapeChar := Options.esc
				EscapeCharChar := EscapeChar EscapeChar
				EscapeComma := EscapeChar ","
				EscapeTmp := chr(2)
				EscapeTmpD := chr(3)
				StringReplace, ttline, ttline, %EscapeCharChar%, %EscapeTmpD%, All
				StringReplace, ttline, ttline, %EscapeComma%, %EscapeTmp%, All
				
				if !RegExMatch(ttline, "i)^FileInstall[ \t]*[, \t][ \t]*([^,]+?)[ \t]*(,|$)", o) || o1 ~= "[^``]%"
					Util_Error("Error: Invalid ""FileInstall"" syntax found. Note that the first parameter must not be specified using a continuation section.")
				_ := Options.esc
				StringReplace, o1, o1, %_%`%, `%, All
				StringReplace, o1, o1, %_%`,, `,, All
				StringReplace, o1, o1, %_%%_%,, %_%,, All
				
				; workaround for `, detection [END]
				StringReplace, o1, o1, %EscapeTmp%, `,, All
				StringReplace, o1, o1, %EscapeTmpD%, %EscapeChar%, All
				StringReplace, ttline, ttline, %EscapeTmp%, %EscapeComma%, All
				StringReplace, ttline, ttline, %EscapeTmpD%, %EscapeCharChar%, All
				
				ScriptText .= (KeepIndent ? indent : "") (KeepComments ? tline : ttline) "`n"
			}else if !contSection && RegExMatch(tline, "i)^#CommentFlag\s+(.+)$", o)
				Options.comm := o1, ScriptText .= (KeepIndent ? indent : "") (KeepComments ? tline : ttline) "`n"
			else if !contSection && RegExMatch(tline, "i)^#EscapeChar\s+(.+)$", o)
				Options.esc := o1, ScriptText .= (KeepIndent ? indent : "") (KeepComments ? tline : ttline) "`n"
			else if !contSection && RegExMatch(tline, "i)^#DerefChar\s+(.+)$", o)
				Util_Error("Error: #DerefChar is not supported.")
			else if !contSection && RegExMatch(tline, "i)^#Delimiter\s+(.+)$", o)
				Util_Error("Error: #Delimiter is not supported.")
			else
				ScriptText .= (contSection ? A_LoopReadLine : (KeepIndent ? indent : "") (KeepComments ? tline : ttline)) "`n"
		}else{
			if KeepComments
				ScriptText .= A_LoopReadLine "`n"
			if StrStartsWith(tline, "*/")
				cmtBlock := false
		}
	}
	
	Loop, % !!IsFirstScript ; equivalent to "if IsFirstScript" except you can break from the block
	{
		static AhkPath := A_IsCompiled ? A_ScriptDir "\..\AutoHotkey.exe" : A_AhkPath
		IfNotExist, %AhkPath%
			break ; Don't bother with auto-includes because the file does not exist
		
		; Auto-including any functions called from a library...
		ilibfile = %A_Temp%\_ilib.ahk
		IfExist, %ilibfile%, FileDelete, %ilibfile%
			AhkType := AHKType(AhkPath)
		if AhkType = FAIL
			Util_Error("Error: The AutoHotkey build used for auto-inclusion of library functions is not recognized.", 1, AhkPath)
		if AhkType = Legacy
			Util_Error("Error: Legacy AutoHotkey versions (prior to v1.1) are not allowed as the build used for auto-inclusion of library functions.", 1, AhkPath)
		tmpErrorLog := Util_TempFile()
		RunWait, "%AhkPath%" /iLib "%ilibfile%" /ErrorStdOut "%AhkScript%" 2>"%tmpErrorLog%", %FirstScriptDir%, UseErrorLevel
		FileRead,tmpErrorData,%tmpErrorLog%
		FileDelete,%tmpErrorLog%
		if (ErrorLevel = 2)
			Util_Error("Error: The script contains syntax errors.",1,tmpErrorData)
		IfExist, %ilibfile%
		{
			PreprocessScript(ScriptText, ilibfile, KeepComments, KeepIndent, KeepEmpties, FileList, FirstScriptDir, Options)
			FileDelete, %ilibfile%
		}
		StringTrimRight, ScriptText, ScriptText, 1 ; remove trailing newline
	}
	
	if OldWorkingDir
		SetWorkingDir, %OldWorkingDir%
}

IsFakeCSOpening(tline)
{
	Loop, Parse, tline, %A_Space%%A_Tab%
		if !StrStartsWith(A_LoopField, "Join") && InStr(A_LoopField, ")")
			return true
	return false
}

FindLibraryFile(name, ScriptDir)
{
	libs := [ScriptDir "\Lib", A_MyDocuments "\AutoHotkey\Lib", A_ScriptDir "\..\Lib"]
	p := InStr(name, "_")
	if p
		name_lib := SubStr(name, 1, p-1)
	
	for each,lib in libs
	{
		file := lib "\" name ".ahk"
		IfExist, %file%
			return file
		
		if !p
			continue
		
		file := lib "\" name_lib ".ahk"
		IfExist, %file%
			return file
	}
}

StrStartsWith(ByRef v, ByRef w)
{
	return SubStr(v, 1, StrLen(w)) = w
}

RegExEscape(String)
{
	return "\Q" StrReplace(String, "\E", "\E\\E\Q") "\E"
}

Util_TempFile(d:="")
{
	if ( !StrLen(d) || !FileExist(d) )
		d:=A_Temp
	Loop
		tempName := d "\~temp" A_TickCount ".tmp"
	until !FileExist(tempName)
	return tempName
}

Util_GetFullPath(path)
{
	VarSetCapacity(fullpath, 260 * (!!A_IsUnicode + 1))
	if DllCall("GetFullPathName", "str", path, "uint", 260, "str", fullpath, "ptr", 0, "uint")
		return fullpath
	else
		return ""
}

Util_Error(txt, doexit=1, extra="")
{
	throw Exception(txt, -2, extra)
}

; Based on code from SciTEDebug.ahk
AHKType(exeName)
{
	FileGetVersion, vert, %exeName%
	if !vert
		return "FAIL"
	
	StringSplit, vert, vert, .
	vert := vert4 | (vert3 << 8) | (vert2 << 16) | (vert1 << 24)
	
	exeMachine := GetExeMachine(exeName)
	if !exeMachine
		return "FAIL"
	
	if (exeMachine != 0x014C) && (exeMachine != 0x8664)
		return "FAIL"
	
	if !(VersionInfoSize := DllCall("version\GetFileVersionInfoSize", "str", exeName, "uint*", null, "uint"))
		return "FAIL"
	
	VarSetCapacity(VersionInfo, VersionInfoSize)
	if !DllCall("version\GetFileVersionInfo", "str", exeName, "uint", 0, "uint", VersionInfoSize, "ptr", &VersionInfo)
		return "FAIL"
	
	if !DllCall("version\VerQueryValue", "ptr", &VersionInfo, "str", "\VarFileInfo\Translation", "ptr*", lpTranslate, "uint*", cbTranslate)
		return "FAIL"
	
	oldFmt := A_FormatInteger
	SetFormat, IntegerFast, H
	wLanguage := NumGet(lpTranslate+0, "UShort")
	wCodePage := NumGet(lpTranslate+2, "UShort")
	id := SubStr("0000" SubStr(wLanguage, 3), -3, 4) SubStr("0000" SubStr(wCodePage, 3), -3, 4)
	SetFormat, IntegerFast, %oldFmt%
	
	if !DllCall("version\VerQueryValue", "ptr", &VersionInfo, "str", "\StringFileInfo\" id "\ProductName", "ptr*", pField, "uint*", cbField)
		return "FAIL"
	
	; Check it is actually an AutoHotkey executable
	if !InStr(StrGet(pField, cbField), "AutoHotkey")
		return "FAIL"
	
	; We're dealing with a legacy version if it's prior to v1.1
	return vert >= 0x01010000 ? "Modern" : "Legacy"
}

GetExeMachine(exepath)
{
	if !(exe := FileOpen(exepath, "r"))
		return
	
	exe.Seek(60), exe.Seek(exe.ReadUInt()+4)
	return exe.ReadUShort()
}
class HelpFile
{
	static BaseURL := "ms-its:" A_AhkPath "\..\AutoHotkey.chm::/docs/"
	static Cache := {"Syntax": {}}
	
	GetPage(Path)
	{
		static xhttp := ComObjCreate("MSXML2.XMLHTTP.3.0")
		html := ComObjCreate("htmlfile")
		Path := this.BaseURL . RegExReplace(Path, "[?#].+")
		xhttp.open("GET", Path, True), xhttp.send()
		html.open(), html.write(xhttp.responseText), html.close()
		while !(html.readyState = "interactive" || html.readyState = "complete")
			Sleep, 50
		return html
	}
	
	GetLookup()
	{
		if this.Lookup
			return this.Lookup
		
		; Scrape the command reference
		this.Commands := {}
		try
			Page := this.GetPage("commands/index.htm")
		
		try ; Windows
			rows := Page.querySelectorAll(".info td:first-child a")
		catch ; Wine
			try
				rows := Page.body.querySelectorAll(".info td:first-child a")
		catch ; IE8
		{
			rows := new this.HTMLCollection()
			trows := Page.getElementsByTagName("table")[0].children[0].children
			loop, % trows.length
				rows.push(trows.Item(A_Index-1).children[0].children[0])
		}
		
		loop, % rows.length
			for i, text in StrSplit((row := rows.Item(A_Index-1)).innerText, "/")
				if RegExMatch(text, "^[\w#]+", Match) && !this.Commands.HasKey(Match)
					this.Commands[Match] := "commands/" RegExReplace(row.getAttribute("href"), "^about:")
		
		; Scrape the variables page
		this.Variables := {}
		try
			Page := this.GetPage("Variables.htm")
		
		try ; Windows
			rows := Page.querySelectorAll(".info td:first-child")
		catch ; Wine
			try
				rows := Page.body.querySelectorAll(".info td:first-child")
		catch ; IE8
		{
			rows := new this.HTMLCollection()
			tables := Page.getElementsByTagName("table")
			loop, % tables.length
			{
				trows := tables.Item(A_Index-1).children[0].children
				loop, % trows.length
					rows.push(trows.Item(A_Index-1).children[0])
			}
		}
		
		loop, % rows.length
			if RegExMatch((row := rows.Item(A_Index-1)).innerText, "(A_\w+)", Match)
				this.Variables[Match1] := "Variables.htm#" row.parentNode.getAttribute("id")
		
		; Combine
		this.Lookup := this.Commands.Clone()
		for k, v in this.Variables
			this.Lookup[k] := v
		
		return this.Lookup
	}
	
	Open(Keyword:="")
	{
		Lookup := this.GetLookup()
		Suffix := Lookup[Keyword] ? Lookup[Keyword] : "AutoHotkey.htm"
		Run, % "hh.exe """ this.BaseURL . Suffix """"
	}
	
	GetSyntax(Keyword:="")
	{
		; Generate this.Commands
		this.GetLookup()
		
		; Only look for Syntax of commands
		if !(Path := this.Commands[Keyword])
			return
		
		; Try to find it in the cache
		if this.Cache.Syntax.HasKey(Keyword)
			return this.Cache.Syntax[Keyword]
		
		; Get the right DOM to search
		Page := this.GetPage(Path)
		Root := Page ; Keep the page root in memory or it will be garbage collected
		if RegExMatch(Path, "#\K.+", ID)
			Page := Page.getElementById(ID)
		
		try ; Windows
			Nodes := page.getElementsByClassName("Syntax")
		catch ; Wine
			try
				Nodes := page.body.getElementsByClassName("Syntax")
		catch ; IE8
			Nodes := page.getElementsByTagName("pre")
		
		try ; Windows
			Text := Nodes.Item(0).innerText
		catch ; Some versions of Wine
			Text := Nodes.Item(0).innerHTML
		
		; Cache and return the result
		this.Cache.Syntax[Keyword] := StrSplit(Text, "`n", "`r")[1]
		return this.Cache.Syntax[Keyword]
	}
	
	class HTMLCollection
	{
		length[]
		{
			get
			{
				; Rounding MaxIndex produces a similar effect
				; to this.Length(), but doesn't trigger recursion
				return Round(this.MaxIndex())
			}
		}
		
		Item(i)
		{
			return this[i+1]
		}
	}
}