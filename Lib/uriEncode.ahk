uriEncode(str)
{
	; Replace characters with uri encoded version except for letters, numbers,
	; and the following: /.~:&=-

	f = %A_FormatInteger%
	SetFormat, Integer, Hex
	pos = 1
	Loop
		If pos := RegExMatch(str, "i)[^\/\w\.~`:%&=-]", char, pos++)
			StringReplace, str, str, %char%, % "%" . Asc(char), All
		Else Break
	SetFormat, Integer, %f%
	StringReplace, str, str, 0x, , All
	Return, str
}