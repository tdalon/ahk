; wrapper for ToolTip command
ToolTip(Message, TimeToDisplay = 500, SleepWhileDisplayed := False)
{
	; display message
	ToolTip,%Message%
	
	; clear tooltip after TimeToDisplay milliseconds
	SetTimer,ToolTipClear,-%TimeToDisplay%
	
	; sleep before returning
	If (SleepWhileDisplayed)
		Sleep,%TimeToDisplay%
        
	Return
	
	; clear tooltip
	ToolTipClear:
	ToolTip
	Return
}