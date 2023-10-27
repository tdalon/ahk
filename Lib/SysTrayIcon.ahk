; wrapper for ToolTip command
SysTrayIcon(IcoFile, TimeToDisplay := 1000)
{
	; display IcoFile
	Menu,Tray,Icon, %IcoFile%
	; restore icon after TimeToDisplay milliseconds
	SetTimer,IcoRestore,-%TimeToDisplay%        
	Return
	
	; clear ico
	IcoRestore:
    RestoreIcoFile  := PathX(A_ScriptFullPath, "Ext:.ico").Full
	If (FileExist(RestoreIcoFile)) 
        Menu,Tray,Icon, %RestoreIcoFile%
	Return
}