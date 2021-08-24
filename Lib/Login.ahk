Login_GetPassword(){
static sPassword
If !(sPassword = "")
return sPassword
InputBox, sPassword, Password, Enter Password for Login, Hide, 200, 125
If ErrorLevel
return
return sPassword
}
Login_SetPassword(){
InputBox, sPassword, Password, Enter Password for Login, Hide, 200, 125
If ErrorLevel
return
sKey := Login_GetPasswordKey()
sPassword := Encrypt(sPassword,sKey)
return sPassword
}
Login_GetPasswordKey(){
RegRead, sPasswordKey, HKEY_CURRENT_USER\Software\PowerTools, PasswordKey
If (sPasswordKey=""){
InputBox, sPasswordKey, Password Key, Enter Password Key for Password encryption, Hide, 200, 125
If ErrorLevel
return
}
RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, PasswordKey, %sPasswordKey%
return sPasswordKey
}
Login_GetPersonalNumber(){
RegRead, sPersNum, HKEY_CURRENT_USER\Software\PowerTools, PersonalNumber
If (sPersNum=""){
sPersNum := Login_SetPersonalNumber()
If (sPersNum="")
return
}
return sPersNum
}
Login_SetPersonalNumber(){
If GetKeyState("Ctrl") {
sUrl := "https://connectionsroot/wikis/home/wiki/Wc4f94c47297c_42c8_878f_525fd907cb68/page/Password%20Setting"
Run, "%sUrl%"
return
}
InputBox, sPersNum, Personal Number, Enter your Personal Number e.g. for KSSE,, 200, 125
If ErrorLevel
return
RegWrite, REG_SZ, HKEY_CURRENT_USER\Software\PowerTools, PersonalNumber, %sPersNum%
return sPersNum
}
Login_IsVPN(){
return Not (A_IPAddress2 = "0.0.0.0")
}
Login_VPNConnect(doWait:=False){
LoginTitle = Cisco AnyConnect |
SetTitleMatchMode, 1
hWnd := WinExist(LoginTitle)
If (hWnd <> 0)  {
WinActivate, ahk_id %hWnd%
GoTo InputVPNPassword
}
SetTitleMatchMode, 3
MainTitle = Cisco AnyConnect Secure Mobility Client
hWnd := WinExist(MainTitle)
If (hWnd <> 0) {
WinActivate, ahk_id %hWnd%
} Else {
VPNexePath := "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"
Run, %VPNexePath%
WinWaitActive, %MainTitle%
}
SetTitleMatchMode, 3
Sleep 500
ControlClick, Connect
If ErrorLevel
return
SetTitleMatchMode, 1
WinWaitActive, %LoginTitle%
InputVPNPassword:
sPassword := Login_GetPassword()
If (sPassword="")
return
ControlSetText,Edit2,%sPassword%
Send,{Enter}
If (doWait=True)
WinWaitClose, %MainTitle%
}
IsConnectedToInternet()
{
IsConnected := DllCall("Wininet.dll\InternetGetConnectedState", "Str", "0x40","Int",0)
return (IsConnected = 0)
}
Login_IsNet(NetName){
sTmpFile = %A_Temp%\ipconfig.txt
If FileExist(sTmpFile)
FileDelete, %sTmpFile%
RunWait, %ComSpec% /c "ipconfig >"%sTmpFile%"",,Hide
FileRead, Output, %sTmpFile%
If InStr(Output,NetName)
return True
Else
return False
}
