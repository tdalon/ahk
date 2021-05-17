; See related blog post https://tdalon.blogspot.com/2021/02/ahk-create-gist.html
#Include <Jxon>
#SingleInstance force

Token := PowerTools_GetSetting("GitHubToken")
If (Token="")
	return

WinGetTitle, Title , A ;Get active window title
RegExMatch(Title,"(.*\\)?([^\\]*)\.([^\s]*)",filename) ;Try to isolate the file name from Window title with RegEx
filename = %filename2%.%filename3%
Code := GetSelection()
If (Code="") { ;Added errorLevel checking		
	MsgBox, No text selected.
	Return 
}

gui, add,text,,Name of file
gui, add,edit,w400 vFileName,%filename% 
gui, add,text,,Description
gui, add,edit,w400 vDescr r5 ;Type in description 
gui, Add, Button, default, Gist ;Create button
gui, show, autosize
return

GuiClose:
ExitApp
ButtonGist:
Gui, Submit  ; Save the input from the user to each control's associated variable.
Body:=Jxon_Dump({content:Code}) ;need to encode it 
Data={"description": "%Descr%","public": true,"files": {"%FileName%": %Body%}} ;build data for post having double quotes
ResponseText := Send(Token,"https://api.github.com/gists","POST",Data)
If (ResponseText="")
	return
Obj := Jxon_Load(ResponseText)
;Obj:=ParseJSON(ResponseText) ;make post and return information about it
sUrl := obj.html_url
Run, "%sUrl%"

;Clipboard:="Use the following for a webpage post:`n<script src='https://gist.github.com/" GitName "/"(SubStr(obj.url,Instr(obj.url,"/",,0)+1))".js'></script>`n`nYou can get the code here: " obj.html_url
return

Send(Token,URL:="",Verb:="",Data:=""){
	static WebRequest:=ComObjCreate("WinHttp.WinHttpRequest.5.1") ;Create COM object
	WebRequest.Open(Verb,URL) ;Open connection
	WebRequest.SetRequestHeader("Authorization","token " Token,0)
	WebRequest.SetRequestHeader("Content-Type","application/json")
	
	If (Login_IsVPN()) {
		ProxyServer := PowerTools_GetSetting("ProxyServer")
		If Not (ProxyServer="n/a") And (ProxyServer="") {
			MsgBox, 0x1011, HttpRequest with VPN?,It seems you are connected with VPN.`nHttpSend does not work via VPN if you use a Proxy. Consider disconnecting VPN.`nContinue now?
			IfMsgBox Cancel
				return
		}
		If Not (ProxyServer="n/a")
			WebRequest.SetProxy(2,ProxyServer) ; default proxy
	}

	WebRequest.Send(Data) ;Send Payload
	return WebRequest.ResponseText
}