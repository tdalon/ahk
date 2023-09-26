; Called by IntelliHtml-> Link2Text -> ConnectionsGetForumTitle, CreateTicket
; sResponseText := HttpGet(sUrl, sPassword*,sUserName*)
; If password is passed as argument, authentification is sent with username/password
; Example: 
; sPassword = Login_GetPassword() - uses password stored within PowerTools
; sResponse := Hpptget(sUrl, sPassword) - uses default username environment variable
HttpGet(sUrl,sPassword :="", sUserName:=""){
WebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
WebRequest.Open("GET", sUrl, false) ; Async=false
If !(sPassword = "") { ; not empty
    If (sUserName = "")
        sUserName := A_UserName	
    ;MsgBox %sUserName%: %sPassword% ; DBG
    WebRequest.SetCredentials(sUserName, sPassword, 0)
    WebRequest.SetCredentials(sUserName, sPassword, 1)
}

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

WebRequest.Send()
sResponse := WebRequest.ResponseText
; Debug
sStatusText := WebRequest.StatusText
;MsgBox %sStatusText% ; DBG
If !(sStatusText = "OK"){
    MsgBox 0x10, Error, Error on WinHttpRequest - GET: %sStatusText%
    ;return
}

sSource := WebRequest.ResponseText
return sSource
}


URLDownloadToVar(URL){
	http:=ComObjCreate("WinHttp.WinHttpRequest.5.1")
	http.Open("GET",URL,1)
	http.SetRequestHeader("Pragma","no-cache")
	http.SetRequestHeader("Cache-Control","no-cache")
	http.Send(),http.WaitForResponse
	return (http.Status=200?http.ResponseText:"Error")
}