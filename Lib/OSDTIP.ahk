OSDTIP_Alert(P*) {                        ; OSDTIP_Alert v0.54 by SKAN on D37P/D383 @ tiny.cc/osdtip
Local
Static FN:="", ID:=0, PS:="", PM:="", P8:=(A_PtrSize=8 ? "Ptr" : "")
  If !IsObject(FN)
    FN := Func(A_ThisFunc).Bind(A_ThisFunc) 

  If (P.Count()=0 || P[1]==A_ThisFunc) {
    If (P[4]=0x201) ;            WM_NCLBUTTONDOWN=0xA1, HTCAPTION=2       ; WM_LBUTTONDOWN=0x201
    Return DllCall("SendMessage", "Ptr",ID, "Int",0xA1,"Ptr",2, "Ptr",0)  ;   
    OnMessage(0x201, FN, 0),  OnMessage(0x010, FN, 0)                     ; WM_LBUTTONDOWN, WM_CLOSE 
    SetTimer, %FN%, OFF
    Progress, 6:OFF                       
    Return ID:=0                                                              
  }                                         

  MT:=P[1], ST:=P[2], OP := P[4] . A_Space, TMR:=P[3], FONT:=P[5] ? P[5] : "Segoe UI",  
  TRN :=Round(P[6]) ? P[6] & 255 : 255, Title := (TMR=0 ? "0x0" : A_ScriptHwnd) . ":" . A_ThisFunc
  OP.= InStr(OP,"V1") ? "CWFFFFE2 CT856442 CBEBB800" : InStr(OP,"V2") ? "CWF0F8FF CT1A4482 CB3399FF" 
    :  InStr(OP,"V3") ? "CWF0FFE9 CT155724 CB429300" : InStr(OP,"V4") ? "CWFFEEED CT721C24 CBE40000" 
    :  InStr(OP,"V0") ? "CW3F3F3F CTDADADA CB797979" : ""
  PBG := (F := InStr(OP,"CB",1)) ? SubStr(OP, F+2, 6) : "797979"
  PBG := Format("0x{5:}{6:}{3:}{4:}{1:}{2:}", StrSplit(PBG)*)

  WinClose, ahk_id %ID%
  DetectHiddenWindows, % ("On", DHW:=A_DetectHiddenWindows)
  SetWinDelay, % (-1, SWD:=A_WinDelay)  
  SetControlDelay, % (0, SCD:=A_WinDelay)

  DllCall("uxtheme\SetThemeAppProperties", "Int",0)
  Progress, 6: ZX6 ZY4 ZH16 FS10 FM11 WS400 WM800 C00 CT222222 %OP% B1 M Hide
          , %ST%, %MT%, %Title%, %FONT%
  DllCall("uxtheme\SetThemeAppProperties", "Int",7)
  WinWait, %Title% ahk_class AutoHotkey2
  ControlGetPos,,,,         PBS, msctls_progress321
  ControlGetPos, X1,,,, Static1
  ControlGetPos, X2,,,, Static2
  NM := X1+Round(PBS//2)
  Progress, 6: ZY4 ZH16 FS10 FM11 WS400 WM800 C00 CT222222 CB797979 %OP% ZX%NM% B1 M Hide
          , %ST%, %MT%, %Title%, %FONT%
  WinWait, %Title% ahk_class AutoHotkey2          

  WinSet, Transparent, %TRN%, % "ahk_id" . (ID:=WinExist())
  WinGetPos, WX, WY, WW, WH
  ControlGetPos,,,,         PBS, msctls_progress321
  ControlGetPos,, Y1, W1, H1, Static1
  ControlGetPos,, Y2, W2, H2, Static2  
  WH := Y1 + H1 + Round(H2) + 2

  SysGet, M, MonitorWorkArea, % Round(P[9])
  mX := mLeft, mY := mTop, mW := mRight-mLeft, mH := mBottom-mTop 
  WX := mX + ( P[7]="" ? (mW//2)-(WW//2) : P[7]<0 ? mW-WW+P[7]+1 : P[7] )
  WY := mY + ( P[8]="" ? (mH//2)-(WH//2) : P[8]<0 ? mH-WH+P[8]+1 : P[8] )    
  WinMove,,, % WX , % WY , % WW, % WH

  ControlMove, Static1, % X1+PBS, % Y1,      % W1, % H1
  ControlMove, Static2, % X2+PBS, % Y1+H1+2, % W2, % H2
  Control, ExStyle, -0x20000,    msctls_progress321                      ; WS_EX_STATICEDGE, removed
  SendMessage, 0x2001, 1, % PBG, msctls_progress321                      ; PBM_SETBACKCOLOR
  ControlMove, msctls_progress321, 0, 0, % PBS, % WH  

  SetControlDelay, %SCD%
  SetWinDelay, %SWD%
  DetectHiddenWindows, %DHW%
  SC := DllCall("GetClassLong" . P8, "Ptr",ID, "Int",-26, "UInt")        ; GCL_STYLE
  DllCall("SetClassLong" . P8, "Ptr",ID, "Int",-26, "Ptr",SC|0x20000)    ; GCL_STYLE, CS_DROPSHADOW    
  Progress, 6:SHOW                                                     
  DllCall("SetClassLong" . P8, "Ptr",ID, "Int",-26, "Ptr",SC)            ; GCL_STYLE

  If (Round(TMR)<0)
    SetTimer, %FN%, %TMR%
  OnMessage(0x202, FN, TMR=0 ? 0 : -1),  OnMessage(0x010, FN)            ; WM_LBUTTONUP,  WM_CLOSE
Return ID := WinExist()
}

;- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

OSDTIP_Desktop(P*) {                    ; OSDTIP_Desktop v0.50 by SKAN on D35P/D36E @ tiny.cc/osdtip
Local
Static FN:="", ID:=0, PS:="", PM:="", P8:=(A_PtrSize=8 ? "Ptr" : "")

  If !IsObject(FN)
    FN := Func(A_ThisFunc).Bind(A_ThisFunc) 

  If (P.Count()=0 || P[1]==A_ThisFunc) {
    If (P[4]=0x201) ;            WM_NCLBUTTONDOWN=0xA1, HTCAPTION=2      ; WM_LBUTTONDOWN=0x201
    Return DllCall("SendMessage", "Ptr",ID, "Int",0xA1,"Ptr",2, "Ptr",0) ;   
    OnMessage(0x201, FN, 0),  OnMessage(0x010, FN, 0)                    ; WM_LBUTTONDOWN, WM_CLOSE 
    SetTimer, %FN%, OFF
    Progress, 7:OFF                       
    Return ID:=0                                                              
  }
 
  MT:=P[1], ST:=P[2], TMR:=P[3], OP:=P[4], FONT:=P[5] ? P[5] : "Segoe UI"
  TRN:=P[6] ? P[6] : "A0A0A0 127", Title := (TMR=0 ? "0x0" : A_ScriptHwnd) . ":" . A_ThisFunc
  
  If (ID) {                           
    Progress, 7:, % (ST=PS ? "" : PS:=ST), % (MT=PM ? "" : PM:=MT), %Title%        
    SetTimer, %FN%, % Round(TMR)<0 ? TMR : "OFF"
    OnMessage(0x201, FN, TMR=0 ? 0 : -1)                                 ; WM_LBUTTONDOWN 
    Return ID
  }                                                                                                        

  DetectHiddenWindows, % ("Off", DHW:=A_DetectHiddenWindows)
  If !hSDV:=DllCall("GetWindow", "Ptr",WinExist("ahk_class Progman"), "UInt",5, "Ptr")  ; GW_CHILD=5
      hSDV:=DllCall("GetWindow", "Ptr",WinExist("ahk_class WorkerW"), "UInt",5, "Ptr")  ; GW_CHILD=5
  DetectHiddenWindows, On     
  SetWinDelay, % (-1, SWD:=A_WinDelay)

  DllCall("uxtheme\SetThemeAppProperties", "Int",0)
  Progress, 7: ZX0 ZY0 ZH1 w200 FS14 FM28 CWA0A0A0 CTFEFEFE B %OP% M HIDE
          , %ST%, %MT%, %Title%, %FONT%
  DllCall("uxtheme\SetThemeAppProperties", "Int",7)
  WinWait %Title% ahk_class AutoHotkey2

  Control, Style,   0x50000000,  msctls_progress321                      ; WS_VISIBLE | WS_CHILD
  Control, ExStyle,-0x20000,     msctls_progress321                      ; WS_EX_STATICEDGE 
  If !InStr(OP,"U4") {
    Control, Style,  0x50000002, Static1                                 ; WS_VISIBLE | WS_CHILD
    Control, Style,  0x50000002, Static2                                 ; | SS_RIGHT
    }
  SendMessage, 0x2001, 0, P[9]!="" ? P[9] : 0xFFFFFF, msctls_progress321 ; PBM_SETBACKCOLOR=0x2001
  WinSet, TransColor, %TRN%
  WinGetPos, X, Y, W, H
  SysGet, M, MonitorWorkArea
  If !InStr(OP,"U5") 
    X:=MRight-W-14, Y:=MBottom-H-14
  Else 
    X := P[7]="" ? (MRight/2) -(W/2) : P[7]<0 ? MRight -W+P[7] : P[7]
  , Y := P[8]="" ? (MBottom/2)-(H/2) : P[8]<0 ? MBottom-H+P[8] : P[8]    
  ID:=WinExist()               ; SetWindowPos HWND_BOTTOM=1, SWP_SHOWWINDOW=0x40 SWP_NOACTIVATE=0x10
  DllCall("SetWindowPos", "Ptr",ID, "Ptr",1, "Int",X, "Int",Y, "Int",W+2, "Int",H, "UInt",0x40|0x10)
  DllCall("SetWindowPos", "Ptr",ID, "Ptr",1, "Int",X, "Int",Y, "Int",W+0, "Int",H, "UInt",0x40|0x10)
  DllCall("SetWindowLong" . P8, "Ptr",ID, "Int",-8, "Ptr",hSDV)          ; GWL_HWNDPARENT
  SetWinDelay, %SWD%
  DetectHiddenWindows, %DHW%
  Progress, 7:SHOW  
  If (Round(TMR)<0) 
    SetTimer, %FN%, %TMR%
  OnMessage(0x201, FN, TMR=0 ? 0 : -1),  OnMessage(0x010, FN)            ; WM_LBUTTONDOWN,  WM_CLOSE
Return ID
}

;- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

OSDTIP_Volume(P*) {                      ; OSDTIP_Volume v0.50 by SKAN on D35P/D369 @ tiny.cc/osdtip
Local
Static FN:="", ID:=0, PV:=0, P8:=(A_PtrSize=8 ? "Ptr" : "") 

  If !IsObject(FN)
    FN := Func(A_ThisFunc).Bind(A_ThisFunc) 

  If (P.Count()=0 || P[1]==A_ThisFunc) {
    OnMessage(0x202, FN, 0),  OnMessage(0x010, FN, 0)                   ; WM_LBUTTONUP, WM_CLOSE 
    SetTimer, %FN%, OFF
    Progress, 8:OFF                       
    Return ID:=0                                                              
  }
  
  M:=P[1], V:=P[2], VSigned:=InStr("+-",SubStr(V,1,1)), TMR:=P[3]
  OP:=P[4], FONT:=P[5] ? P[5] : "Trebuchet MS",  TRN:=Round(P[6]) ? P[6] & 255 : 222
  Title := (TMR=0 ? "0x0" : A_ScriptHwnd) . ":" . A_ThisFunc
  
  If (M!="") 
    SoundSet, %M%,, MUTE
  SoundGet, M,, MUTE
  If ( V!="" && !VSigned)
    SoundSet, %V%  
  SoundGet, VOL
  VOL:=Round(VOL)

  If WinExist("ahk_id" . ID) 
    {
      If (V && VSigned)
        SoundSet, % VOL:=(VOL:=V ? Round((VOL+V)/V)*V : VOL)>100 ? 100 : VOL<0 ? 0 : Round(VOL)
      SendMessage, 0x0409, 1, % (M="On" ? 0x0030FF:0x00FFAA), msctls_progress321 ; PBM_SETBARCOLOR
      SendMessage, 0x2001, 0, % (M="On" ? 0x00175A:0x00402E), msctls_progress321 ; PBM_SETBACKCOLOR
      Progress, 8:%VOL%, % PV!=VOL ? PV:=VOL : "",, %Title% 
      SetTimer, %FN%, % Round(TMR)<0 ? TMR : "OFF"
      Return ID
    }  

  DetectHiddenWindows, % ("On", DHW:=A_DetectHiddenWindows)
  SetWinDelay, % (-1, SWD:=A_WinDelay)  
  DllCall("uxtheme\SetThemeAppProperties", "Int",0)
  Progress, 8:C11 w318 ZH24 ZX28 ZY4 WM400 WS600 FM16 FS22 CT111111 CWF0F0F0 %OP% B1 HIDE
          , % PV:=VOL, V O L U M E, %Title%, %FONT%
  DllCall("uxtheme\SetThemeAppProperties", "Int",7)
  WinWait, %Title% ahk_class AutoHotkey2

  WinSet, Transparent, %TRN%, % "ahk_id" . (ID:=WinExist())
  SendMessage, 0x0409, 1, % (M="On" ? 0x0030FF:0x00FFAA), msctls_progress321 ; PBM_SETBARCOLOR
  SendMessage, 0x2001, 0, % (M="On" ? 0x00175A:0x00402E), msctls_progress321 ; PBM_SETBACKCOLOR
  Control, ExStyle, -0x20000, msctls_progress321
  DetectHiddenWindows, %DHW%
  Progress, 8:%VOL% 
  SC := DllCall("GetClassLong" . P8, "Ptr",ID, "Int",-26, "UInt")       ; GCL_STYLE
  DllCall("SetClassLong" . P8, "Ptr",ID, "Int",-26, "Ptr",SC|0x20000)   ; GCL_STYLE, CS_DROPSHADOW    
  Progress, 8:SHOW
  DllCall("SetClassLong" . P8, "Ptr",ID, "Int",-26, "Ptr",SC)           ; GCL_STYLE
  If (Round(TMR)<0) 
    SetTimer, %FN%, %TMR%
  OnMessage(0x202, FN, TMR=0 ? 0 : -1),  OnMessage(0x010, FN)           ; WM_LBUTTONUP,  WM_CLOSE
Return ID := WinExist()
}

;- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

OSDTIP_KBLeds(P*) {                      ; OSDTIP_KBLeds v0.50 by SKAN on D361/D367 @ tiny.cc/osdtip 
Local
Static FN:="", ID:=0 

  If !IsObject(FN)
    FN := Func(A_ThisFunc).Bind(A_ThisFunc) 

  If (P.Count()=0 || P[1]==A_ThisFunc) {
    OnMessage(0x202, FN, 0),  OnMessage(0x010, FN, 0)                   ; WM_LBUTTONUP, WM_CLOSE 
    SetTimer, %FN%, OFF
    Progress, 9:OFF                       
    Return ID:=0                                                              
  }

  Key := P[1], ST:=P[2], TMR:=P[3], OP:=P[4], FONT:=P[5] ? P[5] : "Trebuchet MS"
  Title := (TMR=0 ? "0x0" : A_ScriptHwnd) . ":" . A_ThisFunc, TRN:=Round(P[6]) ? P[6] & 255 : 222
     
  If WinExist("ahk_id" . ID) {                                          
    ST.=InStr(ST,"off") || InStr(ST,"on") ? "" :  GetKeyState(Key,"T") ? "Off" : "On"
    Switch (Key) {
      Case "CapsLock"   : SetCapsLockState,   %ST%
      Case "ScrollLock" : SetScrollLockState, %ST%
      Case "NumLock"    : SetNumLockState,    %ST%
    }
    C:=GetKeyState("CapsLock","T"), S:=GetKeyState("ScrollLock","T"), N:=GetKeyState("NumLock","T")
    SendMessage, 0x2001, 1,% C ? 0x00FFAA:0x808080, msctls_progress321 ; PBM_SETBACKCOLOR
    SendMessage, 0x2001, 1,% S ? 0x00AAFF:0x808080, msctls_progress322 ; PBM_SETBACKCOLOR  
    SendMessage, 0x2001, 1,% N ? 0x00FFAA:0x808080, msctls_progress323 ; PBM_SETBACKCOLOR

    If (Key="CapsLock" && C=1) || (Key="NumLock" && N=0)
    If ( InStr(OP,"U2",1) && FileExist(WAV:=A_WinDir . "\Media\Windows Default.wav") )
      DllCall("winmm\PlaySoundW", "WStr",WAV, "Ptr",0, "Int",0x220013)  ; SND_FILENAME | SND_ASYNC   

    SetTimer, %FN%, % Round(TMR)<0 ? TMR : "OFF" 
    Progress, 9:,,,%Title%
    Return ID
  }
                                                                                            
  DetectHiddenWindows, % ("On", DHW:=A_DetectHiddenWindows)
  SetWinDelay, % (-1, SWD:=A_WinDelay)
  SetControlDelay, % (0, SCD:=A_ControlDelay)                  
  DllCall("uxtheme\SetThemeAppProperties", "Int",0)
  Progress, 9:ZX32 ZY6 ZH32 W172 WM600 WS400 FM16 FS16 CT101010 CWF0F0F0 %OP% C00 B1 HIDE
          , ScrollLock, CapsLock, %Title%, %FONT%
  WinWait %Title% ahk_class AutoHotkey2                                  

  WinGetPos, WX, WY, WW, WH, % "ahk_id" . (ID:=WinExist())
  Loop, Parse, % "msctls_progress32|msctls_progress32|Static", | 
  DllCall("CreateWindowEx", "Int",0, "Str",A_LoopField, "Str","NumLock" ; WS_VISIBLE | WS_CHILD
       ,"Int",0x50000000, "Int",0, "Int",0, "Int",10, "Int",10, "Ptr",ID, "Ptr",0, "Ptr",0, "Ptr",0)
  DllCall("uxtheme\SetThemeAppProperties", "Int",7)                     
  SendMessage, 0x31, 0, 0,            Static1                           ; WM_GETFONT
  SendMessage, 0x30, %ErrorLevel%, 1, Static3                           ; WM_SETFONT
  ControlGetPos, CX, CY, CW, CH, Static1
  YM:=CY-1, NX:=CX+CH+24, WW:=WW+CH+24, WH:=(CH*3)+(YM*4)+2, PH:=Round(CH/2), PY:=CY+(PH/2) 
  WX:=(A_ScreenWidth/2)-(WW/2), WY := (A_ScreenHeight/2)-(WH/2)
  WinMove,% "ahk_id" WinExist(),,% WX,% WY,% WW, % WH
  ControlMove, Static1,            % NX, % CY,             % CW, % CH
  ControlMove, Static2,            % NX, % CY+CH+YM,       % CW, % CH
  ControlMove, Static3,            % NX, % CY+CH+YM+CH+YM, % CW, % CH
  ControlMove, msctls_progress321, % CX, % PY,             % CH, % PH
  ControlMove, msctls_progress322, % CX, % PY+CH+YM,       % CH, % PH
  ControlMove, msctls_progress323, % CX, % PY+CH+YM+CH+YM, % CH, % PH
  Loop 3
  Control, Style, +0x202, Static%A_Index%                               ; SS_RIGHT | SS_CENTERIMAGE
  WinSet, Transparent, %TRN%
  SetControlDelay, %SCD%
  SetWinDelay, %SWD%
  DetectHiddenWindows, %DHW%

  P8 := (A_PtrSize=8 ? "Ptr":"")
  SC := DllCall("GetClassLong" . P8, "Ptr",ID, "Int",-26, "UInt")       ; GCL_STYLE
  DllCall("SetClassLong" . P8, "Ptr",ID, "Int",-26, "Ptr",SC|0x20000)   ; GCL_STYLE, CS_DROPSHADOW    
  Progress, 9:SHOW
  DllCall("SetClassLong" . P8, "Ptr",ID, "Int",-26, "Ptr",SC)           ; GCL_STYLE

  P[3]:=0, n:=%A_ThisFunc%(P*)
  If (Round(TMR)<0) 
    SetTimer, %FN%, %TMR%
  OnMessage(0x202, FN, TMR=0 ? 0 : -1),  OnMessage(0x010, FN)           ; WM_LBUTTONUP,  WM_CLOSE
Return ID  
}

;- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

OSDTIP_Pop(P*) {                            ; OSDTIP_Pop v0.55 by SKAN on D361/D36E @ tiny.cc/osdtip 
Local
Static FN:="", ID:=0, PM:="", PS:="" 

  If !IsObject(FN)
    FN := Func(A_ThisFunc).Bind(A_ThisFunc) 

  If (P.Count()=0 || P[1]==A_ThisFunc) {
    OnMessage(0x202, FN, 0),  OnMessage(0x010, FN, 0)                   ; WM_LBUTTONUP, WM_CLOSE 
    SetTimer, %FN%, OFF
    DllCall("AnimateWindow", "Ptr",ID, "Int",200, "Int",0x50004)        ; AW_VER_POSITIVE | AW_SLIDE
    Progress, 10:OFF                                                    ; | AW_HIDE
    Return ID:=0
  }

  MT:=P[1], ST:=P[2], TMR:=P[3], OP:=P[4], FONT:=P[5] ? P[5] : "Segoe UI"
  Title := (TMR=0 ? "0x0" : A_ScriptHwnd) . ":" . A_ThisFunc
  
  If (ID) {
    Progress, 10:, % (ST=PS ? "" : PS:=ST), % (MT=PM ? "" : PM:=MT), %Title%
    OnMessage(0x202, FN, TMR=0 ? 0 : -1)                                ; v0.55
    SetTimer, %FN%, % Round(TMR)<0 ? TMR : "OFF" 
    Return ID
  }                                                                                                        

  If ( InStr(OP,"U2",1) && FileExist(WAV:=A_WinDir . "\Media\Windows Notify.wav") )
    DllCall("winmm\PlaySoundW", "WStr",WAV, "Ptr",0, "Int",0x220013)    ; SND_FILENAME | SND_ASYNC   
                                                                        ; | SND_NODEFAULT   
  DetectHiddenWindows, % ("On", DHW:=A_DetectHiddenWindows)             ; | SND_NOSTOP | SND_SYSTEM  
  SetWinDelay, % (-1, SWD:=A_WinDelay)                            
  DllCall("uxtheme\SetThemeAppProperties", "Int",0)
  Progress, 10:C00 ZH1 FM9 FS10 CWF0F0F0 CT101010 %OP% B1 M HIDE,% PS:=ST, % PM:=MT, %Title%, %FONT%
  DllCall("uxtheme\SetThemeAppProperties", "Int",7)                     ; STAP_ALLOW_NONCLIENT
                                                                        ; | STAP_ALLOW_CONTROLS
  WinWait, %Title% ahk_class AutoHotkey2                                ; | STAP_ALLOW_WEBCONTENT
  WinGetPos, X, Y, W, H                                                 
  SysGet, M, MonitorWorkArea
  WinMove,% "ahk_id" . WinExist(),,% MRight-W,% MBottom-(H:=InStr(OP,"U1",1) ? H : Max(H,100)), W, H
  If ( TRN:=Round(P[6]) & 255 )
    WinSet, Transparent, %TRN% 
  ControlGetPos,,,,H, msctls_progress321       
  If (H>2) {
    ColorMQ:=Round(P[7]),  ColorBG:=P[8]!="" ? Round(P[8]) : 0xF0F0F0,  SpeedMQ:=Round(P[9])
    Control, ExStyle, -0x20000,        msctls_progress321               ; v0.55 WS_EX_STATICEDGE
    Control, Style, +0x8,              msctls_progress321               ; PBS_MARQUEE
    SendMessage, 0x040A, 1, %SpeedMQ%, msctls_progress321               ; PBM_SETMARQUEE
    SendMessage, 0x0409, 1, %ColorMQ%, msctls_progress321               ; PBM_SETBARCOLOR
    SendMessage, 0x2001, 1, %ColorBG%, msctls_progress321               ; PBM_SETBACKCOLOR
  }  
  DllCall("AnimateWindow", "Ptr",WinExist(), "Int",200, "Int",0x40008)  ; AW_VER_NEGATIVE | AW_SLIDE
  SetWinDelay, %SWD%
  DetectHiddenWindows, %DHW%
  If (Round(TMR)<0)
    SetTimer, %FN%, %TMR%
  OnMessage(0x202, FN, TMR=0 ? 0 : -1),  OnMessage(0x010, FN)           ; WM_LBUTTONUP,  WM_CLOSE
Return ID:=WinExist()
}

;- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

OSDTIP(hWnd:="") {
Local OSDTIP 
  If (hWnd="")
     Return A_ScriptHwnd . ":OSDTIP_" . "ahk_class AutoHotkey2"
  If !WinExist("ahk_id" . hWnd)  
     Return 0  
  WinGetTitle, OSDTIP
  OSDTIP := StrSplit(OSDTIP,":")
  If ( OSDTIP[1] = A_ScriptHwnd ) 
       OSDTIP[2]()
       
}

;- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
