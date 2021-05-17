TrayTipAutoHide(Title,Text,Time:=2000,Options:=0){
TrayTip %Title%, %Text%,,Options
Sleep %Time%   ; Let it display for %time% in ms
HideTrayTip()
}

HideTrayTip() {
TrayTip  ; Attempt to hide it the normal way.
if SubStr(A_OSVersion,1,3) = "10." {
    Menu Tray, NoIcon
    Sleep 200  ; It may be necessary to adjust this sleep.
    Menu Tray, Icon
    }
}
