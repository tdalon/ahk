; https://stackoverflow.com/a/34357644/2043349
Monitor_WinMove(hWin:="") {

If (hWin="")
    hWin := WinActive("A")
if hWin = 0
    return
; https://autohotkey.com/board/topic/32874-moving-the-active-window-from-one-monitor-to-the-other/

SysGet, Mon2, Monitor, 2
SysGet, Mon1, Monitor, 1

Mon1Width  := Mon1Right - Mon1Left
Mon1Height := Mon1Top - Mon1High
Mon2Width  := Mon2Right - Mon2Left
Mon2Height := Mon2Top - Mon2High

;MsgBox, Left: %Mon1Left% -- Top: %Mon1Top% -- Right: %Mon1Right% -- Bottom %Mon1Bottom%. Mon1Width %Mon1Width%


WinGet, minMax, MinMax, ahk_id %hWin%
if (minMax = 1)
    WinRestore, ahk_id %hWin%
WinGetPos, x, y, width, height, ahk_id %hWin%

MonIndex := Monitor_GetMonitorIndex(hWin)

if (MonIndex = 1) { ; move from 1 to 2
    xScale := Mon2Width / Mon1Width
    yScale := Mon2Height / Mon1Height
    newX := x * xScale
    newX :=  newX + Mon1Width
;MsgBox %Mon1Width% %newX% ; DBG

} else { ; from 2 to 1
    xScale := Mon1Width / Mon2Width
    yScale := Mon1Height / Mon2Height
    newX := x * xScale
    newX := newX - Mon1Width
}
newY := y * yScale
newWidth := width * xScale
newHeight := height * yScale
;WinActivate, ahk_id %hWin% ; required to move
WinMove, ahk_id %hWin%, , %newX%, %newY%, %newWidth%, %newHeight%
if (minMax = 1)
    WinMaximize, ahk_id %hWin%
} ; eofun

; -----------------------------------------------------------------
Monitor_GetMonitorIndex(hWin:="") {
; https://autohotkey.com/board/topic/69464-how-to-determine-a-window-is-in-which-monitor/
If (hWin="")
    hWin := WinActive("A")
if hWin = 0
    return

; Starts with 1.
monitorIndex := 1

VarSetCapacity(monitorInfo, 40)
NumPut(40, monitorInfo)

if (monitorHandle := DllCall("MonitorFromWindow", "uint", hWin, "uint", 0x2)) 
    && DllCall("GetMonitorInfo", "uint", monitorHandle, "uint", &monitorInfo) 
{
    monitorLeft   := NumGet(monitorInfo,  4, "Int")
    monitorTop    := NumGet(monitorInfo,  8, "Int")
    monitorRight  := NumGet(monitorInfo, 12, "Int")
    monitorBottom := NumGet(monitorInfo, 16, "Int")
    workLeft      := NumGet(monitorInfo, 20, "Int")
    workTop       := NumGet(monitorInfo, 24, "Int")
    workRight     := NumGet(monitorInfo, 28, "Int")
    workBottom    := NumGet(monitorInfo, 32, "Int")
    isPrimary     := NumGet(monitorInfo, 36, "Int") & 1

    SysGet, monitorCount, MonitorCount

    Loop, %monitorCount%
    {
        SysGet, tempMon, Monitor, %A_Index%

        ; Compare location to determine the monitor index.
        if ((monitorLeft = tempMonLeft) and (monitorTop = tempMonTop)
            and (monitorRight = tempMonRight) and (monitorBottom = tempMonBottom))
        {
            monitorIndex := A_Index
            break
        }
    }
}
return monitorIndex
} ; eofun