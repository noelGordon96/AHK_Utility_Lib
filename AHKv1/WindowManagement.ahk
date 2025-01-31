;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	WindowManagement (Funtion Library)
; DESCRIPTION:	Contains functions used to manage and move windows on multiple montiors.
; VERSION:		1.11.8.24
; AUTHOR:		Noel Gordon (veggieman1996@gmail.com)


; ##########################################
; OTHER INFO
; ##########################################



; ##########################################
; WINDOW DATA RETRIEVAL FUNCTIONS
; ##########################################



; Get the number of the monitor that the given window is on (0 if minimized)
; Uses window mid-point to determine which monitor it is on
; Used active window if window title is not given
WindowManagement_GetWinMon(winTitle:="A")
{

	; Set title to active window if nesesary
	if (winTitle == "A"){
		WinGetActiveTitle, winTitle
	}

	; If window is minimized return 0
	WinGet, winState, MinMax, %winTitle%
	if (winState==0){
		return 0
	}

	; Calculate windows min-point
	WinGetPos, winX, winY, winW, winH, %winTitle%
	winMidX := winX + (winW/2)
	winMidY := winY + (winH/2)
	
	; Get number of monitors
	SysGet, totalMonitors, 80
	monitorNum := 1
	
	; Go through monitors and check if window is on that monitor
	while (monitorNum<=totalMonitors)
	{
	
		; Define bounding coordinates for monitors
		SysGet, MonCoords, MonitorWorkArea, %monitorNum%

		; Check if window mid-point is in monitor bound
		inBound_H := false
		inBound_V := false
		
		if (winMidX>MonCoordsLeft && winMidX<MonCoordsRight){
			inBound_H := true
		}
		if (winMidY>MonCoordsTop && winMidY<MonCoordsBottom){
			inBound_V := true
		}
		
		if (inBound_H && inBound_V){
			return monitorNum
		}
		
		monitorNum := monitorNum + 1
	}
	
	MsgBox, WindowManagement_GetWinMon: Window Monitor Not Found

}


; ##########################################
; OTHER FUNCTIONS
; ##########################################


; Move the active window to a partucular monitor (but does not maximize)
WindowManagement_MoveToMon(monitorNum=1)
{
	;make sure enough monitors are installed (if not default to highest monitor)
	SysGet, totalMonitors, 80
	if (monitorNum > totalMonitors)
	{
		monitorNum := totalMonitors
	}
	
	;get monitors operating system name and coordinates
	SysGet, MonName, MonitorName, %monitorNum%
	SysGet, MonCoords, MonitorWorkArea, %monitorNum%
	
	;define coordinates to move window to (before maximizing)
	winCoordX := MonCoordsLeft + 100
	winCoordY := MonCoordsTop + 100
	winWidth := abs(abs(MonCoordsRight) - abs(MonCoordsLeft)) - 200
	winHeight := abs(abs(MonCoordsBottom) - abs(MonCoordsTop)) - 200
	
	;move and maximize window
	WinGetActiveTitle, winTitle
	WinMove, %winTitle%,, %winCoordX%, %winCoordY% , %winWidth%, %winHeight%
	WinMove, %winTitle%,, %winCoordX%, %winCoordY% , %winWidth%, %winHeight%
}



; Maximize the active window on a partucular monitor
WindowManagement_MaxOnMon(monitorNum=1)
{

	; Only maximize the window if not already maximized on that window
	winMon := WindowManagement_GetWinMon()
	WinGetActiveTitle, activeTitle
	WinGet, winState, MinMax, %activeTitle%
	if (winState!=1 || winMon!=monitorNum){
		WindowManagement_MoveToMon(monitorNum)
		Sleep, 100
		Send {LWin down}{Up}{LWin up}
	}
	
}

