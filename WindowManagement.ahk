;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	WindowManagement (Funtion Library)
; DESCRIPTION:	Contains functions to more effectivally manage and move windows.
; VERSION:		2.9.3.26
; AUTHOR:		Noel Gordon (veggieman1996@gmail.com)


; ##########################################################
; 	COMPATABILITY AND DEPENDENCIES
; ##########################################################


#Requires AutoHotkey v2


; ##########################################################
; 	LIBRARY GLOBAL VARIABLES AND SETTINGS
; ##########################################################


; none


; ##########################################################
; MULTIPLE MONITOR WINDOW MANAGEMENT FUNCTIONS
; ##########################################################


; Get the number of the monitor that the given window is on (0 if minimized)
; Uses window mid-point to determine which monitor it is on
; Used active window if window title is not given
WindowManagement_GetWinMon(winTitle := "A")
{

	; Set title to active window if nesesary
	if (winTitle == "A"){
		winTitle := WinGetTitle("A")
	}

	; If window is minimized return 0
	winState := WinGetMinMax(winTitle)
	if (winState == -1){
		return 0
	}

	; Calculate windows min-point
	WinGetPos(&winX, &winY, &winW, &winH, winTitle)
	winMidX := winX + (winW/2)
	winMidY := winY + (winH/2)

	; Get number of monitors
	totalMonitors := MonitorGetCount()
	monitorNum := 1

	; Go through monitors and check if window is on that monitor
	while (monitorNum <= totalMonitors)
	{

		; Define bounding coordinates for monitors
		MonitorGetWorkArea(monitorNum, &MonCoordsLeft, &MonCoordsTop, &MonCoordsRight, &MonCoordsBottom)

		; Check if window mid-point is in monitor bound
		inBound_H := false
		inBound_V := false

		if (winMidX > MonCoordsLeft && winMidX < MonCoordsRight){
			inBound_H := true
		}
		if (winMidY > MonCoordsTop && winMidY < MonCoordsBottom){
			inBound_V := true
		}

		if (inBound_H && inBound_V){
			return monitorNum
		}

		monitorNum := monitorNum + 1
	}

	MsgBox("WindowManagement_GetWinMon: Window Monitor Not Found")

}


; Move the active window to a partucular monitor (but does not maximize)
WindowManagement_MoveToMon(monitorNum := 1)
{
	;make sure enough monitors are installed (if not default to highest monitor)
	totalMonitors := MonitorGetCount()
	if (monitorNum > totalMonitors)
	{
		monitorNum := totalMonitors
	}

	;get monitors operating system name and coordinates
	MonName := MonitorGetName(monitorNum)
	MonitorGetWorkArea(monitorNum, &MonCoordsLeft, &MonCoordsTop, &MonCoordsRight, &MonCoordsBottom)

	;define coordinates to move window to (before maximizing)
	winCoordX := MonCoordsLeft + 100
	winCoordY := MonCoordsTop + 100
	winWidth := abs(abs(MonCoordsRight) - abs(MonCoordsLeft)) - 200
	winHeight := abs(abs(MonCoordsBottom) - abs(MonCoordsTop)) - 200

	;move and maximize window
	winTitle := WinGetTitle("A")
	WinMove(winCoordX, winCoordY, winWidth, winHeight, winTitle)
	WinMove(winCoordX, winCoordY, winWidth, winHeight, winTitle)
}



; Maximize the active window on a partucular monitor
WindowManagement_MaxOnMon(monitorNum := 1)
{

	; Only maximize the window if not already maximized on that window
	winMon := WindowManagement_GetWinMon()
	activeTitle := WinGetTitle("A")
	winState := WinGetMinMax(activeTitle)
	if (winState != 1 || winMon != monitorNum){
		WindowManagement_MoveToMon(monitorNum)
		Sleep(100)
		Send("{LWin down}{Up}{LWin up}")
	}

}