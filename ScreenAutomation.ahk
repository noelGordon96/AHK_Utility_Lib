;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	ScreenAutomation (Function Library)
; DESCRIPTION:	Provides functions for easily auotmating mouse clicks
;				and movements to defined locations on the screen.
; VERSION:		2.5.20.26
; AUTHOR:		Noel Gordon (veggieman1996@gmail.com)
; SCOURCE:		none


; TODO:
; - Change locations from hardcoded to settings file
;   - Wait for user function


; FUNCTIONALITY NOTES (MOVE TO OTHER FILE LATER):

; - Anchor To Usage: If a location is created with an anchor location,
;   the coordinates will be recorded as relative to that anchor location.
;   Once a location is recorded as an anchored location, it cannot be
;   properly used as a non-anchored location. This is to prevent confusion
;   since the coordinates will be stored as relative to the anchor.

; - Auto Record Usage: Enabling this will make the script save a locations
;   coordinates automatically when the user uses/confirms clicks on the location.
;   This is useful of the location changes frequently and/or use as an anchor for
;   other locations.



; ##########################################################
; 	COMPATABILITY AND DEPENDENCIES
; ##########################################################


#Requires AutoHotkey v2
#Include "%A_LineFile%\..\UtilityWindows.ahk"


; ##########################################################
; 	LIBRARY GLOBAL VARIABLES AND SETTINGS
; ##########################################################


recordingMouseCoords := false ; flag to indicate if a function in this library is currently recording mouse coordinates
repairLocationName := "" ; name of the location being repaired (used to save coordinates)
repairCoord_screen_x := 0
repairCoord_screen_y := 0

; set DPI awareness context to "per monitor v2" for better scaling on high DPI displays
; this fixes issues with mouse coordinates being off on multiple displays with different DPI settings/scaling
DllCall("SetThreadDpiAwarenessContext", "ptr", -3, "ptr")

; globally set the coordinate mode for mouse functions to screen
CoordMode("Mouse", "Screen")


;###########################################################
;	PUBLIC MISC FUNCTIONS
;###########################################################


; Copy the current address from the browser address bar of a given window
ScreenAutomation_copyBrowserAddressBar(windowTitle := "A"){
	
	; activate the specified window (if not "A" for active)
	if (windowTitle != "A"){
		if WinExist(windowTitle){
			WinActivate(windowTitle)
			WinWaitActive(windowTitle, "", 5)
			if (!WinActive(windowTitle)){
				MsgBox("ERROR: The window `"" windowTitle "`" could not be activated! Exiting...")
				Exit()
			}
		}
		else{
			MsgBox("ERROR: The window `"" windowTitle "`" does not exist! Exiting...")
			Exit()
		}
	}
	
	; copy and return the address from the address bar
	Send("{F6}")
	Sleep(500)
	A_Clipboard := ""
	Send("{Ctrl down}c{Ctrl up}")
	ClipWait(1)
	Sleep(200)
	Send("{Escape}")
	Send("{Escape}")
	Sleep(200)
	tempCopiedText := A_Clipboard
	return tempCopiedText
}


;###########################################################
;	PUBLIC USER COORDINATION FUNCTIONS
;###########################################################


; Wait for user to press a predefined key to continue the script
; Key is defines in the settings for this library
ScreenAutomation_waitForUser(message := "Confirm Action", continueKey := "Enter", messageColor := "fc0303", pauseAfterConfirm := 0){
	activeWindow := WinGetTitle("A")
	msgWin := UtilityWindows_tranparentMessageWindow(message, "1500", "-150", messageColor)
	WinActivate(activeWindow) ; reactivate the previously active window (in case the message window stole focus)
	KeyWait(continueKey, "D")
	KeyWait(continueKey, "U")
	activeWindow := WinGetTitle("A")
	msgWin.Destroy()
	WinActivate(activeWindow) ; reactivate the previously active window (in case the message window stole focus)
	Sleep(pauseAfterConfirm)
}


;###########################################################
;	PUBLIC MOUSE AND SCREEN COORD FUCNTIONS
;###########################################################


; Retrieve a stored set of screen coordinates from the locations file
; If the coordinates do not exist, this function will move the process
; allowing the user to click the location and storing it for later
ScreenAutomation_mouseMoveToLocation(locationName, anchorLocation := ""){

	; attemp to get stored coordinates
	location := {}
	if (anchorLocation == ""){
		location := retrieveAbsoluteLocationCoords(locationName)
	}
	else{
		location := retrieveRelativeLocationCoords(locationName, anchorLocation)
	}
	
	coord_x := location.x
	coord_y := location.y

	; store send mode for later reset
	currentSendMode := A_SendMode
	
	; move the mouse to the specified coordinates
	SendMode("Event")
	BlockInput(true)
	MouseMove(coord_x, coord_y, 7)
	BlockInput(false)
	
	; reset the send mode to the previous value
	SendMode(currentSendMode)
}


; Save the current mouse location to the locations file
; This function can be used independently of the rest of the library
; but is primarily useful for use along with the "...waitForUser()" function
; This can then create ever changing locations that can be used as anchors
; -------------------------------------------
; TODO: Add option to save relative to an anchor location
; This would allow this fucntion to take the place of some of the other
; recording functions in this library
ScreenAutomation_saveCurrentMouseLocation(locationName){
	; get mouse screen coordinates
	currentMouseX := 0
	currentMouseY := 0
	MouseGetPos(&currentMouseX, &currentMouseY)
	; save the coordinates to the locations file
	ScreenAutomation_saveLocation(locationName, currentMouseX, currentMouseY)
}





; Retrieve a stored set of screen coordinates from the locations file
; If the coordinates do not exist, this function will move the process
; allowing the useer to click the location and storing it for later
ScreenAutomation_clickLocation(locationName, pauseAfterClick := 0, requireConfirmation := false, verifyWinTitle := "", anchorLocation := "", autoRecord := false){

	; move mouse to the specified location (record location if necesary)
	ScreenAutomation_mouseMoveToLocation(locationName, anchorLocation)
	
	; wait for user confirmation if required (autoRecord forced confirmation)
	if (requireConfirmation || autoRecord){
		ScreenAutomation_waitForUser()
	}

	; verify the window is active and under the mouse cursor
	if (verifyWinTitle != ""){
		verifyCursorWindow(verifyWinTitle)
	}

	; auto record location if enabled (not allowed with anchored locations)
	if (autoRecord){
		if (anchorLocation != ""){
			MsgBox("ERROR: ScreenAutomation_clickLocation(): AutoRecord cannot be used with an anchor location! Exiting...")
			Exit()
		}
		currentMouseX := 0
		currentMouseY := 0
		MouseGetPos(&currentMouseX, &currentMouseY)
		ScreenAutomation_saveLocation(locationName, currentMouseX, currentMouseY)
	}
	
	; store send mode for later reset
	currentSendMode := A_SendMode

	; move the mouse to the specified coordinates
	SendMode("Event")

	; continue and perform the click
	Sleep(100)
	Click()

	; reset the send mode to the previous value
	SendMode(currentSendMode)
	
	; optionally pause after the click
	if (pauseAfterClick > 0){
		Sleep(pauseAfterClick)
	}

}


; Delete a set of screen coordinates from the locations file
; If no named location is specified, all coordinates will be deleted (file is deleted)
ScreenAutomation_deleteCoordinates(locationName := ""){

	; connect to storage for screen coordinates
	mouseCoordsFile := A_ScriptDir "\screen_coords.ini"
	coordsSectionName := "Screen_Coords"

	; delete the specified location or all locations
	if (locationName != ""){
		IniDelete(mouseCoordsFile, coordsSectionName, locationName "_x")
		IniDelete(mouseCoordsFile, coordsSectionName, locationName "_y")
	}
	else{
		if (FileExist(mouseCoordsFile)){
			FileDelete(mouseCoordsFile)
		}
	}
}



;###########################################################
;	PUBLIC KEYBOARD AUTOMATION FUNCTIONS
;###########################################################


; Send a text string through keyboard automation
; This function is a wrapper for the Send command, allowing for
; easier integration with the rest of the script
ScreenAutomation_sendText(textToSend, pauseAfterSend := 0){

	; set key delay to slow down key presses
	currentKeyDelay := A_KeyDelay
	SetKeyDelay(20)

	; send the text string
	SendEvent(textToSend)

	;restore the key delay to the previous value
	SetKeyDelay(currentKeyDelay)

	; optionally add a delay after sending the text
	if (pauseAfterSend > 0){
		Sleep(pauseAfterSend)
	}
}


;###########################################################
;	PRIVATE UTILITY FUNCTIONS
;###########################################################


; Verify that the specified window is both active and under the mouse cursor
; Will first attempt to activate the window, then check if the mouse cursor is over it
; If the curser is not over the window, it will exit the thread... or...
; (if specified) simply return false
verifyCursorWindow(winTitle, exitThread := true){

	; attempt to activate the window
	if WinExist(winTitle){
		WinActivate(winTitle)
		WinWaitActive(winTitle, "", 5)
	}

	; determine window under the mouse cursor
	MouseGetPos(,, &curserWinID)
	curserWinTitle := WinGetTitle(curserWinID)


	; check if the window is active and the mouse cursor is over it
	if (WinActive(winTitle) && (curserWinTitle == winTitle)){
		return true
	}
	else{
		
		if (exitThread){
			winMessage := "ERROR: The window `"" winTitle "`" is not active or the mouse cursor is not over it.`r`n`r`nExiting thread..."
			MsgBox(winMessage, "Window Not Active", 0x10) ; 0x10 = Exclamation icon
			Exit()
		}
		
		return false
	}

}


; NOT USED YET
; Verify settings file exists otherwise create it with default values
/*
verifyLibrarySettings_SA(){
	
	; TODO STANDARDIZE APPROACH WITH SHORTCUT MANAGEMENT LIBRARY
	settingsFile := A_ScriptDir "\settings.ini"
	IniWrite("RShift", settingsFile, "Settings", "MouseCoordKey")
	IniWrite("0", settingsFile, "Settings", "MouseCoordDelay")
}
*/


; NOT USED YET
; take in a string of options and return true if the string contains a specific option
parseOption(targetOption, optionsString){
	; split the options string into an array
	optionsArray := StrSplit(optionsString, ",")
	
	; loop through the array and check if the target option is present
	for index, option in optionsArray{
		
		; split the option into key and value
		option := Trim(option)
		key := StrSplit(option, ":")[1]
		value := StrSplit(option, ":")[2]


		if (key == targetOption){
			return value
		}
	}
	
	; if the target option was not found, return false
	return false
}


;###########################################################
;	COORDINATE RETRIEVAL AND REPAIR FUNCTIONS
;###########################################################



; Save a set of screen coordinates to the locations file
; currently has no way of differentiating between anchored and non-anchored locations
; (screen coordinates vs relative coordinates)
ScreenAutomation_saveLocation(locationName, coord_x, coord_y){

	; connect to storage for screen coordinates
	mouseCoordsFile := A_ScriptDir "\screen_coords.ini"
	coordsSectionName := "Screen_Coords"
	coord_x_key := locationName "_x"
	coord_y_key := locationName "_y"

	; write the coordinates to the file
	IniWrite(coord_x, mouseCoordsFile, coordsSectionName, coord_x_key)
	IniWrite(coord_y, mouseCoordsFile, coordsSectionName, coord_y_key)
}


; Read a set of saves screen coordinates from the locations file
; if the coordinates no not exist this function will help fix them
retrieveAbsoluteLocationCoords(locationName){

	rawCoords := readLocationCoords_raw(locationName)
	if (rawCoords != "NOT_FOUND"){
		return rawCoords
	}
	else{
		recordStoredLocation(locationName)
		rawCoords := readLocationCoords_raw(locationName)
		if (rawCoords != "NOT_FOUND"){
			return rawCoords
		}
		else{
			MsgBox("ERROR: ScreenAutomation_saveLocation(): Coordinates were not repaired! Exiting...")
			Exit()
		}
	}

}


retrieveRelativeLocationCoords(locationName, anchorLocation){
	
	; get the screen coordinates for the anchor location
	anchorCoords := retrieveAbsoluteLocationCoords(anchorLocation)
	anchor_x := anchorCoords.x
	anchor_y := anchorCoords.y

	; get the stored relative coordinates for the specified location (repair as relative if necesary)
	locationCoords := readLocationCoords_raw(locationName)
	if (locationCoords == "NOT_FOUND"){
		recordStoredLocation(locationName, anchor_x, anchor_y)
		locationCoords := readLocationCoords_raw(locationName)
		if (locationCoords == "NOT_FOUND"){
			MsgBox("ERROR: ScreenAutomation: retrieveRelativeLocationCoords(): Coordinates were not repaired! Exiting...")
			Exit()
		}
	}

	; calculate the relative coordinates based on the anchor location and screen coordinates
	relativeCoords := {
		x: anchorCoords.x + locationCoords.x,
		y: anchorCoords.y + locationCoords.y
	}

	return relativeCoords
}


; Read locations stored corodinates (DOES NOT RECORD THEM IF THEY DO NOT EXIST)
readLocationCoords_raw(locationName){

	; connect to storage for screen coordinates
	mouseCoordsFile := A_ScriptDir "\screen_coords.ini"
	coordsSectionName := "Screen_Coords"
	coord_x_key := locationName "_x"
	coord_y_key := locationName "_y"

	; read the stored coordinates from the file and return then if found
	coord_x := IniRead(mouseCoordsFile, coordsSectionName, coord_x_key, "NOT_FOUND")
	coord_y := IniRead(mouseCoordsFile, coordsSectionName, coord_y_key, "NOT_FOUND")

	; if the coordinates were found, return them
	if (coord_x != "NOT_FOUND" && coord_y != "NOT_FOUND"){
		return {x: coord_x, y: coord_y}
	}
	else {
		return "NOT_FOUND"
	}
}


;###########################################################
;	COORDINATE RECORDING FUNCTIONS/HOTKEY
;###########################################################



; record a location by prompting the user to hover over it and press a predefined key
; the coordinates will then be saved to the locations file
; allows coridinates to be recorded relative to an anchor location
recordStoredLocation(locationName, relativeTo_x := 0, relativeTo_y := 0){
	
	winMessage := "The script could not find the coords defines for `"" locationName "`".`r`n`r`nTo record mouse location, hover your mouse over the location and press RSHIFT..."
	showCoordRepairWin("Mouse Location Not Found", winMessage, "Cancel")
	
	global repairLocationName := locationName ; pass the location name to the hotkey
	global recordingMouseCoords := true ; activates repair hotkey
	WinWaitClose("Mouse Location Not Found")

	; WAIT FOR HOTKEY TO RECORD THE COORDINATES AND CLOSE THE MESSAGE WINDOW

	global repairCoord_screen_x, repairCoord_screen_y ; get the recorded coordinates from the hotkey

	; translate the coordinates (if necessary) and save them
	repairCoord_screen_x := repairCoord_screen_x - relativeTo_x
	repairCoord_screen_y := repairCoord_screen_y - relativeTo_y
	ScreenAutomation_saveLocation(locationName, repairCoord_screen_x, repairCoord_screen_y)
}



; similar to a plain MsgBox except that it does not pause the script
showCoordRepairWin(winTitle, winMessage, btnText := "Cancel"){
    
	; create custom gui win
	global custMsgWin := Gui(, winTitle)
	custMsgWin.Opt("+AlwaysOnTop -MinimizeBox")
    custMsgWin.MarginX := 10
    custMsgWin.MarginY := 10
	custMsgWin.BackColor := "ffffff"
    custMsgWin.SetFont("s9", "Segoe UI")
    custMsgWin.Add("Text", "w400 +Wrap", winMessage)
    
    okBtn := custMsgWin.Add("Button", "Default w80 h25 x170", btnText)
    okBtn.OnEvent("Click", closeRepairWindow)

	custMsgWin.OnEvent("Close", closeRepairWindow)
    custMsgWin.Show("xCenter y150 w420")
	
	; enable PowerToys mouse finder (optional)
	/*
	Sleep 500
	Send("{LCtrl}")
	Sleep 200
	Send("{LCtrl}")
	*/

}



; close repair window and set the flag to false
closeRepairWindow(*){
	custMsgWin.Destroy()
	global recordingMouseCoords := false
}



; This hotkey is almost treated like a function
; Global variables are used to "pass" information in and out
; The hotkey will get the current mouse position store it (pass the info out)
#HotIf (recordingMouseCoords == true)

RSHIFT::{
	KeyWait("RShift", "U")
	
	; debug boolean
	debugMsg := false

	global repairLocationName ; variable to hold "parameter" value
	global repairCoord_screen_x, repairCoord_screen_y ; variables to hold "return" values

	; get and save the current mouse position
	currentMouseX := 0
	currentMouseY := 0
	
	; get coordinates from the screen
	MouseGetPos(&currentMouseX, &currentMouseY)
	repairCoord_screen_x := currentMouseX
	repairCoord_screen_y := currentMouseY

	; show a message box with the recorded coordinates (for debugging)
	if (debugMsg){
		MsgBox("Recorded Coordinates:`r`nX: " currentMouseX "`r`nY: " currentMouseY)
	}

	; display message to user
	closeRepairWindow()
	global recordingMouseCoords := false
	repairLocationName := ""

}

#HotIf

;###########################################################