;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	Debug (Function Library)
; DESCRIPTION:	Provides functions for navigating and managing
;				windows virtual desktops.
; VERSION:		1.3.8.24
; AUTHOR:		Noel Gordon (veggieman1996@gmail.com) using external scource
; SCOURCE:		https://www.computerhope.com/tips/tip224.htm


; ##########################################################
; USAGE AND OTHER INFO
; ##########################################################


; NOTE: Registry keys used in this library are valid for Windows 11

; NOTE: Windows has a built in hotkey to open the Microsoft Office365 App
; Office365 Hotkey: CTRL + SHIFT + ALT + WIN
; The hotkeys used to manage Vitrual Desktops often use CTRL + WIN
; This means if SHIFT + ALT are in a triggering hotkey, this will cause all
; four keys to be down and trigger the Office365 hotkey
; To mitigate this, KeyWait is used below to ensure the ALT is released
; before attempting to send the keystrokes to manage the virtual desktops
; SHIFT can also be used for this purpose


;###########################################################
;	WINDOW MOVEMENT FUNCTIONS
;###########################################################
; These fuctions may be less reliable depending on speed of
; system, updates to Windows menues, etc.




; Move the active window to another virtual desktop
; Destination Virtional Desktop will be created if needed
ManageDesktops_moveWindowToVirtualDesktop(destinationDesktop)
{
	; Pull values from registry to determine current system state
	currentId := getCurrentDesktopId()
	desktopIdList := getVirtualDesktopIdList()
	
	desktopCount := getVirtualDesktopCount(currentId, desktopIdList)
	currentDesktop := getCurrentDesktopNumber(currentId, desktopIdList)
	
	; make sure the destination desktop is not the current desktop
	if (destinationDesktop != currentDesktop){
		
		
		; Toggle the window to be visible on all virtual desktops
		SetTitleMatchMode, 1
		WinGetTitle, winTitle, A
		WinSet, ExStyle, ^0x80, %winTitle%
		
		
		; Create virtual desktops if needed
		while (destinationDesktop > desktopCount){
			ManageDesktops_createVirtualDesktop()
			desktopCount := ManageDesktops_getVirtualDesktopCount()
		}
		
		; Move to destination virtual desktop and sleep to allow system to catch up
		ManageDesktops_switchDesktopByNumber(destinationDesktop)
		; sleep only needed if new desktops were not created (may optimize later)
		sleep % 400 * (abs(destinationDesktop - currentDesktop))
		
		; Toggle the window back to only visible on one virtual desktop
		WinSet, ExStyle, ^0x80, %winTitle%
		WinActivate, %winTitle%
		
	}
}






; Move multiple windows to another virtual desktop
; Windows are passed in as an array of window titles
; Destination Virtional Desktop will be created if needed
ManageDesktops_moveWindowsToVirtualDesktop(destinationDesktop, windowTitleArray)
{

; Pull values from registry to determine current system state
	currentId := getCurrentDesktopId()
	desktopIdList := getVirtualDesktopIdList()
	
	desktopCount := getVirtualDesktopCount(currentId, desktopIdList)
	currentDesktop := getCurrentDesktopNumber(currentId, desktopIdList)
	
	; make sure the destination desktop is not the current desktop
	if (destinationDesktop != currentDesktop){
		
		
		
		; Loop through list of windows and toggle to be visible on all virtual desktops
		SetTitleMatchMode, 1
		For index in windowTitleArray {
			currentValue := windowTitleArray[index]
			WinActivate, %currentValue%
			WinWaitActive, %currentValue%,, 5
			if WinActive(currentValue){
				WinSet, ExStyle, ^0x80, %currentValue%
			}
		}
		
		
		; Create virtual desktops if needed
		while (destinationDesktop > desktopCount){
			ManageDesktops_createVirtualDesktop()
			desktopCount := ManageDesktops_getVirtualDesktopCount()
		}
		
		; Move to destination virtual desktop and sleep to allow system to catch up
		ManageDesktops_switchDesktopByNumber(destinationDesktop)
		; sleep only needed if new desktops were not created (may optimize later)
		sleep % 400 * (abs(destinationDesktop - currentDesktop))
		
		
		; Loop back through list of windows and toggle back to only visible on one virtual desktop
		For index in windowTitleArray {
			currentValue := windowTitleArray[index]
			WinActivate, %currentValue%
			WinWaitActive, %currentValue%,, 5
			if WinActive(currentValue){
				WinSet, ExStyle, ^0x80, %currentValue%
			}
		}
		
	}

}








;###########################################################
;	VIRTUAL DESKTOP MANAGEMENT FUNCTIONS
;###########################################################



; This function switches to the desktop number provided.
; Destiniation Desktop needs to first exist
;
ManageDesktops_switchDesktopByNumber(targetDesktop)
{
	; Pull values from registry to determine current system state
	currentId := getCurrentDesktopId()
	desktopIdList := getVirtualDesktopIdList()
	
	; Map variables for later reference
	currentVirtualDesktop := getCurrentDesktopNumber(currentId, desktopIdList)
	virtualDesktopCount := getVirtualDesktopCount(currentId, desktopIdList)
	
	; Don't attempt to switch to an invalid desktop
	if (targetDesktop > virtualDesktopCount || targetDesktop < 1) {
		return
	}
	
	; Go right until we reach the desktop we want
	while(currentVirtualDesktop < targetDesktop) {
		KeyWait, ALT	;see note above
		;KeyWait, SHIFT
		Send ^#{Right}
		sleep, 500
		currentVirtualDesktop++
	}
	
	; Go left until we reach the desktop we want
	while(currentVirtualDesktop > targetDesktop) {
		KeyWait, ALT	;see note above
		;KeyWait, SHIFT
		Send ^#{Left}
		sleep, 500
		currentVirtualDesktop--
	}
}




; This function creates a new virtual desktop and switches to it
;
ManageDesktops_createVirtualDesktop()
{
	KeyWait, ALT	;see note above
	;KeyWait, SHIFT
	Send, #^d
	sleep, 500
}




; This function deletes the current virtual desktop
;
ManageDesktops_deleteVirtualDesktop()
{
	; send keystrokes to delete current virtual desktop
	; NOTE: KeyWait ensures that the built in Windows Office365
	; hotkey is not triggered: WIN+CRTL+ALT+SHIFT
	; either ALT or SHIFT can be used for this purpose
	KeyWait, ALT
	;KeyWait, SHIFT
	Send, #^{F4}
}




;###########################################################
;	VIRTUAL DESKTOP DATA RETRIEVING FUNCTIONS
;###########################################################





ManageDesktops_getVirtualDesktopName(desktopId := "current"){

	; Get current desktop id if not provided
	if (desktopId == "current"){
		desktopId := getCurrentDesktopId()
	}
	
	; Convert desktop id to uppercase
	StringUpper, destopId, desktopId
	
	; parse id into usable form to access window name
	; NOTE: Windows registry key to get desktop name is the same data
	; as the desktop id however some of the bytes are rearranged
	byte1 := SubStr(desktopId, 1, 2)
	byte2 := SubStr(desktopId, 3, 2)
	byte3 := SubStr(desktopId, 5, 2)
	byte4 := SubStr(desktopId, 7, 2)
	byte5 := SubStr(desktopId, 9, 2)
	byte6 := SubStr(desktopId, 11, 2)
	byte7 := SubStr(desktopId, 13, 2)
	byte8 := SubStr(desktopId, 15, 2)
	
	byte9_10 := SubStr(desktopId, 17, 4)
	byte11_16 := SubStr(desktopId, 21)
	
	; Place bytes in correct order with dividers for later lookup
	desktopId_parsed := byte4 . byte3 . byte2 . byte1 . "-" . byte6 . byte5 . "-" . byte8 . byte7 . "-" . byte9_10 . "-" . byte11_16
	
	; Windows registry location for the currently active virtual desktop name
	regKeyName_currentVirtualDesktopName := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops\Desktops\{" . desktopId_parsed . "}"
	regValName_currentVirtualDesktopName := "Name"
	
	; Get current desktop ID from registry and set IdLength
	RegRead, CurrentVirtualDesktopName, %regKeyName_currentVirtualDesktopName%, %regValName_currentVirtualDesktopName%
	
	; Return desktop name from the registry
	if (CurrentVirtualDesktopName) {
		return CurrentVirtualDesktopName
	}
	; Return generic name given by the system if needed (not stored in the registry)
	else {
		desktopIdList := getVirtualDesktopIdList()
		desktopNumber := getCurrentDesktopNumber(desktopId, desktopIdList)
		desktopName := "Desktop " . desktopNumber
		return desktopName
	}
}


; Return the number of virtual desktops currently open
ManageDesktops_getVirtualDesktopCount(){

	; Pull values from registry to determine current system state
	currentId := getCurrentDesktopId()
	desktopIdList := getVirtualDesktopIdList()
	
	; Return virtual desktop count for this system
	return getVirtualDesktopCount(currentId, desktopIdList)
}


; Return the number (1 indexed) of the currently active desktop
ManageDesktops_getCurrentDesktopNumber(){
	
	; Pull values from registry to determine current system state
	currentId := getCurrentDesktopId()
	desktopIdList := getVirtualDesktopIdList()
	
	; Return the current desktop number
	return getCurrentDesktopNumber(currentId, desktopIdList)
}





;###########################################################
;	UTILITY FUNCTIONS (PRIVATE, USED WITHIN OTHER FUNCTIONS)
;###########################################################






; Get the registry id of current virtual desktop
getCurrentDesktopId(){
	
	; Windows registry location for the currently active virtual desktop id
	regKeyName_currentVirtualDesktopId := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops"
	regValName_currentVirtualDesktopId := "CurrentVirtualDesktop"
	
	; Get current desktop ID from registry
	RegRead, CurrentVirtualDesktopId, %regKeyName_currentVirtualDesktopId%, %regValName_currentVirtualDesktopId%
	
	; Return desktop id if valid or display error
	if (CurrentVirtualDesktopId) {
		return CurrentVirtualDesktopId
	}
	else {
		MsgBox, ERROR Autohotkey ManageDesktops library: Could not find current virtutal desktop id in registry
	}
}


getVirtualDesktopIdList(){

	; Windows registry location for the currently active virtual desktop id
	regKeyName_virtualDesktopIdList := "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops"
	regValName_virtualDesktopIdList := "VirtualDesktopIDs"

	; Get a list of the UUIDs for all virtual desktops on the system
	RegRead, VirtualDesktopList, %regKeyName_virtualDesktopIdList%, %regValName_virtualDesktopIdList%
	
	; Return desktop id list if valid or display error
	if (VirtualDesktopList) {
		return VirtualDesktopList
	}
	else {
		MsgBox, ERROR Autohotkey ManageDesktops library: Could not find virtutal desktop id list in registry
	}
}




getCurrentDesktopNumber(currentId, desktopIdList){

	; Get desktop count
	desktopCount := getVirtualDesktopCount(currentId, desktopIdList)
	currentId_len := StrLen(currentId)
	
	; Parse the REG_DATA string that stores the array of UUID's for virtual desktops in the registry.
	i := 0
	while (i < desktopCount) {
		startPos := (i * currentId_len) + 1
		searchString := SubStr(desktopIdList, startPos, currentId_len)
		; check if search string matches current id
		if (searchString == currentId) {
			currentDesktopNumber := i + 1
			return currentDesktopNumber
		}
		i++
	}
	
	; Display error message if dektop was not found
	MsgBox, ERROR Autohotkey ManageDesktops library: Could not find current virtutal desktop in desktop registry list
}




; Calculate and return number of virtual desktops based of given ids
getVirtualDesktopCount(currentId, desktopIdList){
	; get character length of id and list
	curId_len := StrLen(currentId)
	idList_len := StrLen(desktopIdList)
	; calculate and return desktop count based on lengths
	desktopCount := idList_len / curId_len
	return desktopCount
}


