;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	Shortcut Management (Function Library)
; DESCRIPTION:	Provides functions easily running and managing a script's shortcuts
; VERSION:		2.4.9.25
; AUTHOR:		Noel Gordon (veggieman1996@gmail.com)
; SCOURCE:		none


; ##########################################################
; USAGE AND OTHER INFO
; ##########################################################


; REQUIRED FILES AND SETTINGS

; "shortcutDir" global variable: this variable should denote the full path
; of the hotkey projects shortcut folder, most often this should be defined
; in the following manner near the begining of the main script:

; shortcutDir := A_ScriptDir "\shortcut_files"



;###########################################################
;	PUBLIC SHORTCUT RUNNING FUCNTIONS
;###########################################################



ShortcutManagement_getStoredFolderPath(locationName){

	; connect to storage for folder paths
	global shortcutDir

	pathLocationsFile := shortcutDir "\folder_locations.ini"
	pathSectionName := "Folder_Paths"
	
	; read folder path from storage and return the path name if location exists
	folderPath := IniRead(pathLocationsFile, pathSectionName, locationName, "NOT_FOUND")
	if InStr(FileExist(folderPath), "D"){
		return folderPath
	}
	
	; help the user repair the location if directory does not exist
	
	; if a value was found in the folder storage file (show old value to user to help with repair)
	if (folderPath != "NOT_FOUND"){
		winMessage := "`r`nNOTE: The location specified was found containing the following old path...`r`n`r`n" . folderPath . "`r`n"
		showMessageWindow("Old Target Location", winMessage)
	}
	
	; Open the main dialong to get user decision on corrective action
	winMessage := "The script could not find the path specified for `"" locationName "`".`r`n`r`nPath: " folderPath "`r`n`r`nWould you like repair this location by searching for it?"
	userChoice := Msgbox(winMessage, "Path Not Found", 49)
	
	; Get the target path from user (or exit sub-routine)
	targetPath := ""
	if (userChoice == "OK"){
		targetPath := DirSelect(,, "Select the folder location to repair")
	}
	else if (userChoice == "Cancel"){
		Exit
	}
	
	; write the new folder location to the path storage file
	if (targetPath != ""){
		verifySortcutDir()
		IniWrite(targetPath, pathLocationsFile, pathSectionName, locationName)
		return targetPath
	}
	
	; exit sub-routine if user did not select a folder
	else{
		Exit
	}
}



; Run a shortcut from the shortcut folder
; If a shortcut fails to open (which probably means the shortcut is broken)
; This function will help the user in fixing the shortcut
ShortcutManagement_runShortcut(shortcutName){
	
	; connect to storage for folder paths
	global shortcutDir
	
	;parse shortcut full path (minus the extension)
	shortcutName := "_Shortcut_" shortcutName
	shortcutPath_noExt := shortcutDir "\" shortcutName
	shortcutPath := shortcutPath_noExt ".lnk"

	; check if shortcut file is an lnk instead or a url
	; TEMP SOLUTION: should probably create a "openURL" function instead...
	; ...this would allow opening in new window vs tab (but would require browser specific code)
	; ...also this method does not allow automatic repair/creation of URL shortcuts
	if FileExist(shortcutPath_noExt ".url"){
		shortcutPath := shortcutPath_noExt ".url"
	}
	
	
	; attemp to run the shortcut
	try {
		Run(shortcutPath)
	}

	; if run was unsuccessfull try to create shortcut if necesary
	catch {
		
		; verify shortcut directory exists and create it if it does not
		verifySortcutDir()

		; Check if the shortcut file already existed to assist the user in reparing it
		if (FileExist(shortcutPath)){
			shortcutTarget := ""
			FileGetShortcut(shortcutPath, &shortcutTarget)
			winMessage := "`r`nNOTE: The shortcut specified was found containing the following target...`r`n`r`n" shortcutTarget "`r`n"
			showMessageWindow("Old Shortcut Target", winMessage)
		}
		
		
		; Open the main dialong to get user decision on corrective action
		winMessage := "The script could not open the specified shortcut `"" shortcutName ".lnk`". We can try to repair the shortcut right now...`r`n`r`nIs the shortcut a file?`r`n`r`nIf the shortcut is a folder, press `"No`". Otherwise press `"Cancel`" to manually repair the shortcut or ignore for now."
		userChoice := MsgBox(winMessage, "Shortcut Broken", 51)
		
		; Get the target path from user (or exit)
		if (userChoice == "Yes"){
			targetPath := FileSelect(,,"Select Shortcut Target")
		}
		else if (userChoice == "No"){
			targetPath := DirSelect(,, "Select target folder...")
		}
		else if (userChoice == "Cancel"){
			Exit
		}

		; exit sub-routine if user did not select a file or folder
		if (targetPath == ""){
			MsgBox("ShortcutManagement: No target selected. Exiting...")
			Exit()
		}
		
		; Parse the target file/folder's parent directory
		; (this will serve as the shortcuts working directory)
		targetPathArray := StrSplit(targetPath, "\")
		targetPathArray.Pop()	;remove last element
		arrayLen := targetPathArray.Length
		
		parentDirPath := targetPathArray[1]
		counter := 2
		while (counter <= arrayLen)
		{
			parentDirPath := parentDirPath "\" targetPathArray[counter]
			counter := counter + 1
		}
		
		; Create the new shortcut file
		FileCreateShortcut(targetPath, shortcutPath, parentDirPath)
		
		; Exit the subroutine or hotkey to avoid further issues
		; User will need to press hotkey again to continue workflow
		Exit()
	}
}



;###########################################################
;	PRIVATE UTILITY FUNCTIONS
;###########################################################



; Varify "shortcutDir" variable is defined for error checking
; Also create the required directory if it does not exist
verifySortcutDir(){
	global shortcutDir

	varSet := IsSet(shortcutDir)
	if (varSet){
		if InStr(FileExist(shortcutDir), "D"){
			return
		}
		else{
			DirCreate shortcutDir
		}
	}
	else{
		msg := "ERROR: ShortcutManagement_verifyShortcutDir(): shortcutDir variable is not set. This variable should be defined in the main script."
		MsgBox(msg)
		Exit()
	}
}



;###########################################################
;	PRIVATE CUSTOM GUI MANAGEMENT FUNCTIONS
;###########################################################


; similar to a plain MsgBox except that it does not pause the script
showMessageWindow(winTitle, winMessage, btnText := "OK"){
    
	; create custom gui win
	oldTargetWin := Gui(, winTitle)
	oldTargetWin.Opt("+AlwaysOnTop -MinimizeBox")
    oldTargetWin.MarginX := 10
    oldTargetWin.MarginY := 10
	oldTargetWin.BackColor := "ffffff"
    oldTargetWin.SetFont("s9", "Segoe UI")
    oldTargetWin.Add("Text", "w400 +Wrap", winMessage)
    
    okBtn := oldTargetWin.Add("Button", "Default w80 h25 x170", btnText)
    okBtn.OnEvent("Click", closeMessageWindow)

	oldTargetWin.OnEvent("Close", closeMessageWindow)
    oldTargetWin.Show("xCenter y150 w420")

	; gui internal funtions
	closeMessageWindow(*){
		oldTargetWin.Destroy()
	}
}

