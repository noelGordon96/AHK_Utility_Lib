;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	Shortcut Management (Function Library)
; DESCRIPTION:	Provides functions easily running and managing a scripts shortcuts
; VERSION:		1.11.8.24
; AUTHOR:		Noel Gordon (veggieman1996@gmail.com)
; SCOURCE:		none


; ##########################################################
; USAGE AND OTHER INFO
; ##########################################################


; REQUIRED FILES AND SETTINGS

; "shortcutDir" global variable: this variable should denote the full path
; of the hotkey projects shortcut folder, most often this should be defined
; in the following manner near the begining of the main script:

; shortcutDir := A_ScriptDir . "\shortcut_files"



;###########################################################
;	PUBLIC SHORTCUT RUNNING FUCNTIONS
;###########################################################



ShortcutManagement_getStoredFolderPath(locationName){

	; connect to storage for folder paths
	global shortcutDir

	pathLocationsFile := shortcutDir . "\folder_locations.ini"
	pathSectionName := "Folder_Paths"
	
	
	; read folder path from storage and return the path name if location exists
	IniRead, folderPath, %pathLocationsFile%, %pathSectionName%, %locationName%, NOT_FOUND
	if InStr(FileExist(folderPath), "D"){
		return folderPath
	}
	
	
	; help the user repair the location if directory does not exist
	
	; if a value was found in the folder storage file (show old value to user to help with repair)
	if (folderPath != "NOT_FOUND"){
		winMessage := "`nNOTE: The location specified was found containing the following old path...`n`n" . folderPath . "`n"
		showMessageWindow(winMessage)
	}
	
	; Open the main dialong to get user decision on corrective action
	winMessage := "The script could not find the path specified for """ . locationName . """.`n`nPath: " . folderPath . "`n`nWould you like repair this location by searching for it?"
	Msgbox, 49, Path Not Found, %winMessage%
	
	; Get the target path from user (or exit sub-routine)
	IfMsgBox, OK
		FileSelectFolder, targetPath ,,, Select the folder location to repair...
	IfMsgBox, Cancel
		Exit
	
	
	; write the new folder location to the path storage file
	if (ErrorLevel == 0){
		IniWrite, %targetPath%, %pathLocationsFile%, %pathSectionName%, %locationName%
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

	global shortcutDir
	
	;parse shortcut full path
	shortcutName := "_Shortcut_" . shortcutName
	shortcutPath := shortcutDir . "\" . shortcutName . ".lnk"
	
	;attemp to run the shortcut
	Run, %shortcutPath%,, UseErrorLevel

	;check if run was successfull and try to create shortcut if necesary
	if (ErrorLevel == "ERROR"){
		
		
		; Check if the shortcut file already existed to assist the user in reparing it
		if FileExist(shortcutPath){
			FileGetShortcut, %shortcutPath%, shortcutTarget
			winMessage := "`nNOTE: The shortcut specified was found containing the following target...`n`n" . shortcutTarget . "`n"
			showMessageWindow(winMessage)
		}
		
		
		; Open the main dialong to get user decision on corrective action
		winMessage := "The script could not open the specified shortcut """ . shortcutName . ".lnk"". We can try to repair the shortcut right now...`n`nIs the shortcut a file?`n`nIf the shortcut is a folder, press ""No"". Otherwise press ""Cancel"" to manually repair the shortcut or ignore for now."
		Msgbox, 51, Shortcut Broken, %winMessage%
		
		; Get the target path from user (or exit)
		IfMsgBox, Yes
			FileSelectFile, targetPath,,, Select File,
		IfMsgBox, No
			FileSelectFolder, targetPath ,,, Select Folder...
		IfMsgBox, Cancel
			Exit
		
		
		; exit sub-routine if user did not select a folder
		if (ErrorLevel == 1){
			Exit
		}
		
		
		; Parse the target file/folder's parent directory
		; (this will serve as the shortcuts working directory)
		targetPathArray := StrSplit(targetPath, "\")
		targetPathArray.Pop()	;remove last element
		arrayLen := targetPathArray.MaxIndex()
		
		parentDirPath := targetPathArray[1]
		counter := 2
		while (counter <= arrayLen)
		{
			parentDirPath := parentDirPath . "\" . targetPathArray[counter]
			counter := counter + 1
		}
		
		; Create the new shortcut file
		FileCreateShortcut, %targetPath%, %shortcutPath%, %parentDirPath%
		
		; Exit the subroutine or hotkey to avoid further issues
		; User will need to press hotkey again to continue workflow
		Exit
		
	}

}




;###########################################################
;	PRIVATE GUI MANAGEMENT FUNCTIONS
;###########################################################


; similar to a plain MsgBox except that it does not pause the script
showMessageWindow(winMessage){
	Gui, OldTarget:+AlwaysOnTop -MinimizeBox
	Gui, OldTarget:Margin, 10, 10
	Gui, OldTarget:Color, White
	Gui, OldTarget:Font, s9, Segoe UI
	Gui, OldTarget:Add, Text, w400 +Wrap, %winMessage%
	Gui, OldTarget:Add, Button, w80 h25 x170 gOldTargetGUIClose, OK
	Gui, OldTarget:Show, xCenter y150 w420, Old Target Location
}


OldTargetGUIClose(){
	Gui, OldTarget:Destroy
}