;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	Misc Utilities (Automation Library)
; DESCRIPTION:	Simple miscellaneous utility functions for
;				use with a variery of scripts and tasks.
;				Over time some of these might move to more
;				specialized libraries.
; VERSION:		1.9.9.26
; AUTHOR:		Noel Gordon (veggieman1996@gmail.com)
; SCOURCE:		none


; ##########################################################
; 	COMPATABILITY AND DEPENDENCIES
; ##########################################################


#Requires AutoHotkey v2


; ##########################################################
; 	LIBRARY GLOBAL VARIABLES AND SETTINGS
; ##########################################################


; none


; ##########################################################
; 	MISC UTILITY FUNCTIONS
; ##########################################################


; Attempt to run the specified program with the provided arguments
; The key improvement currently is that it properly handles multiple...
; ... arguments and ensures they are correctly quoted in the final
; ... run string. Allows for flexible argument passing without having
; ... having to manually quote each string argument.
; TODO: allow for handling non-string arguments
MiscUtilities_betterRun(programPath, params*){

	; build argument string (including quotations) for the params passed in
	args := ""
	for index, value in params {
		args := args . Format(' "{1}"', value)
	}

	; build the final script run string with the program path and arguments
	scriptRunString := Format('"{1}" {2}', programPath, args)
	Run(scriptRunString)
}
