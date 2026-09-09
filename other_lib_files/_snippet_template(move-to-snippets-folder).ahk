;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	<Snippet Name> (Automation Snippet)
; DESCRIPTION:	<description of what this snippet does>
; VERSION:		2.mm.dd.yy
; AUTHOR:		Noel Gordon (veggieman1996@gmail.com)
; SCOURCE:		none


; ##########################################################
; 	COMPATABILITY AND DEPENDENCIES
; ##########################################################


#Requires AutoHotkey v2.0

; bundle any utility libraries you need (paths resolve from your main script)
; #Include "%A_ScriptDir%\lib\ScreenAutomation.ahk"
; #Include "%A_ScriptDir%\lib\ExcelCom.ahk"


;###########################################################
;	SNIPPET REGISTRATION
;###########################################################


; register this snippet with the library catalog
; Has a syntax error here, but is available in the included SnippetManager.ahk library...
; so it will work when the snippet is loaded by the main script.
SnippetManager_register({
	name: "<Snippet Name>",
	desc: "<description>",
	targets: [],						; windows / apps it touches
	requires: [],						; libraries it depends on
	file: A_LineFile,					; lets the GUI "Edit" button open this file
	run: mySnippetPrefix_run			; entry point (function reference)
})



;###########################################################
;	SNIPPET LOGIC
;###########################################################


; Main entry point for the snippet (called by the library)
mySnippetPrefix_run(){
	
	

	; TODO: Rename all function in this file to ensure unque names (e.g. mySnippet_*)...
	; This ensure file script (with all snippets Included) can be loaded into the main script...
	; without function name collisions.
	

	; ---- >>> AUTOMATION LOGIC GOES HERE <<< ----
	MsgBox("TEMP: PLACEHOLDER ACTION.`n`r`n`rReplace me with the real automation.", "<Snippet Name>", 0x40)



}