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

; bundle any utility libraries you need
;#Include "%A_ScriptDir%\ScreenAutomation.ahk" ; example - change include path


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
	run: mySnippetPrefix_run,			; entry point (function reference)
	on_bind: mySnippetPrefix_on_bind	; one-time binding logic for the snippet
})



;###########################################################
;	SNIPPET LOGIC
;###########################################################


; Main entry point for the snippet (called by the library)
mySnippetPrefix_run(){
	
	; ---- >>> AUTOMATION LOGIC GOES HERE <<< ----

	MsgBox("TEMP: PLACEHOLDER ACTION.`n`r`n`rReplace me with the real automation.", "<Snippet Name>", 0x40)

}


; One time code that the Snippet Manager calls when binding the snippet
mySnippetPrefix_on_bind(){
	
	; ---- >>> BINDING LOGIC GOES HERE <<< ----

	MsgBox("Replace the ...on_bind() function with logic you would like run here...", "<Snippet Name>", 0x40)
	;ScreenAutomation_deleteCoords() ; example code

}