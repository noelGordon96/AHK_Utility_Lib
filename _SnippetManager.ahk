;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	Snippet Manager (Automation Snippet Library)
; DESCRIPTION:	Library module that catalogs reusable automation "snippets"
;				and shows a GUI to run, bind (to any key), edit, or create
;				them. #Include this file from your main script and call
;				SnippetManager_open() to show the catalog. No programmable
;				keyboard required, only AutoHotkey v2.
; VERSION:		1.9.9.26
; AUTHOR:		Noel Gordon (veggieman1996@gmail.com)
; SCOURCE:		none


; ##########################################################
; USAGE AND OTHER INFO
; ##########################################################


; FOLDER LAYOUT (paths resolve from your MAIN script = A_ScriptDir)
;	<main script>.ahk		<- your script that #Includes this library
;	settings.ini			<- per-machine values (window titles, etc.)
;	lib\					<- utility libraries (SnippetManager, ScreenAutomation, ...)
;	snippets\				<- one file per snippet (plus auto-generated _index.ahk)

; HOW TO USE
;	1) Include this library from your main script:
;			#Include "%A_ScriptDir%\lib\SnippetManager.ahk"
;	2) Call SnippetManager_open() to show the catalog (e.g. bind a key):
;			^+!Q::SnippetManager_open()
;	3) Add snippets with the "New Snippet" button (or copy
;		snippets\snippet_template.ahk), then press "Rescan & Reload".

; PUBLIC FUNCTIONS
;	SnippetManager_open()			Show (or re-show) the snippet catalog GUI
;	SnippetManager_register(info)	Register a snippet (called by each snippet file)
;	SnippetManager_rescan()			Rebuild the snippet index and reload the script

; AHK v2 DOCS REFERENCED
;	Hotkey():		https://www.autohotkey.com/docs/v2/lib/Hotkey.htm
;	GUI/ListView:	https://www.autohotkey.com/docs/v2/lib/GuiControls.htm
;	Loop Files:		https://www.autohotkey.com/docs/v2/lib/LoopFiles.htm


; ##########################################################
; 	COMPATABILITY AND DEPENDENCIES
; ##########################################################


#Requires AutoHotkey v2
#SingleInstance Force


; ##########################################################
; 	LIBRARY GLOBAL VARIABLES AND SETTINGS
; ##########################################################


; Allow other scripts to run this script in their own context (A_ScriptDir)
; so the snippet paths resolve correctly
scriptContext := A_ScriptDir
if (A_Args.Length > 0){
	scriptContext := A_Args[1]
}
else{
	checkForPreviousContext()
}


; catalog of registered snippets (filled by SnippetManager_register as snippets load)
snippetCatalog_obj := Array()

; gui handles (created when the catalog window is first opened)
snippetGui := ""
snippetListView := ""
snippetHotkeyInput := ""

; other global variables
;settingsFile := "settings.ini"
snippetDir := scriptContext "\snippets"

; snippet code editor path
;snippetEditor := "notepad.exe"	;TODO: make this configurable (settings.ini)
snippetEditor := "C:\Users\noel.gordon\AppData\Local\Programs\Microsoft VS Code\Code.exe"


;###########################################################
;	SNIPPET LOADING
;###########################################################

TraySetIcon(A_ScriptDir "\other_lib_files\key_record_icon.ico")

; build Include statement dynamically
buildDynamicContextCaller()
buildSnippetIndex()


; Load every snippet by including an auto-generated list of the files in the
; snippets folder (resolved from the MAIN script via A_ScriptDir). #Include is
; a load-time directive and cannot scan a folder, so the list is rebuilt by
; SnippetManager_rescan(), which then reloads the script. The *i flag keeps

; Actual include for _index.ahk needs to be dynamically generated
; (see comment on "contextCaller")
#Include "%A_ScriptDir%\other_lib_files\_snippet_contextCaller.ahk"

; open main window
SnippetManager_open()



; ##########################################################
; SNIPPET REGISTRATION
; ##########################################################



; Register a snippet with the catalog
; Called at load time by each file in the snippets folder
; snippetInfo is an object: {name, desc, targets, requires, file, run}
SnippetManager_register(snippetInfo_obj){
	global snippetCatalog_obj
	snippetCatalog_obj.Push(snippetInfo_obj)
}


; Rebuild the snippet index from the snippets folder and reload the script so
; new or renamed snippet files are picked up (#Include only runs at load time)
SnippetManager_rescan(){
	savePreviousContext()
	buildSnippetIndex()
	Reload()
}


;###########################################################
;	SNIPPET CATALOG GUI
;###########################################################


; Build (first call) or re-show the snippet catalog window
; Main entry point for the snippet catalog GUI (called by the main script)
SnippetManager_open(*){
	global snippetCatalog_obj, snippetGui, snippetListView, snippetHotkeyInput, snippetDir

	; just re-show the window if it was already built
	if (snippetGui != ""){
		snippetGui.Show()
		return
	}

	; create the main window
	snippetGui := Gui(, "Snippet Manager")
	;snippetGui.Opt("+Resize")
	snippetGui.SetFont("s10", "Segoe UI")

	; snippet list (name / description / target apps)
	snippetGui.Add("Text", "xm y+10 cblue", "Snippets Folder: " snippetDir) ;show custom context snippet directory
	snippetListView := snippetGui.Add("ListView", "r14 w640 Grid", ["Snippet", "Description", "Targets"])
	for index, snippet in snippetCatalog_obj{
		targetText := snippet.HasOwnProp("targets") ? joinArray(snippet.targets, ", ") : ""
		snippetListView.Add(, snippet.name, snippet.desc, targetText)
	}
	snippetListView.ModifyCol(1, 180)
	snippetListView.ModifyCol(2, 320)
	snippetListView.ModifyCol(3, 130)
	snippetListView.OnEvent("DoubleClick", runSelectedSnippet_handle)

	; key picker + bind button
	snippetGui.Add("Text", "xm y+10", "Key:")
	snippetHotkeyInput := snippetGui.Add("Hotkey", "x+6 yp-3 w180")
	bindButton := snippetGui.Add("Button", "x+10 yp-1 w120", "Bind to Key")
	bindButton.OnEvent("Click", bindSelectedSnippet_handle)

	; action buttons
	runButton := snippetGui.Add("Button", "xm y+10 w120", "Run Now")
	runButton.OnEvent("Click", runSelectedSnippet_handle)
	editButton := snippetGui.Add("Button", "x+10 w120", "Edit")
	editButton.OnEvent("Click", editSelectedSnippet_handle)
	newButton := snippetGui.Add("Button", "x+10 w120", "New Snippet")
	newButton.OnEvent("Click", createNewSnippet_handle)
	reloadButton := snippetGui.Add("Button", "x+10 w160", "Rescan && Reload")
	reloadButton.OnEvent("Click", reloadSnippetCatalog_handle)

	; hide (not destroy) on close so the window state is kept between opens
	snippetGui.OnEvent("Close", closeSnippetCatalog_handle)
	snippetGui.OnEvent("Escape", closeSnippetCatalog_handle)
	snippetGui.Show()
}


; Hide the catalog window (keeps the gui in memory for the next open)
closeSnippetCatalog_handle(*){
	global snippetGui
	snippetGui.Destroy()
	ExitApp()
}



;###########################################################
;	GUI ACTION HANDLERS
;###########################################################


; TODO: Reorganize GUI actions and handlers for better maintainability


;###########################################################
;	SNIPPET GUI ACTIONS
;###########################################################


; Return the snippet selected in the list (or "" if nothing is selected)
getSelectedSnippet(){
	global snippetCatalog_obj, snippetListView

	selectedRow := snippetListView.GetNext()
	if (!selectedRow){
		MsgBox("Select a snippet first.", "Snippet Manager", 0x30)
		return ""
	}
	return snippetCatalog_obj[selectedRow]
}


; Run the selected snippet right now (Run Now button / double-click)
runSelectedSnippet_handle(*){
	snippet := getSelectedSnippet()
	if (snippet == ""){
		return
	}
	WinMinimize(snippetGui.Hwnd) ; minimize Snippet Manager window before running the snippet
	snippet.run.Call()
}


; Bind the selected snippet to the key chosen in the hotkey box
bindSelectedSnippet_handle(*){
	global snippetHotkeyInput

	snippet := getSelectedSnippet()
	if (snippet == ""){
		return
	}

	; make sure a key was actually chosen
	chosenKey := snippetHotkeyInput.Value
	if (chosenKey == ""){
		MsgBox("Pick a key in the Key box first.", "Snippet Manager", 0x30)
		return
	}

	; Bind carries the selected snippet into the hotkey callback
	try {
		Hotkey(chosenKey, runBoundSnippet.Bind(snippet), "On")
		MsgBox("Bound `"" snippet.name "`" to " chosenKey, "Snippet Manager", 0x40)
		WinMinimize(snippetGui.Hwnd)
	}
	catch as err {
		MsgBox("Could not bind " chosenKey ":`r`n" err.Message, "Snippet Manager", 0x10)
	}
}


; Open the selected snippet's file in the editor for on-the-fly changes
editSelectedSnippet_handle(*){
	snippet := getSelectedSnippet()
	if (snippet == ""){
		return
	}
	if (!snippet.HasOwnProp("file")){
		MsgBox("This snippet did not report its file path.", "Snippet Manager", 0x30)
		return
	}
	editSnippetFile(snippet.file)
}


editSnippetFile(snippetFilePath){
	global snippetEditor
	runEditorString := "`"" snippetEditor "`" `"" snippetFilePath "`""
	Run(runEditorString)
}



; Create a new snippet file (copied from the template if present) and open it
createNewSnippet_handle(*){
	global snippetDir
	if (!DirExist(snippetDir)){
		DirCreate(snippetDir)
	}

	; ask for the new snippet's name
	namePrompt := InputBox("Name for the new snippet:", "New Snippet")
	if (namePrompt.Result != "OK" || Trim(namePrompt.Value) == ""){
		return
	}

	; normalize into a bare file name
	; - drop a typed .ahk
	; spaces to underscores
	; lowercase everything
	snippetName := Trim(namePrompt.Value)
	snippetName := RegExReplace(snippetName, "i)\.ahk$", "")
	snippetName := RegExReplace(snippetName, "[^\w]", "_")
	snippetName := StrLower(snippetName)
	newFile := snippetDir "\" snippetName ".ahk"

	; don't overwrite an existing snippet, just open it instead
	if (FileExist(newFile)){
		MsgBox("A snippet named `"" snippetName "`" already exists. Opening it instead.", "Snippet Manager", 0x40)
		Run("notepad.exe `"" newFile "`"")
		;TODO: CHANE EDITING PROGRAM...
		return
	}

	; start from the template if we have one, otherwise a minimal stub
	templateFile := snippetDir "\_snippet_template.ahk"
	if (FileExist(templateFile)){
		FileCopy(templateFile, newFile)
	}
	else {
		FileAppend(buildSnippetStub(snippetName), newFile)
	}

	; open the new file so it can be filled in right away
	replaceFileTemplatePlaceholders(snippetName, newFile)
	editSnippetFile(newFile)
	;MsgBox("Created snippets\" snippetName ".ahk`r`nFill it in, then press `"Rescan && Reload`" to load it.", "Snippet Manager", 0x40)
}


; Rebuild the snippet include list and reload so new/edited snippets appear
reloadSnippetCatalog_handle(*){
	SnippetManager_rescan()
}


;###########################################################
;	GUI ACTION HOTKEYS
;###########################################################



#HotIf WinActive("Snippet Manager")



Enter::
{
	global snippetListView

	; Predetermined variables for the focused control
	snippetListView_id := "SysListView321"
	hotkeySelectField_id := "msctls_hotkey321"
	bindSnippet_id := "Button1"

	
	; Get the HWND and class of the currently focused control
	SelectedControl_hwnd := ControlGetFocus()
	selectedControl_class := ControlGetClassNN(SelectedControl_hwnd)


	; Determine action based on the focused control
	if (selectedControl_class == snippetListView_id){

	selectedRow := snippetListView.GetNext()
	;MsgBox("Selected row: " selectedRow)
		
		if (selectedRow > 0){
			ControlFocus(hotkeySelectField_id)
		}
		else{
			;MsgBox("No row selected.")
			snippetListView.Modify(1, "+Focus +Select")
			;selectedRow := snippetListView.GetNext()
			;MsgBox("Selected row after focusing: " selectedRow)
		}

	}
	else if (selectedControl_class == hotkeySelectField_id){
		Send("{Tab}")
		;ControlFocus(bindSnippet_id)
	}
	else{
		Send("{Enter}")
	}

}



#HotIf



;###########################################################
;	PRIVATE UTILITY FUNCTIONS
;###########################################################


; Run a snippet handed in by Bind (used as the dynamic hotkey callback)
runBoundSnippet(snippet, *){
	snippet.run.Call()
}


; (Re)generate snippets\_index.ahk with an #Include line for each snippet file
; The template and the generated list itself are skipped
buildSnippetIndex(){
	global snippetDir
	if (!DirExist(snippetDir)){
		DirCreate(snippetDir)
	}

	; build the include list text (paths resolved from the main script dir)
	includeList := "; AUTO-GENERATED by SnippetManager. Do not edit by hand.`r`n"
	Loop Files, snippetDir "\*.ahk"{
		if (A_LoopFileName == "_index.ahk" || A_LoopFileName == "_snippet_template.ahk"){
			continue
		}
		newInclude := "#Include `"" snippetDir "\" A_LoopFileName "`"`r`n"
		includeList := includeList newInclude
	}

	; overwrite the old list
	indexFile := snippetDir "\_index.ahk"
	if (FileExist(indexFile)){
		FileDelete(indexFile)
	}
	FileAppend(includeList, indexFile)
}




; Bridges the caller's RUNTIME context into a LOAD-TIME #Include.
;
; SnippetManager runs standalone, so the folder it should load snippets
; from only arrives at runtime, as a parameter (scriptContext). But
; #Include is resolved while the script is being PARSED - before any code
; runs - and it only accepts built-in variables in its path. So it can
; never use scriptContext: that value doesn't exist yet at parse time, and
; a custom variable isn't allowed in an #Include anyway.
;
; The workaround: at runtime, write a tiny "bridge" file next to THIS
; script (A_ScriptDir is built-in, so #Include is allowed to use it). The
; bridge holds nothing but one plain, absolute #Include pointing at the
; caller's snippets\_index.ahk. Our own fixed #Include of that bridge file
; then loads it at parse time, which in turn pulls in the caller's snippets.
;
; Easy-to-forget gotcha: because #Include only fires at load time, a newly
; written bridge takes effect on the NEXT load. That's why this runs near
; startup and the script reloads itself once afterward. The file is rebuilt
; every run, so never edit _snippet_contextCaller.ahk by hand.
buildDynamicContextCaller(){
	codeLines := "; AUTO-GENERATED by SnippetManager. Do not edit by hand.`r`n"
	includePath := scriptContext "\snippets\_index.ahk"
	includeStatement := "#Include " includePath "`r`n"
	codeLines := codeLines includeStatement
	
	; overwrite the old list
	contextCallerFile := A_ScriptDir "\other_lib_files\_snippet_contextCaller.ahk"
	if (FileExist(contextCallerFile)){
		FileDelete(contextCallerFile)
	}
	FileAppend(codeLines, contextCallerFile)
}



; Minimal snippet contents used when no template file is available
buildSnippetStub(snippetName){
	stub := "#Requires AutoHotkey v2`r`n`r`n`r`n"
	stub := stub "SnippetManager_register({`r`n"
	stub := stub "`tname: `"" snippetName "`",`r`n"
	stub := stub "`tdesc: `"Describe what this snippet does`",`r`n"
	stub := stub "`ttargets: [],`r`n"
	stub := stub "`trequires: [],`r`n"
	stub := stub "`tfile: A_LineFile,`r`n"
	stub := stub "`trun: " snippetName "_run`r`n"
	stub := stub "})`r`n`r`n`r`n"
	stub := stub snippetName "_run(){`r`n"
	stub := stub "`tMsgBox(`"Replace me with the real automation.`", `"" snippetName "`", 0x40)`r`n"
	stub := stub "}`r`n"
	return stub
}




; TODO: finish function to avoid conflict snippet names (e.g. mySnippet_*)
; Remane placeolder values in template file to ensure unique names (e.g. mySnippet_*)...
replaceFileTemplatePlaceholders(snippetName, filePath){
	
	; DEFINE TEXT TO REPLACE IN FILE
	; Somewhat hard-coded for now making template edits harder
	searchText_1 := "<Snippet Name>"
	replaceText_1 := snippetName

	searchText_2 := "mySnippetPrefix_"
	replaceText_2 := snippetName "_"


	; Read file and perform the search and replaces
	fileContent := FileRead(filePath)

	updatedContent := StrReplace(fileContent, searchText_1, replaceText_1)
	updatedContent := StrReplace(updatedContent, searchText_2, replaceText_2)



	; Overwrite the file with the updated content
	fileObject := FileOpen(filePath, "w")
	fileObject.Write(updatedContent)
	fileObject.Close()
}



; Check if there is a previously saved snippet context
; Called at the beginning of a script to restore the previous snippet context if it exists
; Prevents "Rescann and Reload" button from loosing custom context
checkForPreviousContext(){
	global scriptContext
	contextFile := A_ScriptDir "\other_lib_files\_tempSnippetContext.txt"
	if (FileExist(contextFile)){
		scriptContext := FileRead(contextFile)
		FileDelete(contextFile)
	}
}


; Save the current snippet context to a temporary file
; Called by the "Rescann and Reload" button to save the current snippet context before reloading
savePreviousContext(){
	global scriptContext
	contextFile := A_ScriptDir "\other_lib_files\_tempSnippetContext.txt"
	FileAppend(scriptContext, contextFile)
}



; Join an array of strings into a single delimited string
joinArray(stringArray, delimiter){
	result := ""
	for index, value in stringArray{
		result := (result == "") ? value : result delimiter value
	}
	return result
}