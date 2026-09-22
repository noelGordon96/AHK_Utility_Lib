;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	Snippet Manager (Automation Snippet Library)
; DESCRIPTION:	Library module that catalogs reusable automation "snippets"
;				and shows a GUI to run, bind (to any key), edit, or create
;				them. #Include this file from your main script and call
;				SnippetManager_open() to show the catalog. No programmable
;				keyboard required, only AutoHotkey v2.
; VERSION:		2.9.22.26
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

; list of currently bound snippets, each an object {key, snippet}
; (filled as snippets are bound so the GUI can show what is live)
boundSnippets_obj := Array()

; snippets currently shown in the top list, in row order
; (bound snippets are filtered out, so row number != snippetCatalog_obj index)
displayedSnippets_obj := Array()

; gui handles (created when the catalog window is first opened)
snippetGui := ""
snippetListView := ""
snippetHotkeyInput := ""
snippetBoundListView := ""
snippetBindButton := ""		; kept global so it can be greyed out per active list

; snippet details dialog (Name / Description / File Name) - reused for new
; snippets now and, later, for editing an existing snippet's details
snippetDetailsGui := ""
snippetDetailsName := ""
snippetDetailsDesc := ""
snippetDetailsFile := ""
snippetDetailsOnSubmit := ""		; callback run with (name, desc, fileName) on OK
snippetDetailsFileEdited := false	; user has manually edited the File Name field
snippetDetailsFileAutoValue := ""	; last value we auto-wrote into the File Name field

; other global variables
;settingsFile := "settings.ini"
snippetDir := scriptContext "\snippets"
settingsFile := scriptContext "\settings.ini"

; snippet code editor path
snippetEditor := "C:\Users\" A_UserName "\AppData\Local\Programs\Microsoft VS Code\Code.exe" ;TODO: make this configurable (settings.ini)


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
SnippetManager_openGui()



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
;	MAIN SNIPPET CATALOG GUI
;###########################################################


; Build (first call) or re-show the snippet catalog window
; Main entry point for the snippet catalog GUI (called by the main script)
SnippetManager_openGui(*){
	global snippetCatalog_obj, snippetGui, snippetListView, snippetHotkeyInput, snippetBoundListView, snippetBindButton, snippetDir

	; just re-show the window if it was already built
	if (snippetGui != ""){
		snippetGui.Show()
		return
	}

	; create the main window
	snippetGui := Gui(, "Snippet Manager")
	snippetGui.Opt("+Resize")
	snippetGui.SetFont("s10", "Segoe UI")

	; snippet list (name / description)
	snippetGui.Add("Text", "xm y+10 cblue", "Snippets Folder: " snippetDir) ;show custom context snippet directory
	snippetListView := snippetGui.Add("ListView", "r10 w820 Grid", ["Snippet", "Description"])
	snippetListView.ModifyCol(1, 200)
	snippetListView.ModifyCol(2, 600)
	snippetListView.OnEvent("DoubleClick", runSelectedSnippet_handle)
	snippetListView.OnEvent("ItemSelect", topListSelected_handle)
	refreshSnippetList()

	; key picker + bind button
	snippetGui.Add("Text", "xm y+10", "Key:")
	snippetHotkeyInput := snippetGui.Add("Hotkey", "x+6 yp-3 w180")
	snippetBindButton := snippetGui.Add("Button", "x+10 yp-1 w120", "Bind to Key")
	snippetBindButton.OnEvent("Click", bindSelectedSnippet_handle)

	; action buttons
	runButton := snippetGui.Add("Button", "xm y+10 w120", "Run Now")
	runButton.OnEvent("Click", runSelectedSnippet_handle)
	editButton := snippetGui.Add("Button", "x+10 w120", "Edit Code")
	editButton.OnEvent("Click", editSelectedSnippet_handle)
	newButton := snippetGui.Add("Button", "x+10 w120", "New Snippet")
	newButton.OnEvent("Click", createNewSnippet_handle)
	browseButton := snippetGui.Add("Button", "x+10 w140", "Browse Snippets")
	browseButton.OnEvent("Click", browseSnippets_handle)
	reloadButton := snippetGui.Add("Button", "x+10 w160", "Rescan && Reload")
	reloadButton.OnEvent("Click", reloadSnippetCatalog_handle)

	; currently bound snippets (which key runs which snippet right now)
	snippetGui.Add("Text", "xm y+15", "Currently Bound Snippets:")
	snippetBoundListView := snippetGui.Add("ListView", "r6 w820 Grid", ["Key", "Snippet", "Description"])
	snippetBoundListView.ModifyCol(1, 100)
	snippetBoundListView.ModifyCol(2, 200)
	snippetBoundListView.ModifyCol(3, 500)
	snippetBoundListView.OnEvent("DoubleClick", runSelectedSnippet_handle)
	snippetBoundListView.OnEvent("ItemSelect", boundListSelected_handle)
	refreshBoundSnippets()

	; hide (not destroy) on close so the window state is kept between opens
	snippetGui.OnEvent("Close", closeSnippetCatalog_handle)
	snippetGui.OnEvent("Escape", closeSnippetCatalog_handle)
	snippetGui.Show()
}



;###########################################################
;	MAIN GUI ACTION HANDLERS
;###########################################################


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
		;WinMinimize(snippetGui.Hwnd) ; keep the GUI up after binding (may re-enable later)
		Sleep(200)
		snippet.on_bind.Call()
		Hotkey(chosenKey, runBoundSnippet.Bind(snippet), "On")
		recordBoundSnippet(chosenKey, snippet)
		refreshSnippetList()
		refreshBoundSnippets()
		MsgBox("Bound `"" snippet.name "`" to " chosenKey, "Snippet Manager", 0x40)
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



; Ask for the new snippet's details, then create the file (see the callback)
createNewSnippet_handle(*){
	openSnippetDetailsGui("New Snippet", "", "", "", createSnippetFromDetails)
}


; Open the snippets folder in File Explorer
browseSnippets_handle(*){
	global snippetDir
	if (!DirExist(snippetDir)){
		DirCreate(snippetDir)
	}
	Run("explorer.exe `"" snippetDir "`"")
}


; Rebuild the snippet include list and reload so new/edited snippets appear
reloadSnippetCatalog_handle(*){
	SnippetManager_rescan()
}


; Closing the snippet catalog window and exiting the application
closeSnippetCatalog_handle(*){
	global snippetGui
	snippetGui.Destroy()
	ExitApp()
}


;###########################################################
;	SNIPPET DETAILS DIALOG (Name / Description / File Name)
;###########################################################


; Open the shared details dialog for entering/editing a snippet's Name,
; Description and File Name. On OK, submitCallback(name, desc, fileName) runs.
; Built to be reused later for an "Edit Details" flow (hence the prefill args).
openSnippetDetailsGui(title, prefillName, prefillDesc, prefillFile, submitCallback){
	global snippetGui
	global snippetDetailsGui, snippetDetailsName, snippetDetailsDesc, snippetDetailsFile
	global snippetDetailsOnSubmit, snippetDetailsFileEdited, snippetDetailsFileAutoValue

	; only one details dialog at a time
	if (snippetDetailsGui != ""){
		snippetDetailsGui.Show()
		return
	}

	snippetDetailsOnSubmit := submitCallback
	snippetDetailsFileAutoValue := ""
	; a prefilled file name (edit flow) counts as already set, so typing in the
	; Name field won't clobber it; a blank one (new flow) auto-fills from Name
	snippetDetailsFileEdited := (Trim(prefillFile) != "")

	snippetDetailsGui := Gui("+Owner" snippetGui.Hwnd, title)
	snippetDetailsGui.SetFont("s10", "Segoe UI")

	snippetDetailsGui.Add("Text", "xm y+10", "Name:")
	snippetDetailsName := snippetDetailsGui.Add("Edit", "xm y+2 w360", prefillName)
	snippetDetailsName.OnEvent("Change", snippetDetailsNameChanged_handle)

	snippetDetailsGui.Add("Text", "xm y+10", "Description:")
	snippetDetailsDesc := snippetDetailsGui.Add("Edit", "xm y+2 w360 r3 Multi WantReturn", prefillDesc)

	snippetDetailsGui.Add("Text", "xm y+10", "File Name:")
	snippetDetailsFile := snippetDetailsGui.Add("Edit", "xm y+2 w360", prefillFile)
	snippetDetailsFile.OnEvent("Change", snippetDetailsFileChanged_handle)
	snippetDetailsGui.Add("Text", "xm y+2 cGray", ".ahk is added automatically")

	okButton := snippetDetailsGui.Add("Button", "xm y+15 w110 Default", "OK")
	okButton.OnEvent("Click", snippetDetailsSubmit_handle)
	cancelButton := snippetDetailsGui.Add("Button", "x+10 w110", "Cancel")
	cancelButton.OnEvent("Click", closeSnippetDetails_handle)

	snippetDetailsGui.OnEvent("Close", closeSnippetDetails_handle)
	snippetDetailsGui.OnEvent("Escape", closeSnippetDetails_handle)
	snippetDetailsGui.Show()
}


; Name changed: keep the File Name field in sync with a normalized version of
; the name, until the user manually edits the File Name themselves. We remember
; the value we write so the File Name's own Change event can tell our auto-fill
; apart from a real user edit (the Change event may arrive after this returns)
snippetDetailsNameChanged_handle(*){
	global snippetDetailsName, snippetDetailsFile, snippetDetailsFileEdited, snippetDetailsFileAutoValue
	if (snippetDetailsFileEdited){
		return
	}
	snippetDetailsFileAutoValue := normalizeSnippetFileName(snippetDetailsName.Value)
	snippetDetailsFile.Value := snippetDetailsFileAutoValue
}


; File Name changed: if the new value is one WE wrote (matches the last
; auto-filled value), ignore it; otherwise the user typed it, so lock the
; field and stop syncing it from the Name field
snippetDetailsFileChanged_handle(*){
	global snippetDetailsFile, snippetDetailsFileEdited, snippetDetailsFileAutoValue
	if (snippetDetailsFileEdited){
		return
	}
	if (snippetDetailsFile.Value == snippetDetailsFileAutoValue){
		return
	}
	snippetDetailsFileEdited := true
}


; OK: validate, normalize the file name, hand values to the submit callback
snippetDetailsSubmit_handle(*){
	global snippetDetailsName, snippetDetailsDesc, snippetDetailsFile, snippetDetailsOnSubmit

	snippetName := Trim(snippetDetailsName.Value)
	snippetDesc := Trim(snippetDetailsDesc.Value)
	fileName := normalizeSnippetFileName(snippetDetailsFile.Value)

	if (snippetName == ""){
		MsgBox("Enter a name for the snippet.", "Snippet Manager", 0x30)
		return
	}
	if (fileName == ""){
		MsgBox("Enter a valid file name for the snippet.", "Snippet Manager", 0x30)
		return
	}

	callback := snippetDetailsOnSubmit
	closeSnippetDetails_handle()
	callback.Call(snippetName, snippetDesc, fileName)
}


; Close the details dialog and clear its handles so it can be reopened cleanly
closeSnippetDetails_handle(*){
	global snippetDetailsGui, snippetDetailsOnSubmit
	if (snippetDetailsGui != ""){
		snippetDetailsGui.Destroy()
		snippetDetailsGui := ""
	}
	snippetDetailsOnSubmit := ""
}


;###########################################################
;	MAIN GUI ACTION FUNCTIONS
;###########################################################


; A row was selected in the top (unbound) list: clear the bound list's
; selection so only one list is ever active, and re-enable the Bind button.
; We only react to a real selection (Selected == 1); the deselect we trigger
; on the other list fires Selected == 0, which we ignore (avoids any loop)
topListSelected_handle(ctrl, item, selected){
	global snippetBoundListView, snippetBindButton
	if (!selected){
		return
	}
	snippetBoundListView.Modify(0, "-Select")	; row 0 = all rows
	snippetBindButton.Enabled := true
}


; A row was selected in the bound list: clear the top list's selection and
; grey out Bind (the snippet is already bound, so binding again is meaningless)
boundListSelected_handle(ctrl, item, selected){
	global snippetListView, snippetBindButton
	if (!selected){
		return
	}
	snippetListView.Modify(0, "-Select")		; row 0 = all rows
	snippetBindButton.Enabled := false
}


; Return the snippet selected in either list (or "" if nothing is selected)
; Only one list can have a selection at a time (see the ItemSelect handlers),
; so we check the top list first, then fall back to the bound list
getSelectedSnippet(){
	global displayedSnippets_obj, boundSnippets_obj, snippetListView, snippetBoundListView

	topRow := snippetListView.GetNext()
	if (topRow){
		return displayedSnippets_obj[topRow]
	}

	boundRow := snippetBoundListView.GetNext()
	if (boundRow){
		return boundSnippets_obj[boundRow].snippet
	}

	MsgBox("Select a snippet first.", "Snippet Manager", 0x30)
	return ""
}


editSnippetFile(snippetFilePath){
	global snippetEditor
	runEditorString := "`"" snippetEditor "`" `"" snippetFilePath "`""
	Run(runEditorString)
}


; Create a new snippet file from the details dialog values and open it
; name = display name, desc = description, fileName = normalized file/prefix
createSnippetFromDetails(name, desc, fileName){
	global snippetDir
	if (!DirExist(snippetDir)){
		DirCreate(snippetDir)
	}

	newFile := snippetDir "\" fileName ".ahk"

	; don't overwrite an existing snippet, just open it instead
	if (FileExist(newFile)){
		MsgBox("A snippet file named `"" fileName "`" already exists. Opening it instead.", "Snippet Manager", 0x40)
		Run("notepad.exe `"" newFile "`"")
		;TODO: CHANE EDITING PROGRAM...
		return
	}

	; start from the user's own template if they have one in the snippets
	; folder; otherwise fall back to the standard library template shipped in
	; other_lib_files (users can copy that one into snippets\ and edit it)
	templateFile := snippetDir "\_snippet_template.ahk"
	if (!FileExist(templateFile)){
		templateFile := A_ScriptDir "\other_lib_files\_snippet_template.ahk"
	}
	if (!FileExist(templateFile)){
		MsgBox("Could not find a snippet template:`r`n" templateFile, "Snippet Manager", 0x10)
		return
	}
	FileCopy(templateFile, newFile)
	replaceFileTemplatePlaceholders(name, desc, fileName, newFile)

	; open the new file so it can be filled in right away
	editSnippetFile(newFile)
	;MsgBox("Created snippets\" fileName ".ahk`r`nFill it in, then press `"Rescan && Reload`" to load it.", "Snippet Manager", 0x40)
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


; Remember that a key now runs a snippet so the GUI can list it
; If the key was already bound, its entry is replaced (the newest bind wins)
recordBoundSnippet(key, snippet){
	global boundSnippets_obj
	for index, entry in boundSnippets_obj{
		if (entry.key == key){
			boundSnippets_obj[index] := {key: key, snippet: snippet}
			return
		}
	}
	boundSnippets_obj.Push({key: key, snippet: snippet})
}


; Repaint the "Currently Bound Snippets" list from boundSnippets_obj
; Safe to call before the GUI exists (does nothing until the list is built)
refreshBoundSnippets(){
	global boundSnippets_obj, snippetBoundListView
	if (snippetBoundListView == ""){
		return
	}
	snippetBoundListView.Delete()
	for index, entry in boundSnippets_obj{
		snippet := entry.snippet
		snippetBoundListView.Add(, entry.key, snippet.name, snippet.desc)
	}
}


; Repaint the top snippet list from the catalog, skipping snippets that are
; already bound. displayedSnippets_obj is rebuilt in step with the rows so a
; selected row maps back to the right snippet (see getSelectedSnippet)
refreshSnippetList(){
	global snippetCatalog_obj, snippetListView, displayedSnippets_obj
	if (snippetListView == ""){
		return
	}
	snippetListView.Delete()
	displayedSnippets_obj := Array()
	for index, snippet in snippetCatalog_obj{
		if (isSnippetBound(snippet)){
			continue
		}
		snippetListView.Add(, snippet.name, snippet.desc)
		displayedSnippets_obj.Push(snippet)
	}
}


; True if this snippet currently has a key bound to it
isSnippetBound(snippet){
	global boundSnippets_obj
	for index, entry in boundSnippets_obj{
		if (entry.snippet == snippet){
			return true
		}
	}
	return false
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
	;contextCallerFile_quotes := Format('"{1}"', contextCallerFile)
	FileAppend(codeLines, contextCallerFile)
}



; Normalize a raw name into a bare, safe file/function identifier:
; drop a typed .ahk, turn non-word characters into underscores, lowercase it
normalizeSnippetFileName(rawName){
	fileName := Trim(rawName)
	fileName := RegExReplace(fileName, "i)\.ahk$", "")
	fileName := RegExReplace(fileName, "[^\w]", "_")
	fileName := StrLower(fileName)
	return fileName
}


; Escape text so it is safe to embed inside a double-quoted AHK v2 string
; literal: escape the escape char and quotes, and turn real newlines into `n
; (Chr(96) = backtick escape char, Chr(34) = double quote)
escapeForAhkString(text){
	backtick := Chr(96)
	quote := Chr(34)
	text := StrReplace(text, backtick, backtick backtick)	; ` -> ``
	text := StrReplace(text, quote, backtick quote)			; " -> `"
	text := StrReplace(text, "`r`n", backtick "n")			; CRLF -> `n
	text := StrReplace(text, "`n", backtick "n")			; LF   -> `n
	text := StrReplace(text, "`r", backtick "n")			; CR   -> `n
	return text
}


; Fill the template placeholders with the snippet's real details
; displayName -> the human-friendly name (header + register name:)
; description -> register desc: and the header description line
; filePrefix  -> the normalized identifier used for the snippet's functions
replaceFileTemplatePlaceholders(displayName, description, filePrefix, filePath){

	; Read file and perform the search and replaces
	fileContent := FileRead(filePath)

	; the header line is a comment, so flatten any newlines to keep it one line;
	; the register desc: is a string literal, so escape it for safe embedding
	descForComment := Trim(RegExReplace(description, "\R+", " "))
	descForString := escapeForAhkString(description)

	; replace the longer description placeholder first so it isn't partially
	; matched by the shorter "<description>" token
	updatedContent := StrReplace(fileContent, "<description of what this snippet does>", descForComment)
	updatedContent := StrReplace(updatedContent, "<description>", descForString)
	updatedContent := StrReplace(updatedContent, "<Snippet Name>", displayName)
	updatedContent := StrReplace(updatedContent, "mySnippetPrefix_", filePrefix "_")

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