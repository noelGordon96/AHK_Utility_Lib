;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	ExcelCom (Function Library)
; DESCRIPTION:	Provides functions for comunicating with a local instance of Microsoft Excel

				; NOTE: reframing approve on pulling Excel data
				; new code alredy created but needs to be tested and verified for reliability

; VERSION:		1.9.16.26
; AUTHOR:		Noel Gordon (veggieman1996@gmail.com)
; SCOURCE:		none


; ##########################################################
; USAGE AND OTHER INFO
; ##########################################################


; NOTE: These functions only work with local Excel running
; online Microsoft 365 sessions will not function properly


; ##########################################################
; 	COMPATABILITY AND DEPENDENCIES
; ##########################################################


#Requires AutoHotkey v2


;###########################################################
;	CONNECTION FUNCTIONS
;###########################################################




; Retrieve the COM object for a running Excel instance by reading it directly
; from the spreadsheet grid window. This is more reliable than
; ComObjActive("Excel.Application"), which depends on Excel being registered in
; the Running Object Table -- that can fail (e.g. when Excel and this script run
; at different elevation levels) or attach to the wrong instance when several
; Excel windows are open.
;
; WinTitle (optional): any AutoHotkey WinTitle targeting the desired Excel
;   window. Defaults to the top-most Excel main window (class XLMAIN).
;
; Returns the Excel.Application COM object. Throws an Error on failure.
/*
Excel_Get(WinTitle := "ahk_class XLMAIN")
{
	; Locate the Excel main window (class XLMAIN)
	if !(hwnd := WinExist(WinTitle))
		throw Error("No running Excel window was found. Open your workbook and try again.", -1)

	; Locate the worksheet grid child window (class EXCEL7 -> ClassNN EXCEL71).
	; This control exposes Excel's native object model to Active Accessibility.
	try
		gridHwnd := ControlGetHwnd("EXCEL71", "ahk_id " hwnd)
	catch
		throw Error("Found an Excel window but not its worksheet grid (EXCEL7). Make sure a workbook is open.", -1)

	; OBJID_NATIVEOM asks the window for its native Office object model
	OBJID_NATIVEOM := 0xFFFFFFF0
	IID_IDispatch  := "{00020400-0000-0000-C000-000000000046}"

	; Convert the interface-ID string into a 16-byte GUID buffer
	guid := Buffer(16, 0)
	if (DllCall("ole32\CLSIDFromString", "WStr", IID_IDispatch, "Ptr", guid) < 0)
		throw Error("Failed to build the IDispatch GUID.", -1)

	; Retrieve the Excel Window object from the grid control
	pWindow := 0
	hr := DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", gridHwnd, "UInt", OBJID_NATIVEOM, "Ptr", guid, "Ptr*", &pWindow)
	if (hr != 0 || !pWindow)
		throw Error("Could not read Excel's object model from the window (AccessibleObjectFromWindow failed).", -1)

	; Wrap the raw IDispatch pointer and return the owning Application object
	return ComObjFromPtr(pWindow).Application
}*/




;###########################################################
;	SPREADSHEET DATA RETRIEVAL FUNCTIONS
;###########################################################


/*
; NOT COMPLETED
; Get an array of the given length for the selected row within the Excel instance
ExcelCom_getSelectedRowData(cellLength)
{

	; ENSURE ONLINE VERSION OF SHEET IS NOT OPEN (ONLY WORKS WITH LOCAL INSTALLATION OF EXCEL)
	; Only checks for Chrome and Edge but could check for other popular browsers
	winProcess := WinGet("ProcessName")
	if (winProcess == "msedge.exe" OR winProcess == "chrome.exe"){
		MsgBox("ERROR: This script does not currently work with online version of Microsoft 365. Please open spreadsheet in local copy of Excel to continue.")
	}

	; RETRIEVE DATA FROM LOCAL INSTALLATION OF EXCEL
	else if (winProcess == "EXCEL.EXE"){

		; Connect to the running instance of Excel
		excelAppCom := ComObjActive("Excel.Application")
		
		
		; Define spreadsheet specific data (may need to change to match changes in spreadsheet)
		column_submitDate := 1
		column_manufacturer	:= 2
		column_revNum := 3
		
		
		; Get sheet selected row to determine the evaluation
		selectedRow := excelAppCom.ActiveWindow.RangeSelection.Row
		
		
		; Retrieve data from selected row and parse into usable folder name
		submitDate := excelAppCom.ActiveWorkbook.ActiveSheet.Cells(selectedRow, column_submitDate).Value
		manufacName := excelAppCom.ActiveWorkbook.ActiveSheet.Cells(selectedRow, column_manufacturer).Value
		evalRevNumber := excelAppCom.ActiveWorkbook.ActiveSheet.Cells(selectedRow, column_revNum).Value
		
		
		; Set COM objects to null to prevent lingering issues after close
		excelAppCom := ""
		
		
		; Parse excel data into usable evaluation folder path
		dateYear_2 := SubStr(submitDate, -1)
		submitDate := SubStr(submitDate, 1, StrLen(submitDate)-4)
		submitDate := submitDate . dateYear_2
		submitDate := StrReplace(submitDate, "/", "-")
		evalFolderName := manufacName . " Evaluation (" . evalRevNumber . ")" . " " . submitDate
		
		
		; Get device type from saves data or user if needed
		deviceTypeFolder := getEvalDeviceType(manufacName . " (" . evalRevNumber . ")")
		
		; piece together data for folder path
		currentEvalPath := "C:\Users\Noel Gordon\Oracle Content\Technology Operations Center\Validation\Current Evaluations" 
		evalFolderPath := currentEvalPath . "\" . deviceTypeFolder . "\" . manufacName . "\" . evalFolderName
		
		; Check if target directory exists before opening in file explorer
		if FileExist(evalFolderPath){
			Run, %evalFolderPath%
		}
		
		; Display error if evaluation directory does not exist
		else {
			Run, %currentEvalPath%
			sleep, 1000
			MsgBox, The following directory does not exist: %evalFolderPath%
		}
	}
}
*/


; Copy the contents of the currently selected cell in Excel to the clipboard
; Copies only plain text, not the cell formatting
ExcelCom_copyCellContents(pauseTime := 100){
	A_Clipboard := ""
	Send "{F2}"
	Sleep(pauseTime)
	Send "{Ctrl down}a{Ctrl up}"
	Sleep(pauseTime)
	Send "{Ctrl down}c{Ctrl up}"
	ClipWait(5)
	Send "{Esc}"
	return(A_Clipboard)
}


; Get an array of all cell values from the currently selected row in the
; active Excel instance. Values are read directly from the Excel COM object,
; so extraction is silent -- no clipboard use and no simulated keystrokes.
;
; cellLength (optional): number of columns to read, starting at column 1 (A).
;   When omitted or <= 0, the row's last used column is detected automatically
;   and every column up to it is returned.
;
; Returns an Array of cell values (index 1 = column A). Returns an empty Array
; if no running Excel instance is found or an error occurs.
/*
ExcelCom_getSelectedRowData_NEW(cellLength := 0)
{
	rowData := []

	; Connect to the running instance of Excel (fail gracefully if none)
	try {
		;excelAppCom := ComObjActive("Excel.Application") ;previous method
		excelAppCom := Excel_Get()
	} catch as err {
		MsgBox("ERROR: Could not connect to a running Excel instance.`n`n" . err.Message)
		return rowData
	}

	try {
		activeSheet := excelAppCom.ActiveWorkbook.ActiveSheet

		; Get the row of the current selection
		selectedRow := excelAppCom.ActiveWindow.RangeSelection.Row

		; Determine how many columns to read
		if (cellLength <= 0) {
			; xlToLeft = -4159 -- walk left from the sheet's last column to the
			; last populated cell to find the used width of this row
			lastColumn := activeSheet.Columns.Count
			cellLength := activeSheet.Cells(selectedRow, lastColumn).End(-4159).Column
		}

		; Read each cell value directly from the COM object
		Loop cellLength
			rowData.Push(activeSheet.Cells(selectedRow, A_Index).Value)
	} catch as err {
		MsgBox("ERROR: Failed to read the selected row from Excel.`n`n" . err.Message)
	}

	; Release the COM reference to avoid lingering instances
	excelAppCom := ""

	return rowData
}*/


; Get the contents of the currently selected (active) cell in Excel.
; Reads the value directly from the Excel COM object instead of sending
; keystrokes, so extraction is silent and does not disturb the user's current
; selection or edit state. The value is also mirrored to the clipboard to
; preserve the original "copy" behavior. Returns the plain cell value.
;
; Uses .Value, so a formula cell returns its computed result (not the formula)
; and numbers/text come back without display formatting ($, %, commas, etc.).
/*
ExcelCom_copyCellContents_NEW()
{
	cellValue := ""

	; Connect to the running instance of Excel (fail gracefully if none)
	try {
		excelAppCom := Excel_Get()
	} catch as err {
		MsgBox("ERROR: Could not connect to a running Excel instance.`n`n" . err.Message)
		return cellValue
	}

	; Read the active cell's value directly from the COM object
	try {
		cellValue := excelAppCom.ActiveCell.Value
	} catch as err {
		MsgBox("ERROR: Failed to read the active cell from Excel.`n`n" . err.Message)
	}

	; Release the COM reference to avoid lingering instances
	excelAppCom := ""

	return cellValue
}
*/