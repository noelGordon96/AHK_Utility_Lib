;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	ExcelCom (Function Library)
; DESCRIPTION:	Provides functions for communicating with a local instance of Microsoft Excel
; VERSION:		2.9.23.26
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
;	SPREADSHEET DATA INTERACTION FUNCTIONS
;###########################################################


; Get the value of a single cell by absolute position on the active sheet.
; This is the low-level primitive that the other single-cell read helpers build
; on. The value is read directly from the Excel COM object, so the read is
; silent -- no clipboard use and no simulated keystrokes.
;
; row    : 1-based row index (1 = first row).
; column : 1-based column index (1 = column A).
;
; Returns the cell value. Uses .Value, so a formula cell returns its computed
; result (not the formula) and numbers/text come back without display formatting
; ($, %, commas, etc.). Returns an empty string if no running Excel instance is
; found, the coordinates are invalid, or an error occurs.
ExcelCom_getCell(row, column)
{
	cellValue := ""

	; Guard against invalid coordinates
	if (row <= 0 || column <= 0) {
		MsgBox("ERROR: row and column must be positive integers (1 = first row / column A).")
		return cellValue
	}

	; Connect to the running instance of Excel (fail gracefully if none)
	try {
		excelAppCom := _Excel_Get()
	} catch as err {
		MsgBox("ERROR: Could not connect to a running Excel instance.`n`n" . err.Message)
		return cellValue
	}

	; Read the requested cell value directly from the COM object
	try {
		cellValue := excelAppCom.ActiveWorkbook.ActiveSheet.Cells(row, column).Value
	} catch as err {
		MsgBox("ERROR: Failed to read the cell from Excel.`n`n" . err.Message)
	}

	; Release the COM reference to avoid lingering instances
	excelAppCom := ""

	return cellValue
}



; Set the value of a single cell by absolute position on the active sheet.
; This is the low-level primitive that the other single-cell write helpers build
; on. The value is written directly to the Excel COM object, so the change is
; silent and does not rely on the clipboard or simulated keystrokes.
;
; row    : 1-based row index (1 = first row).
; column : 1-based column index (1 = column A).
; value  : the value to write into the cell.
;
; Returns true on success, false on failure.
ExcelCom_setCell(row, column, value)
{
	success := false

	; Guard against invalid coordinates
	if (row <= 0 || column <= 0) {
		MsgBox("ERROR: row and column must be positive integers (1 = first row / column A).")
		return success
	}

	; Connect to the running instance of Excel (fail gracefully if none)
	try {
		excelAppCom := _Excel_Get()
	} catch as err {
		MsgBox("ERROR: Could not connect to a running Excel instance.`n`n" . err.Message)
		return success
	}

	; Write the value directly to the COM object
	try {
		excelAppCom.ActiveWorkbook.ActiveSheet.Cells(row, column).Value := value
		success := true
	} catch as err {
		MsgBox("ERROR: Failed to write the cell to Excel.`n`n" . err.Message)
	}

	; Release the COM reference to avoid lingering instances
	excelAppCom := ""

	return success
}



; Get an array of all cell values from the currently selected row in the
; active Excel instance. Values are read directly from the Excel COM object,
; so extraction is silent -- no clipboard use and no simulated keystrokes.
;
; maxColumns (optional): number of columns to read, starting at column 1 (A).
;   When omitted or <= 0, the row's last used column is detected automatically
;   and every column up to it is returned.
;
; Returns an Array of cell values (index 1 = column A). Returns an empty Array
; if no running Excel instance is found or an error occurs.
ExcelCom_getSelectedRow(maxColumns := 0)
{
	rowData := []

	; Connect to the running instance of Excel (fail gracefully if none)
	try {
		;excelAppCom := ComObjActive("Excel.Application") ;previous method
		excelAppCom := _Excel_Get()
	} catch as err {
		MsgBox("ERROR: Could not connect to a running Excel instance.`n`n" . err.Message)
		return rowData
	}

	try {
		activeSheet := excelAppCom.ActiveWorkbook.ActiveSheet

		; Get the row of the current selection
		selectedRow := excelAppCom.ActiveWindow.RangeSelection.Row

		; Determine how many columns to read
		if (maxColumns <= 0) {
			; xlToLeft = -4159 -- walk left from the sheet's last column to the
			; last populated cell to find the used width of this row
			lastColumn := activeSheet.Columns.Count
			maxColumns := activeSheet.Cells(selectedRow, lastColumn).End(-4159).Column
		}

		; Read each cell value directly from the COM object
		Loop maxColumns
			rowData.Push(activeSheet.Cells(selectedRow, A_Index).Value)
	} catch as err {
		MsgBox("ERROR: Failed to read the selected row from Excel.`n`n" . err.Message)
	}

	; Release the COM reference to avoid lingering instances
	excelAppCom := ""

	return rowData
}



; Get the value of a single cell on the currently selected row, addressed by
; column number. Resolves the current selection's row and then delegates the
; read to ExcelCom_getCell.
;
; column: 1-based column index to read (1 = column A, 2 = column B, ...).
;
; Returns the cell value (see ExcelCom_getCell). Returns an empty string if no
; running Excel instance is found, the column number is invalid, or an error
; occurs.
ExcelCom_getSelectedRowCell(column)
{
	cellValue := ""

	; Guard against invalid column numbers
	if (column <= 0) {
		MsgBox("ERROR: column must be a positive integer (1 = column A).")
		return cellValue
	}

	; Connect to the running instance of Excel (fail gracefully if none)
	try {
		excelAppCom := _Excel_Get()
	} catch as err {
		MsgBox("ERROR: Could not connect to a running Excel instance.`n`n" . err.Message)
		return cellValue
	}

	; Resolve the row of the current selection
	selectedRow := 0
	try {
		selectedRow := excelAppCom.ActiveWindow.RangeSelection.Row
	} catch as err {
		MsgBox("ERROR: Failed to read the selected row from Excel.`n`n" . err.Message)
	}

	; Release the COM reference (ExcelCom_getCell re-connects to do the read)
	excelAppCom := ""

	; A row of 0 means the lookup above failed
	if (selectedRow <= 0)
		return cellValue

	return ExcelCom_getCell(selectedRow, column)
}



; Set the value of a single cell on the currently selected row, addressed by
; column number. Resolves the current selection's row and then delegates the
; write to ExcelCom_setCell.
;
; column: 1-based column index to write to (1 = column A, 2 = column B, ...).
; value : the value to write into the cell.
;
; Returns true on success, false on failure.
ExcelCom_setSelectedRowCell(column, value)
{
	success := false

	; Guard against invalid column numbers
	if (column <= 0) {
		MsgBox("ERROR: column must be a positive integer (1 = column A).")
		return success
	}

	; Connect to the running instance of Excel (fail gracefully if none)
	try {
		excelAppCom := _Excel_Get()
	} catch as err {
		MsgBox("ERROR: Could not connect to a running Excel instance.`n`n" . err.Message)
		return success
	}

	; Resolve the row of the current selection
	selectedRow := 0
	try {
		selectedRow := excelAppCom.ActiveWindow.RangeSelection.Row
	} catch as err {
		MsgBox("ERROR: Failed to read the selected row from Excel.`n`n" . err.Message)
	}

	; Release the COM reference (ExcelCom_setCell re-connects to do the write)
	excelAppCom := ""

	; A row of 0 means the lookup above failed
	if (selectedRow <= 0)
		return success

	return ExcelCom_setCell(selectedRow, column, value)
}



; Find the column number of a header cell by scanning a header row on the active
; sheet. Reads the entire header row in a single COM call, then scans it for an
; exact, case-sensitive match (values are not trimmed).
;
; header    : exact header text to locate (case-sensitive, untrimmed).
; headerRow : 1-based row that holds the headers (default 1).
;
; Returns the 1-based column number of the first match. Returns 0 (and shows a
; message) if no running Excel instance is found, the header is not found, or an
; error occurs.
ExcelCom_findColumnByHeader(header, headerRow := 1)
{
	columnNumber := 0

	; Connect to the running instance of Excel (fail gracefully if none)
	try {
		excelAppCom := _Excel_Get()
	} catch as err {
		MsgBox("ERROR: Could not connect to a running Excel instance.`n`n" . err.Message)
		return columnNumber
	}

	try {
		activeSheet := excelAppCom.ActiveWorkbook.ActiveSheet

		; xlToLeft = -4159 -- walk left from the sheet's last column to find the
		; used width of the header row
		lastColumn := activeSheet.Cells(headerRow, activeSheet.Columns.Count).End(-4159).Column

		if (lastColumn = 1) {
			; A single-cell range returns a scalar from .Value (not an array)
			if (activeSheet.Cells(headerRow, 1).Value == header)
				columnNumber := 1
		} else {
			; Read the whole header row at once (1-based 2D COM array)
			headerValues := activeSheet.Range(activeSheet.Cells(headerRow, 1), activeSheet.Cells(headerRow, lastColumn)).Value

			; Scan for an exact, case-sensitive match (== is case-sensitive in v2)
			Loop lastColumn {
				if (headerValues[1, A_Index] == header) {
					columnNumber := A_Index
					break
				}
			}
		}

		; Report a header that could not be located on the sheet
		if (columnNumber = 0)
			MsgBox("ERROR: Column header not found on the sheet: '" . header . "'")
	} catch as err {
		MsgBox("ERROR: Failed to read the header row from Excel.`n`n" . err.Message)
	}

	; Release the COM reference to avoid lingering instances
	excelAppCom := ""

	return columnNumber
}



; Get the value of a cell on the currently selected row, addressed by the column
; whose header matches `header`. Combines ExcelCom_findColumnByHeader with
; ExcelCom_getSelectedRowCell so callers can work by header name instead of a
; raw column number.
;
; header    : exact, case-sensitive header text to locate (untrimmed).
; headerRow : 1-based row that holds the headers (default 1).
;
; Returns the cell value, or an empty string if the header is not found or an
; error occurs.
ExcelCom_getSelectedRowByHeader(header, headerRow := 1)
{
	column := ExcelCom_findColumnByHeader(header, headerRow)
	if (column <= 0)
		return ""

	return ExcelCom_getSelectedRowCell(column)
}



; Set the value of a cell on the currently selected row, addressed by the column
; whose header matches `header`. Combines ExcelCom_findColumnByHeader with
; ExcelCom_setSelectedRowCell so callers can work by header name instead of a
; raw column number.
;
; header    : exact, case-sensitive header text to locate (untrimmed).
; value     : the value to write into the cell.
; headerRow : 1-based row that holds the headers (default 1).
;
; Returns true on success, false if the header is not found or an error occurs.
ExcelCom_setSelectedRowByHeader(header, value, headerRow := 1)
{
	column := ExcelCom_findColumnByHeader(header, headerRow)
	if (column <= 0)
		return false

	return ExcelCom_setSelectedRowCell(column, value)
}



; Get the value of the currently active cell in Excel. Reads the value directly
; from the Excel COM object instead of sending keystrokes, so extraction is
; silent and does not disturb the user's current selection or edit state.
;
; Uses .Value, so a formula cell returns its computed result (not the formula)
; and numbers/text come back without display formatting ($, %, commas, etc.).
; Returns an empty string if no running Excel instance is found or an error
; occurs.
ExcelCom_getActiveCell()
{
	cellValue := ""

	; Connect to the running instance of Excel (fail gracefully if none)
	try {
		excelAppCom := _Excel_Get()
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



; Copy the contents of the currently active cell in Excel to the clipboard.
; Copies only plain text using keystrokes instead of pulling data silently from
; the COM object. Useful as a fallback when reading via the COM object is not
; desired.
ExcelCom_copyActiveCell(pauseTime := 100){
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



;###########################################################
;	SPREADSHEET NAVIGATION FUNCTIONS
;###########################################################



; Move selection to a specific cell in the currently active Excel sheet.
; row: 1-based row index to move to (1 = first row).
; column: 1-based column index to move to (1 = column A).
ExcelCom_selectCell(row, column)
{
	; Connect to the running instance of Excel (fail gracefully if none)
	try {
		excelAppCom := _Excel_Get()
	} catch as err {
		MsgBox("ERROR: Could not connect to a running Excel instance.`n`n" . err.Message)
		return false
	}

	try {
		activeSheet := excelAppCom.ActiveWorkbook.ActiveSheet
		activeSheet.Cells(row, column).Select()
	} catch as err {
		MsgBox("ERROR: Failed to select the specified cell in Excel.`n`n" . err.Message)
		return false
	}

	; Release the COM reference to avoid lingering instances
	excelAppCom := ""

	return true
}



; Get the location of the currently active cell in the active Excel sheet.
; Returns an object with 'row' and 'column' properties (both 1-based), or false
; if no running Excel instance is found or no cell is active.
ExcelCom_getActiveCellLocation(){
	try {
		excelAppCom := _Excel_Get()
		activeSheet := excelAppCom.ActiveWorkbook.ActiveSheet
		selectedCell := activeSheet.Application.ActiveCell
		if !selectedCell
			throw Error("No cell is currently selected.", -1)
		return { row: selectedCell.Row, column: selectedCell.Column }
	} catch as err {
		MsgBox("ERROR: Failed to get the location of the active cell in Excel.`n`n" . err.Message)
		return false
	} finally {
		excelAppCom := ""
	}
}



;###########################################################
;	EXCEL INSTANCE FUNCTIONS
;###########################################################


; Get the title(s) of the currently running Excel window(s).
;
; Returns:
;   - the window title (String) when exactly one Excel window is open;
;   - an Array of window titles when several are open, ordered by Z-order
;     (most recently active first);
;   - false when no Excel window is found.
ExcelCom_getExcelWinName()
{
	; Enumerate all top-level Excel windows (class XLMAIN). WinGetList returns
	; handles in Z-order, so the most recently active window comes first.
	winList := WinGetList("ahk_class XLMAIN")

	if (winList.Length = 0)
		return false

	winNames := []
	for hwnd in winList
		winNames.Push(WinGetTitle("ahk_id " . hwnd))

	; Single window -> return its title directly; multiple -> return the array
	return (winNames.Length = 1) ? winNames[1] : winNames
}



;###########################################################
;	CONNECTION FUNCTIONS (PRIVATE)
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
_Excel_Get(WinTitle := "ahk_class XLMAIN")
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
}
