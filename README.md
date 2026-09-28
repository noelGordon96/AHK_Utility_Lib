**NOTICE: This library has been converted to AutoHotkey V2! Old version 1 files can still be found in the respective sub-folder. Most updates moving forward will be for version 2. NOT ALL FILES AND FUNCTIONALILTIES CONVERTED TO VERSION 2 YET.**

# Description
This library contains various utility functions that I have collected over time to perform a variety of tasks within my AutoHotkey projects. See further details below on the various library files and the functions available in each. This library is written for AHKv2.

This is a work in progress... I am slowly adding to this as I consolodate functions from my various projects and work on making them more universaly useful. Let me know if you have suggestions for additions or changes.


# Function Usage
The function names in this library are written acording to the syntax **MyPrefix_MyFunc()** noted in the official AHK documentation: https://www.autohotkey.com/docs/v2/Scripts.htm#lib

~~The files are named as **MyPrifix.ahk** so as long as the files are in an accessible library by your script, you can call them in your scripts **MyPrefix_MyFunc(Params)**.~~

***NOTICE:** The migration to AHKv2 may have changed the above text. With the enhanced error checking in V2 this can cause issues because the functions show as not having a declaration (since the declaration is in a separate file). Manually including the library file with **#Include** solves this for now but I am still looking into alternatives like **#Warn**.*


# Function Libraries

## ShortcutManagement
Run and manages saved shortcuts and folder locations allowing you to easily connect to other system resources, programs, and files. The real benefit of this is that when attempting to run a deleted or moved shortcut, the library will help you repair the target location keeping your overall script running without errors. This can be useful in an organizational environment were rescources you connect to may be changed, migrated, or restructured at any point.

### Available Public Functions
***ShortcutManagement_getStoredFolderPath(locationName)***: Retrieve a saved location directory path. This allows user to include path locations in the parent (including) script. If the path does not exist or was not previouly saved, function assists the user walking them through selecting the folder and storing it for next time.

***ShortcutManagement_runShortcut(shortcutName)***: Run a saved shortcut (.lnk or .url) file. If the shortcut does not exist or is broken, the user is walked through the process of selecting the correct file and the shortcut is auotmatically fixed.<br>
*NOTE: Automatic repair currently only works with .lnk files. URL shortcut files need to be manually added into the "shortcutDir" folder (with the correct naming convention of "\_Shortcut\_\<shortcutName\>").*

***ShortcutManagement_checkRunShortcut(shortcutName, winTitle, userMessage)***: Run a saved shortcut as with previous funtion, but first prompt the user to either confirm, postpone, or cancel the action. (Example Usecase: Include this function in a small script run by Windows Task Scheduler to launch your email program everyday at a specific time, the key being now the scheduled task can be easily postponed at will)

## ManageDesktops
Create, manage, and easily move between Windows 11 virtual desktops. Includes some methods to "pull" windows with you when moving to or creating another desktop (this does not function for all windows).

### Available Public Functions
***ManageDesktops_moveWindowToVirtualDesktop()***: Pull (move) the active window to a specific virtual desktop (by 1 based number). Known non-working windows: Outlook

***ManageDesktops_moveWindowsToVirtualDesktop(destinationDesktop, windowTitleArray)***: Drag a number of windows to specific virtual desktop specified by an array of the window titles.

***ManageDesktops_switchDesktopByNumber(targetDesktop)***: Move to a specific virtual desktop (desktop must already exist)

***ManageDesktops_createVirtualDesktop()***: Create a new virtual desktop and switch to it.

***ManageDesktops_deleteVirtualDesktop()***: Delete the current virtual desktop.

***ManageDesktops_getVirtualDesktopName()***: Mostly used to return the name of the current virtual desktop. Windows desktop ID can also be passed in to get the name of another desktop.

***ManageDesktops_getVirtualDesktopCount()***: Return the current number of Windows virutal desktops.

***ManageDesktops_getCurrentDesktopNumber()***: Return the number (1 based) of the currently active virtual desktop.

***ManageDesktops_getVirtualDesktopNameArray()***: Return an array of virtual desktop names



## WindowManagement
Easily move and maximize windows to specific monitors.

### Available Public Functions

***WindowManagement_GetWinMon(winTitle)***: Return the monitor number on which the specified window (active window if omitted) is located. Monitor number corespond to the there numbers in your display settings. Returns 0 if the window is minimized.

***WindowManagement_MoveToMon(monitorNum)***: Move the active window to the specified monitor. Also resizes the window to fit on the monitor with some buffer area.

***WindowManagement_MaxOnMon(monitorNum, side := "full")***: Maximize the active window on the specified monitor if not already. Can also optionaly specify a side to set the window to take up half the monitor.



## UtilityWindows
This library includes some additional generic windows that can be displayed to the user for various purposes. Althoug not their main purpose, these window GUIs can also serve as good templates to copy and further customize within your code.

### Available Public Functions

***UtilityWindows_tranparentMessageWindow(winMessage, x := "Center", y := "Center", messageColor := "fc0303", messageSize := "24", messageFont := "New Courior", opacity := 0)***: This function displays a transparent window that displays a message on screen.

***UtilityWindows_closeTransparentWindow()***: Close the previous transparent window mentioned above. Since otherwise the window have some way of closing without exiting the script.


***UtilityWindows_parallelMessageBox(winTitle, winMessage, btnText := "OK")***: This function displays a simple message window to the user (similar to MsgBox). However because it it a custom gui, it does not interupt the rest of the script. So it can be used to display info to the user while still executing other lines of code.

***UtilityWindows_dontShowAgainMessage(winTitle, winMessage, hideWinKey, btnText := "OK")***: This function again shows a simple message to the user, however it included a "Don't show again" checkbox. The "hideWinKey" is what is used to remember each unique window type within the settings ini file.



## ExcelCom
This library contains funtions to interact with a localy running Excel, allowing you to easily read data into your script. Data is pulled directly from Excel's COM object (silently, without clipboard use or simulated keystrokes). *Only works with a local instance of Excel - online Microsoft 365 sessions will not function properly.*

### Available Public Functions

***ExcelCom_getCell(row, column)***: Return the value of a single cell by absolute position on the active sheet, addressed by 1-based ***row*** and ***column*** (1 = first row / column A), read silently from the COM object. This is the low-level primitive the other single-cell read helpers build on. Because it uses the cell's value, a formula cell returns its computed result (not the formula) and numbers/text come back without display formatting ($, %, commas, etc.). Returns an empty string if no running Excel instance is found, the coordinates are invalid, or an error occurs.

***ExcelCom_setCell(row, column, value)***: Write ***value*** into a single cell by absolute position on the active sheet, addressed by 1-based ***row*** and ***column*** (1 = first row / column A). The write goes directly to the COM object, so it is silent and does not use the clipboard or simulated keystrokes. This is the low-level primitive the other single-cell write helpers build on. Returns true on success, or false if no running Excel instance is found, the coordinates are invalid, or an error occurs.

***ExcelCom_getSelectedRow(maxColumns := 0)***: Return an Array of all cell values in the currently selected row (index 1 = column A), read silently from the Excel COM object. If ***maxColumns*** is omitted (or <= 0), the row's last used column is detected automatically and every column up to it is returned. Returns an empty Array if no running Excel instance is found.

***ExcelCom_getSelectedRowCell(column)***: Return the value of a single cell on the currently selected row, addressed by 1-based ***column*** (1 = column A, 2 = column B, ...). Resolves the current selection's row and delegates the read to ***ExcelCom_getCell***. Returns an empty string if no running Excel instance is found, the column number is invalid, or an error occurs.

***ExcelCom_setSelectedRowCell(column, value)***: Write ***value*** into a single cell on the currently selected row, addressed by 1-based ***column*** (1 = column A, 2 = column B, ...). Resolves the current selection's row and delegates the write to ***ExcelCom_setCell***. Returns true on success, or false if no running Excel instance is found, the column number is invalid, or an error occurs.

***ExcelCom_findColumnByHeader(header, headerRow := 1)***: Return the 1-based column number of the first cell in ***headerRow*** (default row 1) of the active sheet whose text exactly matches ***header*** (case-sensitive, not trimmed). The header row is read in a single COM call. Returns 0 (and shows a message) if no running Excel instance is found, the header is not found, or an error occurs.

***ExcelCom_getSelectedRowByHeader(header, headerRow := 1)***: Return the value of a cell on the currently selected row, addressed by the column whose header matches ***header*** (case-sensitive). Combines ***ExcelCom_findColumnByHeader*** with ***ExcelCom_getSelectedRowCell*** so you can work by header name instead of a raw column number. Returns an empty string if the header is not found or an error occurs.

***ExcelCom_setSelectedRowByHeader(header, value, headerRow := 1)***: Write ***value*** into a cell on the currently selected row, addressed by the column whose header matches ***header*** (case-sensitive). Combines ***ExcelCom_findColumnByHeader*** with ***ExcelCom_setSelectedRowCell*** so you can work by header name instead of a raw column number. Returns true on success, or false if the header is not found or an error occurs.

***ExcelCom_getActiveCell()***: Return the value of the currently active cell, read directly from the COM object. Because it uses the cell's value, a formula cell returns its computed result (not the formula) and numbers/text come back without display formatting ($, %, commas, etc.).

***ExcelCom_copyActiveCell(pauseTime := 100)***: The keystroke-based alternative - copies the active cell's plain text to the clipboard via F2 / Ctrl+A / Ctrl+C. Useful as a fallback when reading via the COM object is not desired.

***ExcelCom_selectCell(row, column)***: Move the selection to a specific cell on the active sheet, addressed by 1-based ***row*** and ***column*** (1 = column A). Returns true on success, or false if no running Excel instance is found or the cell cannot be selected.

***ExcelCom_getActiveCellLocation()***: Return the location of the currently active cell as an object with ***row*** and ***column*** properties (both 1-based). Returns false if no running Excel instance is found or no cell is active.

***ExcelCom_getExcelWinName()***: Return the title of the running Excel window. When several Excel windows are open, returns an Array of titles ordered by Z-order (most recently active first). Returns false if no Excel window is found.



## MiscUtilities
This library is for misc low level utility functions that might be usefull in a variety of contexts. Over time if patterns emerge, some of these might be moved into more specialized libraries.

### Available Public Functions

***MiscUtilities_betterRun(programPath, params*)***: This is a simple wrapper for the ***Run()*** function. It simply formats all arguments (intended for string params) with quotes around them to avoid manual quote formating everytime you use ***Run()*** with additional string arguments.
**Major Limitation:** very simple, does not yet detect non-string arguments, or strings with quotes already in them.



## ScreenAutomation (BETA)
*This library is a work in progress and it likey to change, so make sure you pay attention to the function parameters as these may also change between versions.* The purpose of this library is to assist with writing fairly complex scripts for things like screen/form automation where regular mouse movement and keyboard input are involved. --- The basic principal is you can use defined "locations" on your screen when writing the script and when you script runs it will prompt the user to store and remember these locations so on subsequent runs the script will know where on your screen which locations. I'll include more details an examples in later versions

### Available Public Functions
More detials on each of these will be included later.

***ScreenAutomation_copyBrowserAddressBar(windowTitle := "A")***: *Works with Google Chrome

***ScreenAutomation_waitForUser(message := "Confirm Action", continueKey := "Enter", messageColor := "fc0303")***

***ScreenAutomation_mouseMoveToLocation(locationName, anchorLocation := "")***

***ScreenAutomation_saveCurrentMouseLocation(locationName)***

***ScreenAutomation_clickLocation(locationName, pauseAfterClick := 0, requireConfirmation := false, verifyWinTitle := "", anchorLocation := "", autoRecord := false)***

***ScreenAutomation_deleteCoordinates(locationName := "")***

***ScreenAutomation_deleteCoordinates(locationName := "")***

***ScreenAutomation_sendText(textToSend, pauseAfterSend := 0)***



## SnippetManager (BETA)
Catalogs reusable automation "snippets" (small, self-contained tasks) and shows a GUI to run them, bind any snippet to any key on the fly, edit them, or create new ones from a template. Built so teammates with less experience with AutoHotkey can use the reuable code snippets as well. This library probably have some clean up that needs to be done, as well as some capabilities I want to add.

### Available Public Functions
***Note:*** Unlike other library files, this runs as a stand-alone script from the library folder. However to retain similarity to other libraries in your main script, this can be called using the Run() command. You can pass in a single "customContext" parameter which is the path you want the snippet manager to run from (A_ScriptDir from the calling script works well). This allows the "snippets" folder to be housed with your main script instead of in the library files.
