**NOTICE: This library has been converted to AutoHotkey V2! Old version 1 files can be found in the respective sub-folder. Most updates moving forward will be for version 2. NOT ALL FILES CONVERTED TO VERSION 2 YET.**

# Description
This library contains various utility functions that I have collected over time to perform a variety of tasks within my AutoHotkey projects. See further details below on the various library files and the functions available in each. This library is written for AHKv2.

This is a work in progress... I am slowly adding to this as I consolodate functions from my various projects and work on making them more universaly useful. Let me know if you have suggestions for additions or changes.


# Function Usage
The function names in this library are written acording to the syntax **MyPrefix_MyFunc()** noted in the official AHK documentation: https://www.autohotkey.com/docs/v2/Scripts.htm#lib

~~The files are named as **MyPrifix.ahk** so as long as the files are in an accessible library by your script, you can call them in your scripts **MyPrefix_MyFunc(Params)**.~~

***NOTICE:** The migration to AHKv2 may have changed the above text. With the enhanced error checking in V2 this can cause issues because the functions show as not having a declaration (since the declaration is in a separate file). Manually including the library file with **#Include** solves this for now but I am still looking into alternatives like **#Warn**.*


# Function Libraries

## ShortcutManagement
Run and manages saved shortcuts and folder locations allowing you to easily connect to other system resources, programs, and files. The real benefit of this is that when attempting to run a deleted or moved shortcut, the library will help you repair the target location. This can be useful in an organizational environment were rescource you connect to may be changed, migrated, or restructured.

### Available Public Functions
***ShortcutManagement_getStoredFolderPath(locationName)***: Retrieve a saved location directory path. This allows user to include path locations in the parent (including) script. If the path does not exist or was not previouly saved, function assists the user walking them through selecting the folder and storing it for next time.

***ShortcutManagement_runShortcut(shortcutName)***: Run a saved shortcut (.lnk) file. If the shortcut does not exist or is broken, the user is walked through the process of selecting the correct file and the shortcut is auotmatically fixed.


## ManageDesktops (Not Converted)
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



## WindowManagement (Not Converted)
Easily move and maximize windows to specific monitors.

### Available Public Functions

***WindowManagement_GetWinMon(winTitle)***: Return the monitor number on which the specified window (active window if omitted) is located. Monitor number corespond to the there numbers in your display settings. Returns 0 if the window is minimized.

***WindowManagement_MoveToMon(monitorNum)***: Move the active window to the specified monitor. Also resizes the window to fit on the monitor with some buffer area.

***WindowManagement_MaxOnMon(monitorNum)***: Maximize the active window on the specified monitor if not already.



## UtilityWindows
This library includes some aditional generic windows that can be displayed to the user for various purposes. Althoug not their main purpose, these window GUIs can also serve as good templates to copy and further customize within your code.

### Available Public Functions

***UtilityWindows_parallelMessageBox(winTitle, winMessage, btnText := "OK")***: This function displays a simple message window to the user (similar to MsgBox). However because it it a custom gui, it does not interupt the rest of the script. So it can be used to display info to the user while still executing other lines of code.

***UtilityWindows_dontShowAgainMessage(winTitle, winMessage, hideWinKey, btnText := "OK")***: This function again shows a simple message to the user, however it included a "Don't show again" checkbox. The "hideWinKey" is what is used to remember each unique window type within the settings ini file.
