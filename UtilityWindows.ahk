;###########################################################
;	INFO HEADER
;###########################################################


; SCRIPT NAME:	Utility Windows (Function Library)
; DESCRIPTION:	Provides custom utility windows for various purposes.
; VERSION:		2.2.3.25
; AUTHOR:		Noel Gordon (veggieman1996@gmail.com)
; SCOURCE:		none


; ##########################################################
; USAGE AND OTHER INFO
; ##########################################################


; REQUIRED FILES AND SETTINGS

; "settingsFile" global variable: this variable should denote file name
; of the hotkey projects settings ini file, most often this should be defined
; in the following manner near the begining of the main script:

; settingsFile := "settings.ini"



;###########################################################
;	TRANSPARENT MESSAGE WINDOW
;###########################################################



; similar to a plain MsgBox except that it does not pause the script
UtilityWindows_tranparentMessageWindow(winMessage, x := "Center", y := "Center", messageColor := "fc0303") {
    
    ; create gui and add contents and controls
    global msgWin := Gui(, "Transparent Window")
	msgWin.Opt("+AlwaysOnTop +LastFound +ToolWindow -Caption")
    msgWin.MarginX := 10
    msgWin.MarginY := 10
	msgWin.BackColor := "ffffff"
    msgWin.SetFont("s24 bold c" . messageColor, "New Courior")
    msgWin.Add("Text",, winMessage)
    WinSetTransparent(100, msgWin.hwnd)
    
	msgWin.OnEvent("Close", UtilityWindows_closeTransparentWindow)
    msgWin.Show("x" . x . " y" . y)

}

; gui internal methods
UtilityWindows_closeTransparentWindow(*){
    msgWin.Destroy()
}



;###########################################################
;	PARALLEL MESSAGE WINDOW
;###########################################################



; similar to a plain MsgBox except that it does not pause the script
UtilityWindows_parallelMessageBox(winTitle, winMessage, btnText := "OK"){
    
    ; create gui and add contents and controls
    msgWin := Gui(, winTitle)
	msgWin.Opt("+AlwaysOnTop -MinimizeBox")
    msgWin.MarginX := 10
    msgWin.MarginY := 10
	msgWin.BackColor := "ffffff"
    msgWin.SetFont("s9", "Segoe UI")
    msgWin.Add("Text", "w400 +Wrap", winMessage)
    
    closeBtn := msgWin.Add("Button", "Default w80 h25 x170", btnText)
    closeBtn.OnEvent("Click", closeMessageWindow)

	msgWin.OnEvent("Close", closeMessageWindow)
    msgWin.Show("xCenter y150 w420")


    ; gui internal methods
    closeMessageWindow(*){
        msgWin.Destroy()
    }
}



;###########################################################
;	DONT SHOW AGAIN WINDOW
;###########################################################


; display a simple message to user along with a "Don't show again" check box
UtilityWindows_dontShowAgainMessage(winTitle, winMessage, hideWinKey, btnText := "OK"){
    global settingsFile

    ; pull relevant settings variable from ini file
    hideWindow := IniRead(settingsFile, "Utility_Windows", hideWinKey, "false")

    ; only show window if marked to show in settings file
    if (hideWindow == "true"){
        ; dont show window (presumably becasue to use has set it to not show in the past)
    }
    else{
        ; create gui and add contents and controls
        msgWin_dsa := Gui(, winTitle)
        msgWin_dsa.Opt("+AlwaysOnTop -MinimizeBox")
        msgWin_dsa.MarginX := 10
        msgWin_dsa.MarginY := 10
        msgWin_dsa.BackColor := "ffffff"
        msgWin_dsa.SetFont("s9", "Segoe UI")
        msgWin_dsa.Add("Text", "w400 +Wrap", winMessage)

        checkBox := msgWin_dsa.Add("Checkbox",, "Don't show again")
        
        closeBtn := msgWin_dsa.Add("Button", "Default w80 h25 x170", btnText)
        closeBtn.OnEvent("Click", closeDSAWindow)

        msgWin_dsa.OnEvent("Close", closeDSAWindow)
        msgWin_dsa.Show("xCenter y150 w420")


        ; gui internal methods
        closeDSAWindow(*){
            if (checkBox.Value == 1){
                IniWrite("true", settingsFile, "Utility_Windows", hideWinKey)
            }
            msgWin_dsa.Destroy()
        }

    }

}