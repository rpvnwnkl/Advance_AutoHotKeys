;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         A.N.Other <myemail@nowhere.com>
;
; Script Function:
;	Template script (you can customize this template by editing "ShellNew\Template.ahk" in your Windows folder)
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2
#Include, ..\Library
;set DB Name (ADVTRN or Production Database)
dbName = ADVTRN

;Open ID in Advance
#Include, Open_ADV_ID.ahk
MButton::
{
    ; Run, Open_ADV_ID_live.ahk, Open_\
    MsgBox, 1
    Open_ADV_ID()
    return
}

;Open Prospect ID in Advance
^MButton::
Run, Open_Prospect_live.ahk, Open_\
return

;Open Prospect Assignment window in Advance
#Include, getProspectID.ahk
#Include, openProsWindow.ahk
+^MButton::
{
    ; Run, Open_Prospect_from_ID_live.ahk, Open_\
    MsgBox, 2
    
    prospectID := getProspectID()
    MsgBox, 3
    
    openProsWindow(prospectID, "PROSA")
    return
}




;close current window with ESC key
#IfWinActive, Production Database
Esc::
Send, {CtrlDown}{F4}{CtrlUp}
return
