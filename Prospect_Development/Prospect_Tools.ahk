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

dbName := "Production Database"

#Include %A_ScriptDir%\..\Library
#Include callGoTo.ahk
#Include moveTo.ahk

;Open ID in Advance
#Include Open_Adv_ID.ahk
MButton::
#O::
{
    BlockInput, On
    Open_ADV_ID()
    BlockInput, Off
}
return

;open prospect assignment window
#Include openProsA.ahk
#Include waitForWindow.ahk
#Include checkAlert.ahk
^MButton::
#A::
{
    BlockInput, On
    Open_ADV_ID()
    checkAlert()
    openProsA()
    BlockInput, Off
    return
}

return
;close current window with ESC key
#IfWinActive, Production Database 
Esc::
Send, {CtrlDown}{F4}{CtrlUp}
return

;open new entity search box and tab to last name
; #IfWinActive, ADVTRN 
; ^N::
; Send, {ShiftDown}{F4}{ShiftUp}
; Send, {TAB 3}
; return