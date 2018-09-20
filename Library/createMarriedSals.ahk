;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         A.N.Other <myemail@nowhere.com>
;
; Script Function:
;	Template script (you can customize this template by editing "ShellNew\Template.ahk" in your Windows folder)
;

; #Include %A_ScriptDir%
#Include callGoTo.ahk
#Include CopyCell.ahk
; #Include moveTo.ahk
#Include windowPresent.ahk

createMarriedSals()
{
    global dbName
    ; SetTitleMatchMode, 2
    If WinActive(dbName)
    {
         ;######################
         ;MUST BE IN MAR WINDOW
         ;########################
        If windowPresent("Marital Information")
        {
            callGoTo("SAL")
            Send, {SHIFTDOWN}{F3}{SHIFTUP} ; go to spouse 2 record
            spouse2 := CopyCell()
            Send, {SHIFTDOWN}{F3}{SHIFTUP} ; go to spouse 1 record
            spouse1 := CopyCell()
            formalSal = %spouse1% and %spouse2%
            Send, {F6}
            Send, %formalSal%
            Send, {TAB  2}
            Send, ji
            Send, {F8}
            Sleep, 100
            Send, {SHIFTDOWN}{F3}{SHIFTUP}
            Send, {F6}
            Send, %formalSal%
            Send, {TAB  2}
            Send, ji
            Send, {F8}
            Sleep, 100
            Send, {CTRLDOWN}{F4}{CTRLUP} ;Remove or add the semi-colon before CTRLDOWN to have the SAL window close itself
        }
    }
    return
}