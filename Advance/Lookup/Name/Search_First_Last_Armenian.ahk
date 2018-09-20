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

;##############################
;### Functions ################

CopyCell()
{
    Sleep, 10
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait, 
    cellOfInterest = %clipboard%
    ;clean up id from excel
    ;check if you are in excel so the newline can be dropped
    IfWinActive, Excel, , , 
    {
        ;clean up id from excel
        StringTrimRight, clipboard, clipboard, 2
        Sleep, 10
        cellOfInterest = %clipboard%
        Send, {Esc}
    }
    clipboard = ;
    return cellOfInterest
}
MoveTo(destination)
{
    ;Move to destination 
    ;MsgBox, Moving to %destination%
    IfWinNotActive, %destination%,
    ;MsgBox "2"
    ;#WinActivateForce
    WinActivate, %destination%, ,
    ;MsgBox "4"
    WinWaitActive, %destination%
    
}

Test_For_Nothing()
{
    noRecordsFound = False
    ;get window text to check No Entities warning
    Sleep, 750
    DetectHiddenText, On
    WinGetText, windowCheck 
    ;MsgBox, The text is:`n%windowCheck%
    ;MsgBox 1 idFound = %idFound%

    ;term to check for
    ;windowText = "No Entities"
    windowText = Looking
    ;Search for windowText in current windows
    IfInString, windowCheck, %windowText%
    {
        noRecordsFound = True
        ;MsgBox, Not Found
        ;Send, {Enter}
        return %noRecordsFound%
    }
    return %noRecordsFound%
}

;###########################
;### Script ################

#A::
Reload
return

;#IfWinActive, Excel
+^L::
Loop
{
; firstName = ;
; lastName = ;
firstName := CopyCell()

IfEqual, firstName, ""
{
    Return
}
;next column
Sleep, 10
Send, {TAB}

lastName := CopyCell()

MoveTo("Production Database")

;open search dialog, tab to last name field
Send, {ShiftDown}{F4}{ShiftUp}
Send, {TAB 3}

;enter the names to search on
Sleep, 25
Send, %lastName%
Send, {TAB}
Send, %firstName%

;execute query
Send, {Enter}
Send, Y

noResults := Test_For_Nothing()
;MsgBox, %noResults%
IfEqual, noResults, True
{
    ;close dialog
    ;close search box
    ;go back to excel
    Sleep, 25
    Send, {ENTER}
    Sleep, 100
    Send, {CtrlDown}{F4}{CtrlUp}
    MoveTo("Excel")
    ;color cell
    Send, {AltDown}HH
    Send, {Right 6}
    Send, {Enter}
    Sleep, 10
    ;next row
    Send, {DOWN}{LEFT}
}Else
{
    ; MoveTo("Excel")
    ; ;color cell
    ; Send, {AltDown}HH
    ; Send, {Right 7}
    ; Send, {Enter}
    ; Sleep, 10
    ; ;next row
    ; Send, {DOWN}
    ; Send, {LEFT}
    return
}
; MsgBox, %noResults%
; IfEqual, noResults, 1
; {
;     MsgBox, No Entities
; }
}
return

