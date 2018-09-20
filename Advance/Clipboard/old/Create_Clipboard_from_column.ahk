/*
This AutoHotKey file will copy an entire column
in excel, and then add a new clipboard entry 
in Advance, for each row copied
Requires the clipboard to be open
You should also choose the appropriate clipboard type
*/
;To STart:
; Press Shift + Ctrl + G

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#MaxMem 256 ; Allow 256 MB per variable

dbName = Production Database 

CopyCol()
{
    Sleep, 10
    clipboard = ;
    Send, {ShiftDown}{CtrlDown}{Down}{ShiftUp}{CtrlUp}
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait, 0 ;
    cellOfInterest = %clipboard%
    /*
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
    */
    clipboard = ;
    Send, {ESC}
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
    ;WinWaitActive, %destination%
    
}

checkEmptyQuiet(testVar)
{
    ;check to see if the cell is empty
    IfEqual, testVar,
    {
        ;MsgBox, No %varName%
        return True
    }
    return False
}

;Assumes ADvance is open and clipboard is current window
;clipboard of choice (allocation, entity) should be selected

+^G::

SetTitleMatchMode, 2

MoveTo("Excel")
;correct cell should already have been selected

;collect entire column of ids
idList := CopyCol()

MoveTo(dbName)

; Read from the array:
Loop, parse, idList, `n,
{
    ; The following line uses the := operator to retrieve an array element:
    reggie = ["\s\t\r\n]+
    cleanID := RegExReplace(A_LoopField, reggie,"")
    ;msgBox, %cleanID%
    ;nextID := %A_LoopField%
    noID := checkEmptyQuiet(cleanID)
    if noID
    {
        Send, {TAB}
        Exit
    }else
    {
        Send, {F6}
        Sleep, 75
        Send, %cleanID%
    }
}
ExitApp