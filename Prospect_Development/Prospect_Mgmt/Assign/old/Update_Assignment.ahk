;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         A.N.Other <myemail@nowhere.com>
;
; Script Function:
; This script takes a row from a spreadsheet and uses it to create
; prospect assignment updates
; row fields
; ___ADV ID_____DO ID_____School Code_____Type Code_____Comment___
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2

dbName = ADVTRN

today()
{
    FormatTime, month,,MM
    FormatTime, day,,dd
    FormatTime, year,,yyyy
    today = %month%/%day%/%year%
    return today
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

CopyCell()
{
    Sleep, 10
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait, 0 ;
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

pullupEntity(entID)
{
    ;pull up entity window
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    Send, %entID%
    Send, {Enter}
}

nextColumn(nextCol = 1, prevCol = 0)
{
    diff := (nextCol - prevCol)
    ;MsgBox, %diff%
    Loop, %diff%
    {
        Send, {TAB}
        Sleep, 50
    }
    return
}

errorFound(errorFound = False)
{
    ;get window text to check Saving Prospect warning
    Sleep, 2250
    IfWinExist, Saving Prospect
    {
        errorFound := True
        ;Send, {Enter}
        return %errorFound%
    }
    return %errorFound%
}

getToProsA()
{
    global dbName
    ; this gives 'control' to the current window
    ; Send, {ENTER}
    ; Send, {CtrlDown}{F4}{CtrlUp}
    ; alt method below
    Send, {F7}
    Send, {ShiftDown}{TAB 4}
    Send, {ShiftUp}
    ;now we can tab to prospect window
    ;on the main bio view
    ;Send, {TAB 8}
    Send, {Enter}

    ;now pop up a dialog and exit
    ;allowing us to navigate this window
    Send, {F3}
    Send, {ESC}

    ;now tab to assignment box
    Send, {TAB 3}
    Send, {Enter}

    ;now would be a good time to capture
    ;prospect ID, if you're interested
    
    ;to prevent errors when accessing the same record repeatedly
    ;we go back and close the first prospect window
    Send, {ShiftDown}{F6}{ShiftUp}
    ;should be back to normal now
    
    ;open Assignment Window
    Send, {AltDown}A{AltUp}
    
    ;move control back to advance, jik
    MoveTo(dbName) 
    
    return
}

newAssignment(doID, schoolCode, typeCode, comment)
{
    ;this is just to ensure
    ;the cursor starts in field 1
    Send, {F7}
    Send, {F6}
    Sleep, 100
    Send, %doID%
    Send, {TAB 2}
    Send, %schoolCode%
    Send, {TAB}
    Send, %typeCode%
    Send, {TAB 4}
    today := today()
    Send, %today%
    Send, {TAB 2}
    Send, %comment%
    Send, {F8}
    Sleep, 100
    return
}

inactivateAssignment()
{
    Send, {F7}
    ;tab to active box
    Send, {TAB 4}
    Send, {SPACE}
    ;tab to stop date
    Send, {TAB 4}
    today := today()
    Send, %today%
    Send, {F8}
    Sleep, 100
}
#IfWinActive, Excel
+^G::
SendMode, Input

entityID := CopyCell()
nextColumn()
doID := CopyCell()
nextColumn()
schoolCode := CopyCell()
nextColumn()
typeCode := CopyCell()
nextColumn()
comment := CopyCell()

;go back to original field


moveTo(dbName)

pullupEntity(entityID)
; MsgBox, getting to pros a
getToProsA()
; MsgBox, adding new assignment
newAssignment(doID, schoolCode, typeCode, comment)
if errorFound()
{
    ;need to uncheck other active box
    Send, {ENTER}
    ;refresh
    SEnd, {F7}
    MsgBox, 4097, , Select the Row to Inactivate
        IfMsgBox, Ok
        {
            inactivateAssignment()
            newAssignment(doID, schoolCode, typeCode, comment)
        } else
        {
            exit
        }
}
MoveTo("Excel")
Send, {ENTER}
return