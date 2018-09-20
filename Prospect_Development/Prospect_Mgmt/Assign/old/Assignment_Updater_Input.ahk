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


Gui, Font, S24 CDefault Italic, Verdana
Gui, Font, S24 CDefault Italic, Verdana
Gui, Font, S18 CDefault Italic, Verdana
Gui, Font, S16 CDefault Bold, Verdana
Gui, Font, S14 CDefault Bold, Comic Sans MS
Gui, Add, Button, x192 y329 w160 h70 gPauseButton, &Pause
Gui, Add, Button, x228 y7 w62 h36 gResetButton, reset
Gui, Add, Button,  x295 y8 w57 h36 gQuitButton,  quit
Gui, Font, S14 CDefault Bold, Comic Sans MS
Gui, Font, S14 CDefault Bold Italic, Comic Sans MS
Gui, Font, S14 CDefault Bold, Comic Sans MS
Gui, Font, S14 CGreen, Comic Sans MS
Gui, Add, Button, x22 y329 w140 h70 , Start
Gui, Font, S14 CGreen Bold, Comic San
Gui, Font, S24 CGreen Italic, Comic San
Gui, Add, Text, x10 y57 w340 h66 +Center, Assignment Updater
Gui, Font, S16 CGreen Italic, Comic San
Gui, Font, S22 CBlue, Constantia
Gui, Add, Text, x15 y10 w107 h35 , Moira's
Gui, Font, S8 CDefault, Verdana
Gui, Add, Edit, x192 y139 w100 h20 vdoID,
Gui, Add, Edit, x192 y169 w70 h20 vschoolCode,
Gui, Add, Edit, x192 y199 w70 h20 vtypeCode,
Gui, Add, Edit, x132 y239 w210 h80 vcomment,
Gui, Add, Text, x22 y139 w100 h20 +Right, DO ID:
Gui, Add, Text, x22 y169 w100 h20 +Right, School Code:
Gui, Add, Text, x22 y199 w100 h20 +Right, DO Type Code:
Gui, Add, Text, x22 y239 w100 h20 +Right , Comment:
; Generated using SmartGUI Creator 4.0
Gui, Show, x1083 y59 h415 w369, assignmentGUI
WinSet, AlwaysOnTop, On, assignmentGUI 
; MsgBox, it's working?
return

QuitButton:
    ExitApp, 

ResetButton:
    Reload

PauseButton:
{
    Pause, Toggle, 1
    return
}

ButtonStart:
{
    Gui, submit, NoHide
    ;gray out button
    GuiControl, Disable, Start
    assignProspectDO(doID, schoolCode, typeCode, comment)
    GuiControl, Enable, Start
    return
}

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
    global dbName
    ;get window text to check Saving Prospect warning
    moveTo(dbName)
    Sleep, 2500
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
    delay = 1000
    ; this gives 'control' to the current window
    ; Send, {ENTER}
    ; Send, {CtrlDown}{F4}{CtrlUp}
    ; alt method below
    Sleep, %delay%
    Send, {F7}
    Send, {ShiftDown}{TAB 4}
    Send, {ShiftUp}
    ;now we can tab to prospect window
    ;on the main bio view
    ;Send, {TAB 8}
    Send, {Enter}
    Sleep, %delay%

    ;now pop up a dialog and exit
    ;allowing us to navigate this window
    Send, {F3}
    Send, {ESC}
    Sleep, %delay%
    ;now tab to assignment box
    Send, {TAB 3}
    Send, {Enter}
    Sleep, %delay%
    ;now would be a good time to capture
    ;prospect ID, if you're interested
    
    ;to prevent errors when accessing the same record repeatedly
    ;we go back and close the first prospect window
    Send, {CtrlDown}{F6}{CtrlUp}
    Sleep, %delay%
    Send, {CtrlDown}{F4}{CtrlUp}
    Sleep, %delay%
    Send, {CtrlDown}{F6}{CtrlUp}

    ;should be back to normal now
    Sleep, %delay%
    ;open Assignment Window
    Send, {AltDown}A
    Send, {AltUp}
    Sleep, %delay%
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
assignProspectDO(doID, schoolCode, typeCode, comment) 
{
    global dbName
    SendMode, Input
    IfWinNotExist, Excel 
    {
        MsgBox, Excel isn't open
        Exit
    }
    MoveTo("Excel")
    entityID := CopyCell()
    ; nextColumn()
    ; doID := CopyCell()
    ; nextColumn()
    ; schoolCode := CopyCell()
    ; nextColumn()
    ; typeCode := CopyCell()
    ; nextColumn()
    ; comment := CopyCell()
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
}

; ^+G::
; assignProspectDO()
; return