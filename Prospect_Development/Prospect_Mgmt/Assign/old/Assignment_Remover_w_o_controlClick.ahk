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

; dbName = ADVTRN

doID = ;
schoolCode = ;
typeCode = ;
comment = ;

createGUI()
{
    global stopDate
    global comment
    global dbName

    today := today()
    Gui, Font, S14 CDefault , Wingdings 3 
    Gui, Add, Button, x185 y3 w40 h30 gResetButton, Q
    Gui, Font, S14 CDefault, Wingdings
    Gui, Add, Button, x226 y3 w40 h30 gQuitButton, x
    ; Gui, Font, S14 CDefault Italic Bold, Comic Sans MS
    ; Gui, Add, Button, x192 y255 w140 h70 gPauseButton, &Pause
    ; Gui, Add, Button, x240 y8 w57 h36 gResetButton, reset
    ; Gui, Add, Button,  x300 y8 w57 h36 gQuitButton,  quit
    Gui, Font, S14 CGreen Bold, Comic Sans MS
    Gui, Add, Button, x50 y255 w175 h70 , Start
    Gui, Font, S18 CRed Italic Bold, Comic San
    Gui, Add, Text, x5 y45 w265 h66 +Center, Assignment Remover 
    Gui, Font, S22 CBlue Bold, Constantia
    Gui, Add, Text, x15 y10 w107 h35 , Moira's
    Gui, Font, S8 CDefault, Verdana
    Gui, Add, Edit, x100 y135 w100 h20 vstopDate, %today%
    Gui, Add, Edit, x100 y165 w165 h80 vcomment,
    Gui, Add, DropDownList, x100 y105 w100 h20 R2 vdbName, Production||Training 
    Gui, Add, Text, x5 y105 w75 h20 +Right, Database:
    Gui, Add, Text, x5 y135 w75 h20 +Right, Stop Date:
    Gui, Add, Text, x5 y165 w75 h20 +Right , Comment:
    ; Generated using SmartGUI Creator 4.0
    Gui, Show, x1083 y59 h341 w275, Assignment Remover
    WinSet, AlwaysOnTop, On, Assignment Remover 
    ; MsgBox, it's working?
    return


    QuitButton:
        ExitApp, 

    ResetButton:
        Reload

    ; PauseButton:
    ; {
    ;     Pause, Toggle, 1
    ;     return
    ; }

    ButtonStart:
    {
        Gui, submit, NoHide
        ;gray out button
        GuiControl, Disable, Start
        IfEqual, dbName, Production
        {
            dbName = Production Database
        }else
        {
            dbName = ADVTRN
        }
        inactivateProspectDO(stopDate, comment)
        GuiControl, Enable, Start
        return
    }
}
createGUI()
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
    SendInput, {CTRLDOWN}c{CTRLUP}
    ClipWait, ;
    cellOfInterest = %clipboard%
    ;clean up id from excel
    ;check if you are in excel so the newline can be dropped
    IfWinActive, Excel, , , 
    {
        ;clean up id from excel
        StringTrimRight, clipboard, clipboard, 2
        Sleep, 10
        cellOfInterest = %clipboard%
        SendInput, {Esc}
    }
    clipboard = ;
    return cellOfInterest
}

pullupEntity(entID)
{
    ;pull up entity window
    SendInput, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    SendInput, %entID%
    SendInput, {Enter}
}

nextColumn(nextCol = 1, prevCol = 0)
{
    diff := (nextCol - prevCol)
    ;MsgBox, %diff%
    Loop, %diff%
    {
        SendInput, {TAB}
        Sleep, 50
    }
    return
}

errorFound(errorFound = False)
{
    global dbName
    ;get window text to check Saving Prospect warning
    moveTo(dbName)
    Sleep, 1000
    IfWinExist, Saving Prospect
    {
        errorFound := True
        ;SendInput, {Enter}
        return %errorFound%
    }
    return %errorFound%
}

getProspectID()
{
    global dbName
    delay = 250
    ; this gives 'control' to the current window
    ; SendInput, {ENTER}
    ; SendInput, {CtrlDown}{F4}{CtrlUp}
    ; alt method below
    Sleep, %delay%
    SendInput, {F7}
    SendInput, {ShiftDown}{TAB 4}
    SendInput, {ShiftUp}
    ;now we can tab to prospect window
    ;on the main bio view
    ;SendInput, {TAB 8}
    SendInput, {Enter}
    Sleep, %delay%

    ;now pop up a dialog and exit
    ;allowing us to navigate this window
    SendInput, {F3}
    SendInput, {ESC}
    Sleep, %delay%
    ;now tab to assignment box
    SendInput, {TAB 3}
    SendInput, {Enter}
    Sleep, %delay%
    ;now would be a good time to capture
    ;prospect ID, if you're interested
    SendInput, {F3}
    prosID := CopyCell()
    ; msgBox, %prosID%
    ;close the ID window now
    SendInput, {ESC}
    ;chance to close other windows here
    SendInput, {CtrlDown}{F4}
    SendInput, {CtrlDown}{F4}
    SendInput, {CtrlDown}{F4}
    SendInput, {CtrlUp}

    return %prosID%
}

openProsWindow(prospectID, window)
{
    global dbName
    SendInput, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 

    SendInput, %window%
    SendInput, {TAB}
    SendInput, {TAB}
    ; MsgBox, %prospectID%
    SendInput, %prospectID%
    SendInput, {ENTER} ;Call up Prospect window
    Sleep, 250
    ;MsgBox, I've just entered the info
    IfWinNotActive, %dbName%
    WinActivate, %dbName% 
    return
}

inactivateAssignment(stopDate, comment)
{
    ;this is just to ensure
    ;the cursor starts in field 1
    SendInput, {F7}
    ;now inactivate
    SendInput, {TAB 4}
    ;uncheck active box
    SendInput, {SPACE}
    SendInput, {TAB 4}
    ;enter stop date
    SendInput, %stopDate%
    SendInput, {TAB}
    ;add comment
    SendInput, %comment%;%A_Space%
    SendInput, {F8}
    Sleep, 100
    return
}

inactivateProspectDO(stopDate, comment) 
{
    global dbName
    SendMode, Input
    IfWinNotExist, Excel 
    {
        MsgBox, Excel isn't open
        GuiControl, Enable, Start
        Exit
    }
    MoveTo("Excel")
    entityID := CopyCell()

    moveTo(dbName)
    pullupEntity(entityID)
    ; MsgBox, getting to pros a
    prosID := getProspectID()
    ;open assignment and pros a windows
    Sleep, 100
    openProsWindow(prosID, "PROS")
    openProsWindow(prosID, "PROSA")
    ;select row to deactivate
    MsgBox, 4097, , Select the Row to Inactivate
        IfMsgBox, Ok
        {
            inactivateAssignment(stopDate, comment)
        } else
        {
            return
        }
    MoveTo("Excel")
    SendInput, {ENTER}
    return
}

; ^+G::
; assignProspectDO()
; return