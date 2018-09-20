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
    global doID
    global schoolCode
    global typeCode
    global comment
    global dbName

    Gui, Font, S24 CDefault Italic, Verdana
    Gui, Font, S24 CDefault Italic, Verdana
    Gui, Font, S18 CDefault Italic, Verdana
    Gui, Font, S16 CDefault Bold, Verdana
    Gui, Font, S14 CDefault Bold, Comic Sans MS
    Gui, Add, Button, x192 y329 w140 h70 gPauseButton, &Pause
    Gui, Add, Button, x240 y8 w57 h36 gResetButton, reset
    Gui, Add, Button,  x300 y8 w57 h36 gQuitButton,  quit
    Gui, Font, S14 CDefault Bold, Comic Sans MS
    Gui, Font, S14 CDefault Bold Italic, Comic Sans MS
    Gui, Font, S14 CDefault Bold, Comic Sans MS
    Gui, Font, S14 CGreen, Comic Sans MS
    Gui, Add, Button, x40 y329 w140 h70 , Start
    Gui, Font, S14 CGreen Bold, Comic San
    Gui, Font, S24 CGreen Italic, Comic San
    Gui, Add, Text, x10 y57 w340 h66 +Center, Assignment Updater
    Gui, Font, S16 CGreen Italic, Comic San
    Gui, Font, S22 CBlue, Constantia
    Gui, Add, Text, x15 y10 w107 h35 , Moira's
    Gui, Font, S8 CDefault, Verdana
    Gui, Add, Text, x22 y109 w100 h20 +Right, Database:
    Gui, Add, DropDownList, x132 y109 w100 h20 R2 vdbName, Production||Training 
    Gui, Add, Edit, x132 y139 w100 h20 vdoID,
    Gui, Add, Edit, x132 y169 w70 h20 vschoolCode,
    Gui, Add, Edit, x132 y199 w70 h20 vtypeCode,
    Gui, Add, Edit, x132 y239 w210 h80 vcomment,
    Gui, Add, Text, x22 y139 w100 h20 +Right, Assigned ID:
    Gui, Add, Text, x22 y169 w100 h20 +Right, School:
    Gui, Add, Text, x22 y199 w100 h20 +Right, Type:
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
        IfEqual, dbName, Production
        {
            dbName = Production Database
        }else
        {
            dbName = ADVTRN
        }
        IfWinNotExist, %dbName%
        {
            MsgBox, %dbName% isn't open.
            
        } else 
        {
            assignProspectDO(doID, schoolCode, typeCode, comment)
        }
        GuiControl, Enable, Start
        return
    }
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
    moveTo(dbName)
    Sleep, 1000
    IfWinExist, Saving Prospect
    {
        errorFound := True
        return %errorFound%
    }
    return %errorFound%
}

getProspectID()
{
    global dbName
    delay = 250
    ; this gives 'control' to the current window
    Sleep, %delay%
    Send, {F7}
    ;now we can tab to prospect window
    ;on the main bio view
    Send, {ShiftDown}{TAB 4}
    Send, {ShiftUp}
    Send, {Enter}
    ; Sleep, %delay%

    waitForWindow("Prospect Tracking Summary")
    
    ;now pop up a dialog and exit
    ;allowing us to navigate this window
    Send, {F3}
    Send, {ESC}
    Sleep, %delay%
    ;now tab to assignment box
    Send, {TAB 3}
    Send, {Enter}
    ; Sleep, %delay%
    prosWindow := waitForWindow("Prospect - (Prospect")
    ; msgBox, %prosWindow%
    IfEqual, prosWindow, 0
    {
        ; msgBox, proswindow is false
        return False
    }
    ;now would be a good time to capture
    ;prospect ID, if you're interested
    Send, {F3}
    prosID := CopyCell()
    ; msgBox, %prosID%
    ;close the ID window now
    Send, {ESC}
    ;chance to close other windows here
    Send, {CtrlDown}{F4}
    Send, {CtrlDown}{F4}
    Send, {CtrlDown}{F4}
    Send, {CtrlUp}

    return %prosID%

}

openProsWindow(prospectID, window)
{
    global dbName
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, %window%
    Send, {TAB}
    Send, {TAB}
    ; MsgBox, %prospectID%
    Send, %prospectID%
    Send, {ENTER} ;Call up Prospect window
    Sleep, 250
    ;MsgBox, I've just entered the info
    IfWinNotActive, %dbName%
    WinActivate, %dbName% 
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

waitForWindow(searchString, countmax = 50)
{
    ; msgBox, starting string search
    count = 0
    Loop
    {
        IfGreaterOrEqual, count, %countmax%
        {
            return False
        }
        DetectHiddenText, On
        WinGetText, windowCheck 
        ; MsgBox, The text is:`n%windowCheck%
        firstLine = ;
        ;get the first 
        Loop, parse, windowCheck, `n, `r
        {
            ; MsgBox, 4,,%A_LoopField%
            firstLine = %A_LoopField%
            break
        }
        IfInString, firstLine, %searchString%
        {
            ; msgBox, found %searchString%
            return True
        }
        Sleep, 10
        count += 1
    } 
    return True
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
    prosID := getProspectID()
    IfEqual, prosID, 0
    {
        MsgBox, No Prospect Record
        ;close window
        Send, {CtrlDown}{F4}{CtrlUp}
        GuiControl, Enable, Start
        Exit
    }
    ; msgBox, %prosID%
    ;open assignment and pros a windows
    Sleep, 100
    ; openProsWindow(prosID, "PROS")
    openProsWindow(prosID, "PROSA")
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
                GuiControl, Enable, Start
                exit
            }
    }
    MoveTo("Excel")
    Send, {ENTER}
    return
}
createGUI()
; ^+G::
; assignProspectDO()
; return