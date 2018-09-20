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
#Include ..\..\..\Library
#Include pullUpEntity.ahk
#Include callGoTo.ahk
#Include moveTo.ahk
#Include CopyCell.ahk
#Include windowPresent.ahk
#Include waitForWindow.ahk
#Include closeWindows.ahk
#Include safeClick.ahk
#Include openProsA.ahk
#Include checkAlert.ahk

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
    return


    QuitButton:
        ExitApp, 

    ResetButton:
        Reload

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
        IfWinExist, %dbName%
        {
            inactivateProspectDO(stopDate, comment)
        } Else
        {
            MsgBox, The Database isn't open.
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

resetGUI()
{
    GuiControl, Enable, Start
    Exit
}

inactivateAssignment(stopDate, comment)
{
    ;given a stopdate and comment, with a row selected
    ;in the ProsA window, this function inactivateds the 
    ;current row, adding the stopdate and prepending the comment

    global dbName
    BlockInput, On
    ensureAdvance()
    ;this is just to ensure
    ;the cursor starts in field 1
    SendInput, {F7}
    ;gain control over assignment window
    safeClick("FNUDO3120208", dbName, "Prospect Assignments", "resetGUI", "Assignments Window")
    ;now inactivate
    SendInput, {TAB 4}
    ;uncheck active box
    SendInput, {SPACE}
    SendInput, {TAB 4}
    ;enter stop date
    ;need to account for fille din stop date.
    Send, {CtrlDown}{ShiftDown}{Right}{ShiftUp}{CtrlUp}
    SendInput, %stopDate%
    SendInput, {TAB}
    ;add comment
    IfNotEqual, comment,
    {
        SendInput, %comment%;%A_Space%
    }
    SendInput, {F8}
    BlockInput, Off
    Sleep, 100
    return
}
closeBackgroundWindow(num)
{
    BlockInput, On
    ensureAdvance()
    Loop, %num%
    {
        Send, {AltDown}FU
        Send, {AltUp}
    }
    closeWindows(num)
    BlockInput, Off
    return
}
inactivateProspectDO(stopDate, comment) 
{
    global dbName
    SendMode, Input

    ;get entity ID
    IfWinNotExist, Excel 
    {
        MsgBox, Excel isn't open
        resetGUI()
    }
    BlockInput, On
    MoveTo("Excel")
    entityID := CopyCell()

    ;open entity record
    ; moveTo(dbName)
    ; ensureAdvance()
    pullupEntity(entityID)

    ; check for alert
    checkAlert()
    ensureAdvance()
    ; click on window buttons to open prosA window
    openProsA()
    BlockInput, Off
    ;select row to deactivate
    MsgBox, 4097, , Select the Row to Inactivate
        IfMsgBox, Ok
        {
            inactivateAssignment(stopDate, comment)
        } else
        {
            return
        }
    ;close ProsA background Windows
    closeBackgroundWindow("3")

    
    ;move back to excel
    BlockInput, On
    MoveTo("Excel")
    SendInput, {ENTER}
    BlockInput, Off
    return
}

createGUI()
