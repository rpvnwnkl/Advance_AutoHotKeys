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
#Include ..\..\Library
#Include callGoTo.ahk
#Include moveTo.ahk
#Include CopyCell.ahk
#Include windowPresent.ahk
#Include waitForWindow.ahk
#Include closeWindows.ahk

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

pullupEntity(entID)
{
    ;pull up entity window
    SendInput, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    SendInput, %entID%
    SendInput, {Enter}
}

inactivateAssignment(stopDate, comment)
{
    global dbName
    moveTo(dbName)
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
    SendInput, %stopDate%
    SendInput, {TAB}
    ;add comment
    IfNotEqual, comment,
    {
        SendInput, %comment%;%A_Space%
    }
    SendInput, {F8}
    Sleep, 100
    return
}

resetGUI()
{
    GuiControl, Enable, Start
    Exit
}
safeClick(buttonName, windowName, windowText, callBack, easyWindowName)
{
    ;buttonName, windowName, windowText are all as described in ControlClick docs
    ;easyWindowName is used for display message upon failure
    ;callback can execute custom behavior on failure (ie, it calls a function)
    ;set callback to 0 to default as return

    moveTo(windowName)
    delay = 75
    windowFound := waitForWindow(windowText, delay)
    IfEqual, windowFound, 1
    {
        Sleep, %delay%
        ControlClick, %buttonName%, %windowName%, %windowText%,,2,NA
    } Else
    {
        Msgbox, Could not open %easyWindowName%
        IfNotEqual, callBack, 0
        {
            %callBack%()
        } Else
        {
            return
        }
        
    }
}
openProsA()
{
    global dbName
    ;click prospect window from bio overview
    ; Send, {F7}
    safeClick("Button1", dbName, "Custom Tab Profile", "resetGUI", "Prospect Window")
    ;click Prospect INformation button from tracking summary view
    ; Send, {F7}
    safeClick("Button5", dbName, "Prospect Tracking Summary", "resetGUI", "Prospect Tracking Summary Window")
    ;click Assignment button from prospect window
    ; this reload seems to help maintain window control
    Send, {F7}
    safeClick("Button12", dbName, "Prospect - (", "resetGUI", "Prospect Window")
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
    MoveTo("Excel")
    entityID := CopyCell()

    ;open entity record
    moveTo(dbName)
    pullupEntity(entityID)

    ; check for alert
    Sleep, 150
    alertFound := windowPresent("Alerts -")
    ; msgBox, %alertFound%
    IfEqual, alertFound, 1
    {
        closeWindows()
    }
    moveTo(dbName)
    ; click on window buttons to open prosA window
    openProsA()

    ;select row to deactivate
    MsgBox, 4097, , Select the Row to Inactivate
        IfMsgBox, Ok
        {
            inactivateAssignment(stopDate, comment)
        } else
        {
            return
        }
    ;move back to excel
    MoveTo("Excel")
    SendInput, {ENTER}
    return
}

createGUI()
