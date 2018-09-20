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

task_helper_gui()
{
    global dbName
    global prosComment


    Gui, Font, S14 CDefault , Wingdings 3 
    Gui, Add, Button, x175 y3 w40 h30 gResetButton, Q
    Gui, Font, S14 CDefault, Wingdings
    Gui, Add, Button, x216 y3 w40 h30 gQuitButton, x
    Gui, Font, S10 CDefault, courier new
    Gui, Add, DropDownList, x122 y59 w120 h30 R2 vdbName, Production||Training
    Gui, Font, S12 CDefault Bold Italic, Consolas
    Gui, Add, Text, x12 y59 w100 h30 , Database:
    Gui, Add, Text, x12 y109 w100 h30 , Comment:
    Gui, Font, S10 CDefault norm, courier new
    Gui, Add, Edit, x12 y139 w230 h100 vprosComment,
    Gui, Font, S14 CDefault Bold Italic, consolas 
    Gui, Add, Button, x43 y250 w166 h62 Center, Start
    Gui, Font, S14 CBlue Bold Italic, Showcard Gothic
    Gui, Add, Text, x8 y7 w142 h20 , New Parent 
    Gui, Add, Text, x8 y33 w142 h20 , Task Helper
    ; Generated using SmartGUI Creator 4.0
    Gui, Show, x1195 y108 h327 w260, Task Helper
    WinSet, AlwaysOnTop, On, Task Helper
    Return

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
        IfWinNotExist, %dbName%
        {
            MsgBox, %dbName% isn't open.
            
        } else 
        {
           completeProspectTask() 
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
#Include moveTo.ahk
#Include callGoTo.ahk

CopyCell()
{
    Sleep, 10
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait, 1
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

getResearchRating()
{
    ;get research rating
    callGoTo("ENTE")
    ;check type field for CR value
    Loop, 10
    {
        currVal := copyCell()
        IfEqual, currVal, CR
        {
            ; msgbox, found CR rating
            Send, {TAB}
            ratingVal := CopyCell()
            ; msgBox, 1%ratingVal%1
            IfEqual, ratingVal,
            {
                MsgBox, No Rating Score Given
                GuiControl, Enable, Start
                Exit
            }
            ; msgbox, rating equals %ratingVal%
            ;close window
            closeWindows()
            return %ratingVal%
        }Else
        {
            Send, {TAB 13}
        }
    }
    MsgBox, Couldn't find Research Rating 
    GuiControl, Enable, Start
    Exit

    ;tab to field, copy to value
    ;dummy val below
    ; ratingVal = 07
    ; return %ratingVal%
}
#Include closeWindows.ahk
getSchoolCode()
{
    callGoTo("REC")
    Loop, 10
    {
        currVal := copyCell()
        IfEqual, currVal, PA
        {
            ;found parent row
            Send, {TAB 2}
            schoolCode := CopyCell()
            closeWindows()
            return %schoolCode%
        }Else IfEqual, currVal, GP
        {
            ;found grad parent row
            Send, {TAB 2}
            schoolCode := CopyCell()
            closeWindows()
            return %schoolCode%
        }Else
        {
            Send, {TAB 3}
        }
    }
    MsgBox, Couldn't find Parent Record Type
    GuiControl, Enable, Start
    Exit
    ; Send, {ENTER 2}
    ; schoolCode := CopyCell()
    ; closeWindows()
    ; return %schoolCode%
}
createNewProspect()
{
    global prosType := "I"
    global prosSchool 
    global prosRating
    global prosRatingStatus = A
    global prosComment

    callGoTo("PROS")
    Send, {TAB 2}
    Send, %prosType%
    Send, {TAB}
    prosSchool := getSchoolCode()
    Send, %prosSchool%
    Send, {TAB 3}
    prosRating := getResearchRating()
    Send, %prosRating%
    Send, {TAB}
    Send, A
    Send, {TAB}
    today := today()
    Send, %today%
    Send, {TAB 5}
    SEnd, {Space}
    Send, {AltDown}U
    Send, {AltUp}
    Send, {TAB 2}
    ;need Moira to enter comment here?
    Send, %prosComment%
    Send, {F8}
    closeWindows()
}

updateProspectTask()
{
    ;we'll assume the window is how we left it
    ;to be sure, we'll reset 
    Send, {F7}
    Send, {TAB}
    SEnd, 2
    Send, {TAB 11}
    today := today()
    Send, %today%
    Send, {F8}
    closeWindows()
}

getID()
{
    ;pull up id window
    Send, {F3}
    Sleep, 250
    businessID := CopyCell()
    ;close the Id window
    Send, {Esc}
    return %businessID%
}

checkDesc(keyword)
{
    ;get to description field
    Send, {AltDown}D
    Send, {AltUp}
    Send, {AltDown}EL
    Send, {AltUp}
    fieldText := CopyCell()
    IfInString, fieldText, %keyword%
    {
        return True
    }
    return False
}

#Include windowPresent.ahk
completeProspectTask()
{
    global dbName
    global prosType
    global prosSchool
    global prosRating
    global prosRatingStatus

    moveTo(dbName)
    If windowPresent("Task Lookup")
    {
        ;assume that first row is active, ie can be entered on
        Send, {ENTER}
        sleep, 10
    }
    If windowPresent("Tasks -")
    {
        ;reset to ensure we are on first field
        Send, {F7}
        ;check task type
        taskType := CopyCell()
        IfNotEqual, taskType, ID
        {
            MsgBox, Not Prospect Identification
            return
            ;close window
            ;next row
        } Else
        {
            ;check description for parent mention
            parentInDesc := checkDesc("parent")
            ; MsgBox, %parentInDesc%
            IfNotEqual, parentInDesc, 1
            {
                MsgBox, Parent not mentioned in task description
                return
            }
            entID := getID()
            
            ;create prospect record
            createNewProspect()
            updateProspectTask()
            Send, {TAB 2}
            Send, {Down}
            ; pullupEntity(entID)
            ; callGoTo("PROSV")
        }
    } Else
    {
        MsgBox, Not in Task Viewer
    }
}

task_helper_gui()
; ^+G::
; assignProspectDO()
; return