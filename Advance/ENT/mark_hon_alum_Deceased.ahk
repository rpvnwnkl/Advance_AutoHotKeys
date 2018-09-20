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
#MaxMem 256 ; Allow 256 MB per variable
dbName = ;
#Include %A_ScriptDir%\..\..\Library
#Include moveTo.ahk
#Include CopyCell.ahk
#Include callGoToBlank.ahk
#Include closeWindows.ahk
createGUI()
{
    global cbType
    global dbName
    global LoopCheck
    Gui, Font, S10 CBlue, Comic Sans MS
    ; Gui, Add, Text, x16 y56 w125 h31 , Clipboard Type:
    Gui, Font, S10, Verdana 
    ; Gui, Add, DropDownList, x142 y56 w140 h35 R11 vcbType AltSubmit, Allocation|Contact Report|Documents|Entity||Primary Gift|Matching Gift|Membership|Primary Pledge|Program Prospect|Proposal|Prospect
    Gui, Add, Text, x16 y87 w125 h31 , Database:
    Gui, Add, DropDownList, x142 y87 w140 h20 R2 vdbName, Production||Training 
    Gui, Add, Checkbox, x16 y117 w100 h31 vLoopCheck, Loop?
    Gui, Font, S8 CBlack, Helvetica

    Gui, Font, S12 CBlack Italic Bold, Comic Sans MS
    Gui, Add, Button, x22 y170 w272 h42 , Start
    Gui, Add, Button, x233 y6 w70 h37 gQuitButton, Quit
    Gui, Add, Button, x197 y6 w35 h37 gReloadButton, Re 
    Gui, Font, S18 CRed, Papyrus
    Gui, Add, Text, x8 y8 w195 h43 , Mark Deceased 
    ; Generated using SmartGUI Creator 4.0
    Gui, Show, x1163 y95 h230 w310, Mark Deceased 
    WinSet, AlwaysOnTop, On, Mark Deceased 
    ; MsgBox, it's working?
    return


    QuitButton:
        ExitApp, 

    ReloadButton:
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
        IfEqual, LoopCheck, 1
        {
            Loop
            {
                mark_honAlumDeceased()
            }
        } Else
        {
            mark_honAlumDeceased()
        }
        GuiControl, Enable, Start
        return
    }
}


mark_honAlumDeceased()
{
    global dbName
    IfWinNotExist, Excel
    {
        MsgBox, Excel isn't open.
        return
    } Else
    {
        MoveTo("Excel")
    }

    ;capture info from excel
    ;start on id cell
    entID := CopyCell()
   
    ;enter it into Advance

    IfWinNotExist, %dbName% 
    {
        MsgBox, %dbName% isn't open.
        return
    } Else
    {
        MoveTo(dbName)
    }

    ;open bdet
    callGoToBlank("BDET", entID)
    ;update bdet
    Send, {TAB 10}
    Send, IN
    ;save
    Send, {F8}
    ;popup dismissal
    ; WinActivate, Update
    WinWaitActive, Update Birth
    Send, {Enter}
    ;close
    closeWindows()

    ;open ent window
    callGoToBlank("ENT", entID)
    ;update ent window
    Send, {TAB 10}
    Send, D
    ;save
    Send, {F8}
    
    ;close
    closeWindows()

    ;go back to excel
    moveTo("Excel")
    ;color
    Send, {Home}
    Send, {AltDown}HH
    Send, {Right 7}
    Send, {AltUp}
    Send, {Enter 2}

    return
}

createGUI()