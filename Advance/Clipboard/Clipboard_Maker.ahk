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

createGUI()
{
    global cbType
    global dbName
    ;stop reload buttons
    Gui, Font, S14 CDefault , Wingdings 3 
    Gui, Add, Button, x225 y3 w40 h30 gResetButton, Q
    Gui, Font, S14 CDefault, Wingdings
    Gui, Add, Button, x265 y3 w40 h30 gQuitButton, x
    
    Gui, Font, S10 CBlue, Comic Sans MS
    Gui, Add, Text, x16 y56 w125 h31 , Clipboard Type:
    Gui, Font, S10, Verdana 
    Gui, Add, DropDownList, x142 y56 w140 h35 R11 vcbType AltSubmit, Allocation|Contact Report|Documents|Entity||Primary Gift|Matching Gift|Membership|Primary Pledge|Program Prospect|Proposal|Prospect
    Gui, Add, Text, x16 y87 w125 h31 , Database:
    Gui, Add, DropDownList, x142 y87 w140 h20 R2 vdbName, Production||Training 
    Gui, Font, S8 CBlack, Helvetica
    Gui, Add, Text, x22 y119 w270 h40 , 
    (
Select the clipboard and database above. 
In Excel, select the cells containing IDs. 
Click "Start" to copy the IDs over to Advance Clipboard.
    )
    Gui, Font, S12 CBlack Italic Bold, Comic Sans MS
    Gui, Add, Button, x22 y170 w272 h42 , Start
    ; Gui, Add, Button, x233 y6 w70 h37 gQuitButton, Quit
    Gui, Font, S18 CRed, Papyrus
    Gui, Add, Text, x8 y8 w200 h43 , Clipboard Maker
    ; Generated using SmartGUI Creator 4.0
    Gui, Show, x1163 y95 h230 w310, Clipboard Maker 
    WinSet, AlwaysOnTop, On, Clipboard Maker 
    ; MsgBox, it's working?
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
        create_clipboard_from_selection(cbType)
        GuiControl, Enable, Start
        return
    }
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

CopyCol()
{
    BlockInput, On
    Sleep, 10
    clipboard = ;
    ; SendInput, {ShiftDown}{CtrlDown}{Down}{ShiftUp}{CtrlUp}
    SendInput, {CTRLDOWN}c{CTRLUP}
    ClipWait, 1
    cellOfInterest = %clipboard%
    clipboard = ;
    SendInput, {ESC}
    BlockInput, Off
    return cellOfInterest
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

create_clipboard_from_selection(clipboardType)
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

    ;correct cell should already have been selected
    ; MsgBox, %clipboardType%
    ;collect entire column of ids
    idList := CopyCol()

    IfWinNotExist, %dbName% 
    {
        MsgBox, %dbName% isn't open.
        return
    } Else
    {
        MoveTo(dbName)
    }

    ;pull up the clipboard
    ;clear it
    ;close it
    ;open it again
    ;(a fresh start)
    ;open clipboard
    ControlSend,,{F4},%dbName%,,,
    Sleep, 50
    ;clear clipboards
    ControlSend,,{ShiftDown}{CtrlDown}A{CtrlUp}{ShiftUp}, %dbName%
    Sleep, 50
    ;in case, close the window
    ControlSend,,{CtrlDown}{F4}{CtrlUp},%dbName%,,,
    Sleep, 50
    ;open clipboard
    ControlSend,,{F4},%dbName%,,,
    Sleep, 50
    ;now we select the clipboard chosen by user
    ControlSend,PBTreeView32_1001, {Down %clipboardType%},%dbName%
    Sleep, 50
    ; tab tab to do a context switch which clears leftover cells
    ControlSend,,{Tab 2}, %dbName%
    Sleep, 25
    
    ; Read from the array:
    Loop, parse, idList, `n,
    {
        BlockInput, On
        ; The following line uses the := operator to retrieve an array element:
        reggie = ["\s\t\r\n]+
        cleanID := RegExReplace(A_LoopField, reggie,"")
        noID := checkEmptyQuiet(cleanID)
        if noID
        {
            SendInput, {TAB}
        }else
        {
            MoveTo(dbName)
            SendInput, {F6}
            SendInput, %cleanID%
            SendInput, {Enter 2}
            Sleep, 10
        }
        BlockInput, Off
    }
    return
}

createGUI()