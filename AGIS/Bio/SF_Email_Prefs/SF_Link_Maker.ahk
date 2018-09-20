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
dbName = Production Database
#Include ..\..\..\Library
#Include moveTo.ahk
#Include CopyCell.ahk
#Include callGoTo.ahk
#Include ensureAdvance.ahk
#Include closeWindows.ahk

createGUI()
{
    global cbType
    global dbName
    global linkDisplay
    global newLink
    ;stop reload buttons
    Gui, Font, S14 CDefault , Wingdings 3 
    Gui, Add, Button, x425 y3 w40 h30 gResetButton, Q
    Gui, Font, S14 CDefault, Wingdings
    Gui, Add, Button, x465 y3 w40 h30 gQuitButton, x
    
    Gui, Font, S10 CBlack, Comic Sans MS
    Gui, Add, Text, x16 y56 w155 h31 , Email Preferences Link:
    Gui, Font, s8 CBlue Underline,
    ; Gui, Add, Text, x16 y87 w125 h31 grunLink, %newLink% 
    Gui, Add, Text, x16 y87 w500 h23 grunLink vlinkDisplay, %newLink% 
;     Gui, Font, S8 CBlack, Helvetica
;     Gui, Add, Text, x22 y119 w270 h40 , 
;     (
; Select the clipboard and database above. 
; In Excel, select the cells containing IDs. 
; Click "Start" to copy the IDs over to Advance Clipboard.
;     )
    Gui, Font, S12 CBlack Norm Italic Bold, Comic Sans MS
    Gui, Add, Button, x50 y120 w422 h42 , Start
    ; Gui, Add, Button, x233 y6 w70 h37 gQuitButton, Quit
    Gui, Font, S18 CGreen, Papyrus
    Gui, Add, Text, x8 y8 w415 h43 , SalesForce Email Pref. Link Maker
    ; Generated using SmartGUI Creator 4.0
    Gui, Show, x975 y95 h175 w515, SF Link Maker 
    WinSet, AlwaysOnTop, On, SF Link Maker 
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
        newLink := createLink()
        GuiControl,, linkDisplay, %newLink%
        ; postLink(newLink)
        GuiControl, Enable, Start
        return
    }
    runLink:
    {
        run, %newLink%
        return
    }
}

getSfID()
{
    BlockInput, On
    ensureAdvance()
    callGoTo("ID")
    currVal = ;
    prevVal = ;
    Loop, 10
    {
        currVal := copyCell()
        IfEqual, currVal, SFC 
        {
            ;found Salesforce row
            Send, {TAB}
            sfID := CopyCell()
            closeWindows()
            BlockInput, Off
            return %sfID%
        } Else
        {
            IfEqual, currVal, %prevVal%
            {
                ;no id
                BlockInput, Off
                MsgBox, No ID found
                GuiControl, Enable, Start
                Exit
            } Else
            {
                prevVal = %currVal%
                Send, {Enter 3}
                Send, {Tab}
            }
        }
    }
}

createLink()
{
    newLink := ""
    GuiControl,, linkDisplay, %newLink%
    sfID := getSfID()
    begLink := "http://pages.e.tufts.edu/preferences/?s="
    endLink := "&m=0000000&j=0000000&"
    newLink = %begLink%%sfID%%endLink%
    ; GuiControl,, linkDisplay, %newLink%
    return %newLink%
}

; postLink(newLink)
; {
;     global linkDisplay
;     GuiControl,, vlinkDisplay, %newLink%
;     ; MsgBox, %newLink%
;     return
; }


createGUI()