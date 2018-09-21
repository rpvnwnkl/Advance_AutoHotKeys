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
#Include nextColumn.ahk
#Include moveTo.ahk
#Include CopyCell.ahk
#Include windowPresent.ahk
#Include checkAlert.ahk
#Include waitForWindow.ahk
#Include closeWindows.ahk

createGUI()
{
    global advIDcolumn
    global giftRatioColumn
    global giftAmtColumn 
    global comment
    global dbName
    global loopCheck

    FormatTime, thisYear, YYYY, yyyy

    Gui, Font, S14 CDefault , Wingdings 3 
    Gui, Add, Button, x240 y3 w40 h30 gResetButton, Q
    Gui, Font, S14 CDefault, Wingdings
    Gui, Add, Button, x281 y3 w40 h30 gQuitButton, x
    
    Gui, Font, S22 CRed, Papyrus
    Gui, Add, Text, x15 y10 w150 h35, Art Sale 
    Gui, Font, S24 CBlack Bold Italic, Comic Sans MS
    Gui, Add, Text, x19 y57 w340 h66, GIK Entry
    
    Gui, Font, S8 CBlack, Verdana
    Gui, Add, DropDownList, x220 y40 w100 h20 R2 vdbName, Production||Training 
    Gui, Add, Checkbox, x225 y65 w100 h20 -wrap vloopCheck, Loop?  
    
    Gui, Font, S8 CGreen, Verdana
    Gui, Add, Text, x22 y110 w100 h20 , Field:
    Gui, Add, Text, x132 y110 w100 h20, Column #:

    Gui, Font, CBlack Norm,
    Gui, Add, Text, x32 y130 w100 h20 , Advance ID:
    Gui, Add, Edit, x142 y130 w70 h20 vadvIDcolumn, 1

    Gui, Add, Text, x32 y160 w100 h20 , Gift Ratio:
    Gui, Add, Edit, x142 y160 w70 h20 vgiftRatioColumn, 2
    
    Gui, Add, Text, x32 y190 w100 h20 , Gift Amount:
    Gui, Add, Edit, x142 y190 w70 h20 vgiftAmtColumn, 3
    
    Gui, Add, Text, x32 y220 w200 h20 , GIK Reciept Description:
    Gui, Add, Edit, x32 y239 w265 h80 vcomment, `n[gift ratio]`% of the value of artwork sold during the SMFA at Tufts Art Sale, Nov. %thisYear%. 
    
    Gui, Font, S14 CGreen, Comic Sans MS
    Gui, Add, Button, x15 y329 w140 h70 , Start
    Gui, Add, Button, x165 y329 w140 h70 , Stop
    ; Generated using SmartGUI Creator 4.0
    Gui, Show, x1083 y59 h415 w325, GIK Entry
    WinSet, AlwaysOnTop, On, GIK Entry 
    ; MsgBox, it's working?
    return


    QuitButton:
        ExitApp, 

    ResetButton:
        Reload

    ButtonStop:
    {
        GuiControl, Disable, Stop
        ; GuiControl, Enable, Start
        Pause, On
        return
    }

    ButtonStart:
    {
        ; Pause, Off, 1
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
            IfEqual, loopCheck, 0
            {
                enterGIK_ArtSale(advIDcolumn, giftRatioColumn, giftAmtColumn, comment)
            } Else
            {
                Loop, 
                {
                    stillValid := enterGIK_ArtSale(advIDcolumn, giftRatioColumn, giftAmtColumn, comment)
                    IfEqual, stillValid, 0
                    {
                        break
                    } 
                    IfEqual, A_IsPaused, 1
                    {
                        break
                    }
                }
            }
        }
        GuiControl, Enable, Start
        return
    }
}

createNewGIK(advID, giftAmount, giftRatio, comment)
{
    global dbName

    delay = 10
    BlockInput, On
   
    ;open new transaction
    Send, {F6}
    Sleep, %delay%

    ;Enter ID Num
    Send, %advID%
    Send, {Tab}
    Sleep, %delay%

    ; check for alert
    checkAlert(100)

    ; close donor info window
    moveTo("Gift Ledger") 
    Send, O
    moveTo(dbName)

    ;Enter total gift amount
    Send, %giftAmount%
    Sleep, %delay%

    ; Save
    Send, {F8}
    Sleep, 20*%delay%
    ; catch a missing required value error
    IfWinExist, Entry Ledger
    {
        WinActivate, Entry Ledger
        WinWaitActive, Entry Ledger
        Send, {Enter}
        BlockInput, Off
        MsgBox, Defaults for the batch were not set.
        GuiControl, Enable, Start
        Exit
    }

    ; Open tender window
    Send, {AltDown}{AltUp}
    Send, E
    Send, T
    tenderOpen := waitForWindow("Tender - (", 65)    
    IfEqual, tenderOpen, 0
    {
        BlockInput, Off
        MsgBox, Tender window couldn't open.
        GuiControl, Enable, Start 
        Exit
    }

    ; enter GIK type
    Send, {TAB 7}
    Sleep, %delay%
    Send, A
    Sleep, %delay%

    ; enter description
    Send, {TAB}
    Sleep, %delay%
    ; enter a newline for printer preferences
    ; Send, {Enter}
    StringReplace, commentString, comment, [gift ratio], %giftRatio%, All
    Send, %commentString%
    Sleep, %delay%

    ; save and exit tender window
    Send, {F8}
    ; Sleep, 20*%delay%
    closeWindows()
    waitForWindow("Batch Gift Ledger")
    return
}

enterGIK_ArtSale(advIDcolumn, giftRatioColumn, giftAmtColumn, comment)
{
    global dbName

    ; go to excel
    BlockInput, On
    IfWinExist, Excel
    {
        moveTo("Excel")
    } Else
    {
        MsgBox, Excel isn't open.
        GuiControl, Enable, Start
        Exit
    }

    ;###################################
    ;#### Copy data from Excel #########

    ;go to beginning of row 
    Send, {HOME}
    ; get advance ID
    nextColumn(advIDcolumn)
    advID := CopyCell()
    IfEqual, advID,
    {
        ; no ID
        BlockInput, Off
        MsgBox, No ID captured.
        return 0
    }
    ; get gift ratio
    nextColumn(giftRatioColumn, advIDcolumn)
    giftRatio := CopyCell()
    ; get gift total
    nextColumn(giftAmtColumn, giftRatioColumn)
    giftAmount := CopyCell()

    ;####################################
    ;##### Enter data into advance ######

    ; go to advance
    IfWinExist, %dbName% 
    {
        WinActivate, %dbName%
        WinWaitActive, %dbName%
    } Else
    {
        BlockInput, Off
        MsgBox, Advance isn't open.
        GuiControl, Enable, Start 
        Exit
    }
    ; ensure batch gift ledger window is active
    Sleep, 15*%delay%
    ledgerOpen := windowPresent("Batch Gift Ledger")
    IfEqual, ledgerOpen, 0
    {
        BlockInput, Off
        MsgBox, Open the Batch Gift Ledger and try again.
        GuiControl, Enable, Start
        Exit
    } 
    ; enter new gift
    createNewGIK(advID, giftAmount, giftRatio, comment)
    ; are there any other popups to look for?

    ; update excel
    moveTo("Excel")
    ; need to go to beginning of row before coloring
    Send, {Home}
    Send, {AltDown}{AltUp}
    Send, H
    Send, H
    Send, {Right 5}
    Send, {Enter}

    ; prepare for next row
    Send, {Enter}
    BlockInput, Off

    return 1
}



createGUI()