;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         A.N.Other <myemail@nowhere.com>
;
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
    global primaryIdNum
    global donorIdColumn 
    global spouseIdColumn 
    global fullGiftAmtColumn
    global legalGiftAmtColumn
    global runnerIdColumn
    global anonFlagColumn
    ; global comment
    global dbName
    global loopCheck

    FormatTime, thisYear, YYYY, yyyy

    Gui, Font, S14 CDefault , Wingdings 3 
    Gui, Add, Button, x240 y3 w40 h30 gResetButton, Q
    Gui, Font, S14 CDefault, Wingdings
    Gui, Add, Button, x281 y3 w40 h30 gQuitButton, x
    
    Gui, Font, S18 CTeal, Consolas
    Gui, Add, Text, x15 y10 w40 h35, PMC 
    Gui, Font, S24 CBlack Bold Italic, Comic Sans MS
    Gui, Add, Text, x55 y10 w175 h66, Gift Entry
    
    Gui, Font, S8 CPurple, Verdana
    Gui, Add, Text, x190 y75 w85 h20 , DB:
    Gui, Add, DropDownList, x220 y74 w100 h20 R2 vdbName, Production||Training 
    
    Gui, Font, S9 Norm CMaroon Bold, Verdana
    Gui, Add, Text, x10 y75 w85 h20 , Primary ID:
    Gui, Font, CBlack
    Gui, Add, Edit, x95 y75 w75 h20 , 293573
    
    Gui, Font, S8 Norm CRed Bold Italic, Verdana
    Gui, Add, Text, x22 y110 w100 h20 , Field:
    Gui, Add, Text, x162 y110 w100 h20, Column #:

    Gui, Font, CBlack Norm,
    Gui, Add, Text, x32 y130 w120 h20 , Donor ID:
    Gui, Add, Edit, x172 y130 w70 h20 vdonorIdColumn, 2

    Gui, Add, Text, x32 y160 w120 h20 , Spouse ID:
    Gui, Add, Edit, x172 y160 w70 h20 vspouseIdColumn, 4
    
    Gui, Add, Text, x32 y190 w120 h20 , Donation Amount:
    Gui, Add, Edit, x172 y190 w70 h20 vfullGiftAmtColumn, 6
    
    Gui, Add, Text, x32 y220 w120 h20 , Net Amount:
    Gui, Add, Edit, x172 y220 w70 h20 vlegalGiftAmtColumn, 7
    
    Gui, Add, Text, x32 y250 w120 h20 , Runner ID:
    Gui, Add, Edit, x172 y250 w70 h20 vrunnerIdColumn, 20
    
    Gui, Add, Text, x32 y280 w120 h20 , Anon Flag:
    Gui, Add, Edit, x172 y280 w70 h20 vanonFlagColumn, 22
    
    Gui, Font, S10 CBlue Bold Italic, Verdana
    Gui, Add, Checkbox, x42 y305 w100 h20 vloopCheck, Loop?  
    
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
        Pause, On, 1
        GuiControl, Disable, Stop
        GuiControl, Enable, Start
        return
    }

    ButtonStart:
    {
        Pause, Off
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
                enterPMC_Gift(primaryIdNum 
                              ,donorIdColumn 
                              ,spouseIdColumn 
                              ,fullGiftAmtColumn
                              ,legalGiftAmtColumn
                              ,runnerIdColumn
                              ,anonFlagColumn)
            } Else
            {
                Loop, 
                {
                    stillValid := enterPMC_Gift(primaryIdNum 
                                                ,donorIdColumn 
                                                ,spouseIdColumn 
                                                ,fullGiftAmtColumn
                                                ,legalGiftAmtColumn
                                                ,runnerIdColumn
                                                ,anonFlagColumn)
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
    StringReplace, commentString, comment, [gift ratio], %giftRatio%+5, All
    Send, %commentString%
    Sleep, %delay%

    ; save and exit tender window
    Send, {F8}
    ; Sleep, 20*%delay%
    closeWindows()
    waitForWindow("Batch Gift Ledger")
    return
}

enterPMC_Gift(primaryIdNum, donorIdColumn, spouseIdColumn, fullGiftAmtColumn
              , legalGiftAmtColumn, runnerIdColumn, anonFlagColumn)
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
    nextColumn(donorIdcolumn)
    advID := CopyCell()
    IfEqual, advID,
    {
        ; no ID
        BlockInput, Off
        MsgBox, No ID captured.
        return 0
    }
    ; get spouse ID
    nextColumn(spouseIdColumn, donorIdColumn)
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