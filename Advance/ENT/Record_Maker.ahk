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
#Include callGoTo.ahk
#Include moveTo.ahk
#Include waitForWindow.ahk
#Include closeWindows.ahk
#Include CopyCell.ahk
#Include createInformal.ahk

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
    Gui, Add, Text, x8 y8 w195 h43 , Record Maker
    ; Generated using SmartGUI Creator 4.0
    Gui, Show, x1163 y95 h230 w310, Record Maker 
    WinSet, AlwaysOnTop, On, Record Maker 
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
                create_new_hon_alum_record()
            }
        } Else
        {
            create_new_hon_alum_record()
        }
        GuiControl, Enable, Start
        return
    }
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

createEntity(prefix, firstName, middleName, lastName, suffix, proSuffix, recYear, recSchool, recType, recStatus, comment)
{
    global dbName
    ;pull up new entity window
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    Send, {F6}

    ;###########################
    ;start with name
    Send, {AltDown}L{AltUp}
    Send, %lastName%
    Send, {Tab}
    ;####################
    ;DOES NOT ACCOUNT FOR INITIAL FIRST NAME
    Send, %firstName%
    Send, {Tab}
    Send, %middleName%
    Send, {Tab}
    Send, %prefix%
    Send, {Tab}
    Send, %suffix%
    Send, {Tab}
    Send, %proSuffix%

    ;############################
    ;Now record type info
    ;year, school + type
    Send, {Tab 2}
    Send, %recYear%
    Send, {Tab}
    Send, %recSchool%
    Send, {Tab}
    Send, %recType%
    Send, {Tab}
    Send, %recStatus%

    ;###############################
    ;add comment
    Send, {AltDown}m{AltUp}
    Send, %comment%
    
    ;########################
    ;Now, save and review
    Send, {F8}

    IfNotEqual, proSuffix,
    {
        ;need to ok on validate entity popup
        WinWaitActive, Validate Entity
        Send, {Enter}
        ; Msgbox, just closed validate box?
    }

    ;check for duplicate entities window
    moveTo(dbName)
    dupeExists := waitForWindow("Duplicate Entities", 25)
    ;prompt to save anyway/cancel
    IfEqual, dupeExists, 1
    {
        MsgBox, 4097, ,Possible Dupe found, Save Anyway?
            IfMsgBox, Ok
            {
                ;save
                moveTo(dbName)
                ;close dupe box
                Send, {Enter}
                ;submit on save anyway box
                Send, {Enter}

                ;check for validate window again
                IfNotEqual, proSuffix,
                {
                    ;need to ok on validate entity popup
                    WinWaitActive, Validate Entity
                    Send, {Enter}
                    ; Msgbox, just closed validate box?
                }
                ;now continue to save regular stuff
            } else
            {
                ;don't save
                moveTo(dbName)
                Send, {Enter}
                Send, N
                closeWindows()
                return 0
            }
    } Else
    {
        ; Msgbox, No dupe found
    }
    
    WinWaitActive, Primary Entity
    IfWinActive, Primary Entity
    {
        Send, {Enter}
    }
    
    ;now no address box
    IfWinActive, Primary Entity
    {
        Send, {Enter}
    }

    ;retrieve ID
    moveTo(dbName)
    entityReady := waitForWindow("Entity - (")
    ; msgBox, %entityReady%
    IfEqual, entityReady, 1
    {
        ;gathering new ID number
        ;get window text to get ID
        WinGetText, entityWindow 
        Sleep, 200
        ;while we could find the position of the id num
        ;we know where it is for now, so we will hard code it
        StringMid, idNum, entityWindow, 15, 6
        return %idNum%
    } Else
    {
        MsgBox, check for errors, Id not captured
        return 0
    }
}


create_new_hon_alum_record()
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
    global prefix
    global firstName
    global prevFirstName
    global middleName
    global prevMiddleName
    global lastName
    global prevLastName
    global suffix
    global proSuffix
    global recYear
    global recSchool
    global recType
    global recStatus
    global comment
    ; captureInfo()
    ;start on blank id cell
    Send, {TAB}
    prefix := CopyCell()
    IfEqual, prefix,
    {
        msgBox, NO prefix, end of rows?
        Exit
    }
    Send, {TAB}
    firstName := CopyCell()
    Send, {TAB}
    middleName := CopyCell()
    Send, {TAB}
    lastName := CopyCell()
    Send, {TAB}
    suffix := CopyCell()
    Send, {TAB}
    proSuffix := CopyCell()
    Send, {TAB}
    recYear := CopyCell()
    ; Send, {TAB}
    recSchool := "UN"
    recType := "AL"
    recStatus := "L"
    comment := "Team 250521"
    ; MsgBox, %prefix%`n%firstName%`n%middleName%`n%lastName%`n%suffix%`n%proSuffix%`n%recYear%`n%recSchool%`n%recType%`n%recStatus%`n%comment%
    ;enter it into Advance

    IfEqual, firstName, %prevFirstName%
    {
        IfEqual, middleName, %prevMiddleName%
        {
            IfEqual, lastName, %prevLastName%
            {
                MsgBox, Duplicate person row.
                GuiControl, Enable, Start
                exit
            }
        }
    }
    IfWinNotExist, %dbName% 
    {
        MsgBox, %dbName% isn't open.
        return
    } Else
    {
        MoveTo(dbName)
    }

    newID := createEntity(prefix, firstName, middleName, lastName, suffix, proSuffix, recYear, recSchool, recType, recStatus, comment)
    ; msgBox, %newID%
    ;create salutation
    createInformal(firstName)
    ;create degree

    ;close window
    moveTo(dbName)
    Send, {CtrlDown}{F4}{CtrlUp}
    Send, N
    moveTo(dbName)
    Sleep, 25
    Send, {CtrlDown}{F4}{CtrlUp}
    ;go back to excel adn drop new id
    moveTo("Excel")
    ; Send, {Left 7}
    Send, {Home}
    Send, %newID%
    ; Send, {Right}
    ; Send, {Left}
    ; pastedID := CopyCell()
    Send, {Enter}

    ;assign lastName values
    prevFirstName := firstName
    prevLastName := lastName
    prevMiddleName := middleName
    ; WinActivate, Record Maker
    return
}

createGUI()