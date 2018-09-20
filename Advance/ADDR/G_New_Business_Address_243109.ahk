/*This script is designed to add new business contact
for Team Request 243109
*/

#A::
Reload
Return

#MaxMem 256 ; Allow 256 MB per variable

pullupEntity(entID)
{
    ;pull up entity window
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    Send, %entID%
    Send, {Enter}
}

callGoTo(tableText)
{
    SetTitleMatchMode, 2
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, %tableText%{ENTER} ;Call up given window
    Sleep, 250
    ;MsgBox, I've just entered the info
    MoveTo("Production Database") 
    return
}

callGoToBlank(tableText, entID)
{
    SetTitleMatchMode, 2
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, %tableText%
    Send, {TAB}
    Send, %entID%{ENTER} ;Call up given window
    Sleep, 250
    ;move control back to advance, jik
    MoveTo("Production Database") 
    
    
    return
}

CopyCell()
{
    Sleep, 10
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait, 0 ;
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

MoveTo(destination)
{
    ;Move to destination 
    ;MsgBox, Moving to %destination%
    IfWinNotActive, %destination%,
    ;MsgBox "2"
    ;#WinActivateForce
    WinActivate, %destination%, ,
    ;MsgBox "4"
    ;WinWaitActive, %destination%
    
}

nextColumn(nextCol, prevCol)
{
    diff := (nextCol - prevCol)
    ;MsgBox, %diff%
    Loop, %diff%
    {
        Send, {Tab}
    }
    return
}
checkEmpty(testVar, varName)
{
    ;check to see if the cell is empty
    IfEqual, testVar,
    {
        MsgBox, No %varName%
        Exit
    }
}
updateExcel()
{
    ;back to excel
    ;MoveTo("- Excel")
    Send, {AltDown}{TAB}{AltUp}
    MoveTo("Excel") 
    Sleep, 200
    ;MsgBox, Should be in Excel
    Send, {ENTER}
    Send, {CtrlDown}{Left}{CtrlUp}{Right}
    Send, {Up}
    Send, {Alt}HH{Right 7}{ENTER}
    Send, {Down}
    return
}
/*
#####################################
SCRIPT START
*/
+^g::

SetTitleMatchMode, 2
TeamID = 243109
FormatTime, month,,MM
FormatTime, day,,dd
FormatTime, year,,yyyy

;#####################################
;COLLECT DATA FROM EXCEL
;#####################################
MoveTo("Excel") 
;MsgBox afterwards
;###############################
; entity info
entityIDcol := 2


;entityID
;should already be selected
entityID := CopyCell()

;check to see if the row is empty/without primary donor id
checkEmpty(entityID, "Entity ID")

; Employment
empIDcol := 5
empNamecol := 6
empTitlecol := 7
workStreetcol := 8
workCitycol := 9
workStatecol := 10
workZipcol := 11
workCountrycol := 12
workTelcol := 13

;empID
nextColumn(empIDcol, entityIDcol)
empID := CopyCell()

;empName
nextColumn(empNamecol, empIDcol)
empName := CopyCell()

;empTitle
nextColumn(empTitlecol, empNamecol)
empTitle := CopyCell()

;workStreet
nextColumn(workStreetcol, empTitlecol)
workStreet := CopyCell()
checkEmpty(workStreet, "Street Address")

;workCity
nextColumn(workCitycol, workStreetcol)
workCity := CopyCell()

;workState
nextColumn(workStatecol, workCitycol)
workState := CopyCell()

;workZip
nextColumn(workZipcol, workStatecol)
workZip := CopyCell()

;workCountry
nextColumn(workCountrycol, workZipcol)
workCountry := CopyCell()

;workTel
nextColumn(workTelcol, workCountrycol)
workTel := CopyCell()

;########################################
;ENTER DATA INTO ADVANCE
;########################################

;move to Advance
MoveTo("Production Database")
Sleep, 200

;pullupEntity(entityID)

;##############################
; now for the mailing address
;##################
callGoToBlank("ADDR", entityID)
addrVar =
(
%empTitle%
%empName%
%workStreet%
%workCity%, %workState% %workZip%
)
justAddr = 
(
%workStreet%
%workCity%, %workState%
%workZip% %workCountry%
)
MsgBox, 8196, New Business Address, Does this Address entry exist?`n%addrVar%
IfMsgBox, Yes
{
    addrExist = True
} Else
{
    addrExist = False
}
IfEqual, addrExist, False
{
    ;make new address
    Send, {F6}
    ;change to Type B
    Send, {AltDown}Y{AltUp}
    Send, B
    ;now move back to zip, then up to line 1
    Send, {AltDown}Z{AltUp}
    Send, {ShiftDown}{Tab 6}{ShiftUp}
    ;now call address tool
    Send, {AltDown}33{AltUp}
    Sleep, 300
    ;pop up msgbox to show rest of address
    ;after msgbox is close, address tool should have already been exited
    message = 
    (
    Enter the address:

    %justAddr%

    Click OK afterwards
    )
    addressHelper()
}Else
{
    ;back to excel
    Sleep, 200
    updateExcel()
    Exit
}
return

addressHelper()
{
    global
    Gui, Add, Text, x2 y9 w230 h230 , %justAddr%
    Gui, Add, Button, x2 y239 w110 h60 gA_OK, OK
    Gui, Add, Button, x122 y239 w110 h60 gA_Cancel, Cancel
    ; Generated using SmartGUI Creator 4.0
    Gui, Show, x1269 y29 h302 w233, New GUI Window
    MoveTo("Production Database")
    MoveTo("Instant Address")
    return

    A_OK:
    {
        Gui, Cancel
        MoveTo("Production Database")
        ;MsgBox, Filling out Address
        fillOutAddress()
        return
    }
    A_Cancel:
    {
        Gui, Cancel
        MoveTo("Excel")
        send, {ENTER}
        return
    }
}

fillOutAddress()
{
    global
    ;address has been entered, now we fill other fields
    ;move to type field
    Send, {AltDown}Y{AltUp}
    ;now over to dates
    Send, {TAB 2}
    ;month, day year
    Send,  %month%
    Send, {TAB}
    Send, %day%
    Send, {TAB}
    Send, %year%
    ;to src link
    Send, {TAB 6}
    Send, BIO
    ;move to phone type
    Send, {TAB 3}
    reggie = [)(-\s"]+
    cutPhone := RegExReplace(workTel, reggie,"")
    ;msgBox, %cutPhone%
    StringLen, telLen, cutPhone
    If (telLen > 9)
    {
        Send, B
        ;move to phone number
        Send, {TAB}
        if workCountry = USA
        {
            ;two phone fields
            StringLeft, firstThree, cutPhone, 3
            StringTrimLeft, restOfNum, cutPhone, 3
            ;send first section (split on space)
            Send, %firstThree%
            Send, {Tab}
            ;send second section
            Send, %restOfNum%
        }
        else
        {
            Send, %workTel%
        }
    }else
    {
        Send, {TAB 2}
    }
    ;move to comment
    Send, {TAB 5}
    Send, Team %TeamID%
    ;link employer
    MsgBox, Link employer:`n`n%empTitle%`n%empName%

    MsgBox, 4100, Save New Entry, Ready to Save?
    IfMsgBox, Yes
    {
        ;MsgBox, Saving
        Send, {F8}
        Sleep, 200
        Send, {ENTER 2}
        ;MsgBox, We've saved, now let's update Excel
        updateExcel()
        ;MsgBox, Excel updated
        return
    }else
    {
        Exit
    }
}