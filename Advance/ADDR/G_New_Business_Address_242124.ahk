/*This script is designed to enter new parents for 
medical students 
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
    ;MsgBox "1"
    ;IfWinNotActive, %destination%,
    ;MsgBox "2"
    WinActivate, %destination%, ,
    ;MsgBox "4"
    ;WinWaitActive, %destination%
    
}
/*
#####################################
SCRIPT START
*/
+^g::
SetTitleMatchMode, 2

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
entityID = ;
entityFirst = ;
entityLast = ;

;entityID
;capture entityID (col 1)
;should already be selected
entityID := CopyCell()
;MsgBox, entityID = %entityID%

;check to see if the row is empty/without primary donor id
IfEqual, entityID, ""
{
    MsgBox, No Entity ID
    Exit
}

;entityFirst
;capture entityFirst (col 2)
Send, {Tab 1}
entityFirst := CopyCell()

;entityLast
;capture entityLast (col 3)
Send, {Tab 1}
entityLast := CopyCell()

;################################
; Employment
gradYear = ;
empName = ;
empTitle = ;
workAddr = ;
workCity = ;
workState = ;
workZip = ;

;gradYear is not used, so we will skip it
;(col 4)
Send, {Tab 1}

;empID (col 5)
Send, {Tab 1}
empID := CopyCell()

;empName (col 6)
Send, {Tab 1}
empName := CopyCell()

;empTitle (col 7)
Send, {Tab 1}
empTitle := CopyCell()

;workAddr (col 8)
Send, {Tab 1}
workAddr := CopyCell()
;check to see if the cell is empty
IfEqual, workAddr,
{
    MsgBox, No Street Address
    Exit
}

;workCity (col 9)
Send, {Tab 1}
workCity := CopyCell()
;check to see if the cell is empty
IfEqual, workCity,
{
    MsgBox, No City Given
    Exit
}

;workState (col 10)
Send, {Tab 1}
workState := CopyCell()
;check to see if the cell is empty
IfEqual, workState,
{
    MsgBox, No State Given
    Exit
}

;workZip (col 11)
Send, {Tab 1}
workZip := CopyCell()
;check to see if the cell is empty
IfEqual, workZip,
{
    MsgBox, No Zip Given
    Exit
}

;########################################
;ENTER DATA INTO ADVANCE
;########################################

;move to Advance
MoveTo("Production Database")
Sleep, 200

pullupEntity(entityID)

;##############################
; now for the mailing address
;##################
callGoTo("ADDR")
addrVar =
(
%empTitle%
%empName%
%workAddr%
%workCity%, %workState% %workZip%
)
justAddr = 
(
%workAddr%
%workCity%, %workState%
%workZip%
)
MsgBox, 8196, New Business Address, Does this Address entry exist?`n%addrVar%
IfMsgBox, Yes
    addrExist = True
Else
    addrExist = False

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
    ;MsgBox, 8192, Enter the address:`n%justAddr%
    ;ToolTip, Enter the address:`n%justAddr%, -0, -0
    message = 
    (
    Enter the address:
    %justAddr%

    Click OK afterwards
    )
    Gui, Add, Text, x2 y9 w230 h230 , %justAddr%
    Gui, Add, Button, x2 y239 w110 h60 gA_OK, OK
    Gui, Add, Button, x122 y239 w110 h60 gA_Cancel, Cancel
    ; Generated using SmartGUI Creator 4.0
    Gui, Show, x1269 y29 h302 w233, New GUI Window
    MoveTo("Production Database")
    MoveTo("Instant Address")
    return

    ;MoveTo("Production Database")
    ;MoveTo("Instant Address")
    ;hijinks to select field
    ;Send, {TAB}
    ;Send, {ShiftDown}{Tab}{ShiftUp}
    ;send zip
    ;Send, %workZip%
    ;Send, {Enter}

    A_OK:
    {
        Gui, Cancel
        MoveTo("Production Database")
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
        ;move to comment
        Send, {TAB 10}
        Send, Team 242124
        ;link employer
        MsgBox, Link employer:`n`n%empTitle%`n%empName%

        MsgBox, 4100, Save New Entry, Ready to Save?
        IfMsgBox, Yes
        {
            Send, {F8}
            Send, {ENTER 2}
            WinActivate, Excel, ,
            ;MsgBox, 8092, , Finished?
            ;back to excel            
            /*
            Sleep, 100
            MoveTo("Excel")
            ;WinActivate, Excel, ,
            Sleep, 100
            Send, {ENTER}
            Send, {Up}
            Send, {Alt}HH{Right 7}{ENTER}
            Send, {Down}
            */
            ;Exit
        }else
        {
            Exit
        }
    }
    A_Cancel:
    {
        Gui, Cancel
        MoveTo("Excel")
        send, {ENTER}
        Exit
    }
    ;back to excel
    MoveTo("Excel")
    Send, {ENTER}
    Send, {Up}
    Send, {Alt}HH{Right 7}{ENTER}
    Send, {Down}
    return
}Else
{
    ;back to excel
    MoveTo("Excel")
    Send, {ENTER}
    Send, {Up}
    Send, {Alt}HH{Right 7}{ENTER}
    Send, {Down}
    return
}
MoveTo("Excel")

return
