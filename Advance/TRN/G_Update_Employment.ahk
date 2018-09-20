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
    MoveTo("ADVTRN") 
    return
}

CopyCell()
{
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait ;
    ;clean up id from excel
    ;check if you are in excel so the newline can be dropped
    IfWinActive, Excel, , , 
    {
        ;clean up id from excel
        StringTrimRight, clipboard, clipboard, 2
        Send, {Esc}
    }
    cellOfInterest = %clipboard%
    clipboard = ;
    return cellOfInterest
}

MoveTo(destination)
{
    ;Move to destination 
    ;MsgBox "1"
    IfWinNotActive, %destination%,
    ;MsgBox "2"
    WinActivate, %destination%, ,
    ;MsgBox "4"
    
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

;workCity (col 9)
Send, {Tab 1}
workCity := CopyCell()

;workState (col 10)
Send, {Tab 1}
workState := CopyCell()

;workZip (col 11)
Send, {Tab 1}
workZip := CopyCell()

;########################################
;ENTER DATA INTO ADVANCE
;########################################

;move to Advance
MoveTo("ADVTRN")
Sleep, 200

pullupEntity(entityID)

;callVar = EMPT
callGoTo("EMPT")

;MsgBox, %empName%`n%empTitle%
;#SingleInstance
;SetTimer, ChangeButtonNames, 50
MsgBox, 4, New Employment, Does this employment entry exist?`n`n%empName%`n%empTitle%
IfMsgBox, Yes
    empExist = True
Else
    empExist = False

;if not empExist, add new
IfEqual, empExist, False
{
    MsgBox, 4, Inactivate Row, Should any employment rows be Inactivated?
    IfMsgBox, Yes
    {
        MsgBox, 1, Select the Row to Inactivate
        IfMsgBox, Ok
        {
            ;this moves the cursor to the employer id field
            Send, {TAB}
            currCell := CopyCell()
            StringLen, empIdLength, currCell
            ;this checks to see if employer id is blank
            ;if so, we have more fields to tab through
            if empIdLength = 0
            {
                Send, {TAB 2}
                Msgbox, big id
            }
            Send, {TAB}
            Send, FE

            ;stop day, month, year
            Send, {TAB 12}
            ;Send, %month%
            Send, {TAB}
            ;Send, %day%
            Send, {TAB}
            ;Send, %year%
            ;primary uncheck
            Send, {TAB}
            Send, {SPACE}
            MsgBox, Check the Match Box as it hasn't been updated.
            Send, {F8}

            /*MsgBox, 4, Save New Entry, Ready to Save?
                IfMsgBox, Yes
                {
                    Send, {F8}
                }else
                {
                    Send, {F7}
                }
            */
        }
    }
    ;done inactivating current employment
    Sleep, 200
    ;MsgBox, starting new entry
    ;now, add new employment
    Send, {F6}
    ;MsgBox, %empID%
    If empID > 0
    {
        Send, %empID%
        Send, {TAB}
    } else
    {
        Send, {TAB}
        ;MsgBox, Pasting employer name: `n%empName%
        Send, %empName%
        Send, {TAB 2}
    }
    /*;search for employer record
    searchForEmployer(empName)
    MsgBox, 4,, Find Employer?
        IfMsgBox, No
        {
            Send, {ESC 2}
            ;move to Emp Name 1 field
            Send, {TAB}
            Send, %empName%
            Send, {TAB 2}
        } else
        {
            Send, {ENTER}
            Send, {TAB}
        }
    */
    ;move to Job Type
    Send, PE
    ;move to Job Status
    Send, {TAB}
    Send, F
    ;move to Job Title
    Send, {TAB}
    Send, %empTitle%
    ;move to Data Src
    Send, {TAB 6}
    Send, BIO
    ;move to month, day year
    Send, {TAB}
    Send, %month%
    Send, {TAB}
    send, %day%
    Send, {TAB}
    Send, %year%
    ;move to preferred box
    Send, {TAB 4}
    Send, {SPACE}
    ;move to Fld of Work
    Send, {TAB 3}
    Send, FAAA
    ;mvoe to comment
    Send, {TAB 5}
    Send, Team 242124

    MsgBox, 4, Save New Entry, Ready to Save?`nCheck for Employer record, and matching status.
    IfMsgBox, Yes
    {
        Send, {F8}
    }else
    {
        ;Send, {F7}
        Exit
    }
}

/*
;##############################
; now for the mailing address
;##################
callVar = ADDR
callGoTo(callVar)
addrVar =
(
%workAddr%
%workCity%, %workState% %workZip%
)
MsgBox, 4, New Business Address, Create this business address?`n%addrVar%
IfMsgBox, No
    addrExist = True
Else
    addrExist = False

IfEqual, addrExist, False
{
    ;make new address
    Send, {F6}
    ;first line addr
    Send, %workAddr%
    ;addr City
    MsgBox, hold
    Send, {TAB 4}
    Send, %workCity%
    ;addr state
    Send, {TAB}
    Send, %workState%
    ;addr zip
    Send, {TAB}
    Send, %workZip%
    ;two tabs will set zip adn move past in case there is drop down
    ;but now we aren't sure where we are
    Send, {TAB 2}
    ;try to move to addr type
    ;you are either there, or after it
    Send, {TAB 5}
    ;MsgBox, about to copy currCell
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait ;
    currCell = %clipboard%
    MsgBox, %currCell%
    IfEqual, currCell, H
    {
        Send, {DELETE}B
    } else
    {
        Msgbox, error with zip code, check fields, will pick up at type
    }
    ;Send, B
    ;skip status
    Send, {TAB}
    ;month, day year
    Send, {TAB}
    Send,  %month%
    Send, {TAB}
    Send, %day%
    Send, {TAB}
    Send, %year%
    ;to empt link
    Send, {TAB 5}
    MsgBox, about to pick drop down
    Send, {DOWN}
    ;diversion caused by drop down, kicks back to status
    Send, {ENTER 8}
    ;now send Down again
    Send, {DOWN}
    ;move back to src field
    Send, {TAB 9}
    Send, BIO
    ;move to comment
    Send, {TAB 10}
    Send, Team 242124

    MsgBox, 4, Save New Entry, Ready to Save?
    IfMsgBox, Yes
    {
        Send, {F8}
    }else
    {
        Send, {F7}
    }
}
*/
;back to excel
MoveTo("Excel")
Send, {ENTER}
return
/*ChangeButtonNames:
IfWinNotExist, New Employment
    return ; keep waiting
SetTimer, ChangeButtonNames, off
WinActivate
ControlSetText, Button1, &Yes
ControlSetText, Button2, &No
return
*/