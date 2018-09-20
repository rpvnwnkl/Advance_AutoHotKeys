/*This script is designed to update employment info
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
    IfWinNotActive, %destination%,
    ;MsgBox "2"
    WinActivate, %destination%, ,
    ;MsgBox "4"
    
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
/*
#####################################
SCRIPT START
*/
+^g::
SetTitleMatchMode, 2
;Team Request numer
TeamID = 242729
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
entityIDcol := 1

;entityID
;capture entityID (col 1)
;should already be selected
entityID := CopyCell()

;check to see if the row is empty/without primary donor id
IfEqual, entityID, ""
{
    MsgBox, No Entity ID
    Exit
}

;################################
; Employment
empIDcol := 5
empNamecol := 6
empTitlecol := 7

;empID (col 5)
nextColumn(empIDcol, entityIDcol)
empID := CopyCell()

;empName (col 6)
nextColumn(empNamecol, empIDcol)
empName := CopyCell()

;empTitle (col 7)
nextColumn(empTitlecol, empNamecol)
empTitle := CopyCell()

;########################################
;ENTER DATA INTO ADVANCE
;########################################

;move to Advance
MoveTo("Production Database")
Sleep, 200

pullupEntity(entityID)

callGoTo("EMPT")

;MsgBox, %empName%`n%empTitle%

MsgBox, 8196, New Employment, Does this employment entry exist?`n`n%empName%`n%empTitle%
IfMsgBox, Yes
    empExist = True
Else
    empExist = False

;if not empExist, add new
IfEqual, empExist, False
{
    MsgBox, 4100, Inactivate Row, Should any employment rows be Inactivated?
    IfMsgBox, Yes
    {
        MsgBox, 4097, , Select the Row to Inactivate
        IfMsgBox, Ok
        {
            ;##############################            
            ;this either selects id field 
            ;or it shoots past to Job Type or emp name 1
            Send, {TAB}
            ;copy the selected contents
            currCell := CopyCell()
            ;check the copied lengths
            StringLen, empIdLength, currCell
            if empIdLength = 0
            {
                ;no employer id, at field 1
                ;MsgBox, No Employer ID
                Send, {TAB 3}
            }
            if empIdLength > 2
            {
                /*
                this part is necessary because advance highlights fields differently
                depending on whether the window loads with the row selected
                or whether you chose the current row seelction
                */
                ;at field 1 or 2
                ;need to determine which
                StringLeft, firstVal, currCell, 1
                IfEqual, firstVal, 0
                {
                    ;MsgBox, Employer ID Present
                    Send, {TAB}
                } else
                {
                    ;MsgBox, In Field 2
                    Send, {Tab 2}
                }
            }
            if empIdLength = 2
            {
                ;employer id, but too far
                ;MsgBox, Emp exists, on Job Type
                ;no movement
            }
            ;now we should be at job type (or therabouts)
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
            MsgBox, 8192, ,Check the Match Box as it hasn't been updated.
            Send, {F8}
            ;sending enter in case of popups
            Send, {ENTER 2}
            ;f7 to resent cursor to first field
            Send, {F7}
        }
    }
    ;done inactivating current employment
    Sleep, 200
    MsgBox, starting new entry
    ;now, add new employment
    Send, {F6}
    StringLen, empIdLength, empID
    If empIdLength > 1
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
    Send, Team %TeamID%

    MsgBox, 8196, Save New Entry, Ready to Save?`nCheck for Employer record, and matching status.
    IfMsgBox, Yes
    {
        MoveTo("Production Database")
        Sleep, 200
        Send, {F8}
        ;enter for popup or two
        Send, {ENTER 2}
    }else
    {
        Exit
    }
}


;back to excel
MoveTo("Excel")
Send, {ENTER}
Send, {Up}
Send, {Alt}HH{Right 5}{ENTER}
Send, {Down}
return
