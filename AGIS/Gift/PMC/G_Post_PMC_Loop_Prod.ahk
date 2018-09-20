/*This script is designed to enter PMC info from an excel doc
Start with excel spreadsheet giving id num, gift ratio, and gift total
start with cursor on id num
*/

#A::
Reload
Return

#MaxMem 256 ; Allow 256 MB per variable
/*
#####################################
SCRIPT START
*/
#G::
SetTitleMatchMode, 2

Loop {
;#####################################
;COLLECT DATA FROM EXCEL
;#####################################
WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel, ,WinWaitActive, Excel 

primaryIdNum = 293573
donorIdNum = ;
spouseIdNum = ;
fullGiftAmt = ;
legalGiftAmt = ;
runnerIdNum = ;
anonGift = ;

;donorIdNum
;capture entity ID (col 1)
;should already be selected
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
;clean up id from excel
donorIdNum = %clipboard%
StringTrimRight, donorIdNum, donorIdNum, 2 ; trim two char from end of string
Send, {Esc}
Sleep, 25

;check to see if the row is empty/without primary donor id
if (donorIdNum < 1) {
    MsgBox, No Primary Donor ID
    return
}

;spouseIdNum
;clear clipboard, get spouseIdNum (col 3)
clipboard = ;
Send, {Tab 2}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
;clean up id from excel
spouseIdNum = %clipboard%
StringTrimRight, spouseIdNum, spouseIdNum, 2 ; trim two char from end of string
Send, {Esc}
Sleep, 25

;fullGiftAmt
;clear clipboard, get full donation amount (col 5)
clipboard = ;
Send, {Tab 2}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
;clean up number from excel
fullGiftAmt = %clipboard%
StringTrimRight, fullGiftAmt, fullGiftAmt, 2 ; trim two char from end of string
Send, {Esc}
Sleep, 25

;legalGiftAmt
;clear clipboard, get legal gift amount (col 6)
clipboard = ;
Send, {Tab}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
;clean up number from excel
legalGiftAmt = %clipboard%
StringTrimRight, legalGiftAmt, legalGiftAmt, 2 ; trim two char from end of string
Send, {Esc}
Sleep, 25

;runnerIdNum
;capture entity ID (col 19)
clipboard = ;
Send, {Tab 13}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
;clean up id from excel
runnerIdNum = %clipboard%
StringTrimRight, runnerIdNum, runnerIdNum, 2 ; trim two char from end of string
Send, {Esc}
Sleep, 25

;anonGift
;capture anon flag (col 21)
clipboard = ;
Send, {Tab 2}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
;clean up id from excel
anonGift = %clipboard%
StringTrimRight, anonGift, anonGift, 2 ; trim two char from end of string
Send, {Esc}
Sleep, 25

;reset for next row
Sleep, 100
clipboard = ;
Sleep, 100
Send, {ENTER}

/*
MsgBox, donorIdNum is %donorIdNum% 
    , spouseIdNum is %spouseIdNum% 
    , fullGiftAmt is %fullGiftAmt%
    , legalGiftAmt is %legalGiftAmt%
    , runnerIdNum is %runnerIdNum%
    , anonGift is %anonGift%.
*/


;########################################
;ENTER DATA INTO ADVANCE
;########################################

;move to Advance
;batch is created, defaults are set, ledger is open
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, , 
WinWaitActive, Production Database, , 
Sleep, 100

;batch and ledge are already open, with defaults set
;open new transaction
Send, {F6}

;Enter ID Num for primary donor
Send, %primaryIdNum%
Send, {Tab}
Sleep, 50
;Press 'O' for OK
Send, O
Sleep, 50

;Enter legal amount of gift
Send, %legalGiftAmt%

;check for Anon flag
if (anonGift = "Y") {
    ;move to "Asc Type" field
    Send, {AltDown}k{AltUp}
    Sleep, 25
    Send, {ShiftDown}{Tab}{ShiftUp}
    Sleep, 25
    Send, 2
}
;Save
Send, {F8}
Sleep, 100

;###############################
;ASSOCIATED DONORS WINDOW
;##################################

;Open Assoc. donors window
Send, {AltDown}{AltUp}
Send, G
Send, A
Sleep, 100

;##################################
;ASSOCIATED DONOR 1
;##################################
;add addt'l donor
Send, {F6}

;Enter donorIdNum
Send, %donorIdNum%
;move to "Asc Type" field
Send, {AltDown}A{AltUp}
;Sleep, 500
;Press 'O' for OK in ensuing pop-up
Send, O
Sleep, 500
;get window text to check for alerts
WinGetText, alertCheck 
;MsgBox, The text is:`n%alertCheck%

;If found, alert popped up on record and we close it
IfInString, alertCheck, alert
{
    ;MsgBox, Closing this window
    ;close window
    Send, {CtrlDown}{F4}{CtrlUp}
    ;move to donor info window
    WinWait, Donor Information, 
    IfWinNotActive, Donor Information, , WinActivate, Donor Information, , 
    WinWaitActive, Donor Information, , 
    Sleep, 100
    ;close window
    Send, O
    ;move back to advance
    WinWait, Production Database, 
    IfWinNotActive, Production Database, , WinActivate, Production Database, , 
    WinWaitActive, Production Database, , 
    Sleep, 100
}

;Press 'O' for OK in ensuing pop-up
;Send, O
;Sleep, 50

;ENTER  'D' for Donor advised type
Send, D
Sleep, 25

;move to credit field
Send, {AltDown}m{AltUp} 
Send, {TAB 4}
Sleep, 25

;change entry to full amount
Send, {ShiftDown}{CtrlDown}{Right}{CtrlUp}{ShiftUp}
Send, %fullGiftAmt%
Sleep, 25

;save entry
Send, {F8}
Sleep, 100

;###########################################
;Check for spouse ID and add to Assoc Donors 
;###########################################

if (spouseIdNum > 0) {
    ;MsgBox Spouse Found!
    ;add spouse as addt'l donor
    Send, {F6}

    ;Enter spouseIdNum
    Send, %spouseIdNum%
    ;move to "Asc Type" field
    Send, {AltDown}A{AltUp}
    Sleep, 50

    ;Press 'O' for OK in ensuing pop-up
    Send, O
    Sleep, 50

    ;ENTER  'D' for Donor advised type
    Send, D
    Sleep, 25

    ;move to credit field
    Send, {AltDown}m{AltUp} 
    Send, {TAB 4}
    Sleep, 25

    ;change entry to full amount
    Send, {ShiftDown}{CtrlDown}{Right}{CtrlUp}{ShiftUp}
    Send, %fullGiftAmt%
    Sleep, 25

    ;save entry
    Send, {F8}
    Sleep, 100

    ;Fix first donor credit amount now
    Send, {DOWN}
    ;move to credit field
    Send, {AltDown}m{AltUp} 
    Send, {TAB 4}
    Sleep, 25
    ;change entry to full amount
    Send, {ShiftDown}{CtrlDown}{Right}{CtrlUp}{ShiftUp}
    Send, %fullGiftAmt%
    Sleep, 25

    ;save entry again
    Send, {F8}
    Sleep, 100
}
;###################################
;IN HONOR OF DONOR/RUNNER
;###################################
if (donorIdNum = runnerIdNum) {
    ;change 'D' to 'H' in associated donor window
    Send, {Down}
    ;move to "Asc Type" field
    Send, {AltDown}A{AltUp}
    Sleep, 25
    Send, H
    ;save entry again
    Send, {F8}
    Sleep, 100

}
if (spouseIdNum = runnerIdNum) {
    ;change 'D' to 'H' in associated donor window
    Send, {Down 2}
    ;move to "Asc Type" field
    Send, {AltDown}A{AltUp}
    Sleep, 25
    Send, H
    ;save entry again
    Send, {F8}
    Sleep, 100
} else if (runnerIdNum > 0 
    and donorIdNum != runnerIdNum 
    and spouseIdNum != runnerIdNum) {
        ;add RUNNER as assoc. donor
        Send, {F6}

        ;Enter runnerIdNum
        Send, %runnerIdNum%
        ;move to "Asc Type" field
        Send, {AltDown}A{AltUp}
        Sleep, 50

        ;Press 'O' for OK in ensuing pop-up
        Send, O
        Sleep, 50

        ;ENTER  'H' for in-honor type
        Send, H
        Sleep, 25

        ;save entry
        Send, {F8}
        Sleep, 100
}
;######################################
;PRIMARY DONOR WINDOW
;###############################
;Update credit amount under primary donor 
;Open Primary donors window
Send, {AltDown}{AltUp}
Send, G
Send, P
Sleep, 100

;move to Credit field, in two steps
Send, {AltDown}i{AltUp}
Send, {TAB}
Sleep, 25
Send, %fullGiftAmt%

;save entry
Send, {F8}
Sleep, 100


;back to square one
}
;WinWait, Excel, 
;IfWinNotActive, Excel, , Excel, 
;WinActivate, Excel, 
;clipboard = ;

return