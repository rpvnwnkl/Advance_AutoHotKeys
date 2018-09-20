/*This script is designed to enter new parents for 
medical students 
*/

#A::
Reload
Return

#MaxMem 256 ; Allow 256 MB per variable

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
    WinWait, Production Database, 
    IfWinNotActive, Production Database, , WinActivate, Production Database, 
    WinWaitActive, Production Database, 
    
    
    return
}

CopyCell()
{
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait ;
    ;clean up id from excel
    StringTrimRight, clipboard, clipboard, 2 ; trim two char from end of string
    cellOfInterest = %clipboard%
    Send, {Esc}
    return cellOfInterest
}

IsDeceased(deceasedStatus)
{
    if deceasedStatus = Y
    {
        return True
    } else
        {
            return False
        }
}

createInformal(informalName)
{
    ;call up salutation window
    callVar = SAL
    callGoTo(callVar) 
    ;Now create a new entry
    Send, {F6}
    ;send informal salutation
    Send, %informalName%
    ;move to type field
    Send, {TAB}{TAB}
    ;set type as informal individual
    Send, ii
    ;save and then exit
    Send, {F8}
    Send, {CTRLDOWN}{F4}{CTRLUP} 
}


createChildRel(childID)
{
    ;call up child window
    callVar = CHILD
    callGoTo(callVar)
    ;Now create a new entry
    Send, {F6}CH{TAB}
    Send, %childID%
    ;save and then exit
    Send, {F8}
    Send, {CTRLDOWN}{F4}{CTRLUP}
    
}
createEntity(childID, Male, firstName, middleName, lastName, Suffix, addressOne, city, state, zip)
    {
        ;pull up entity window
        Send, {ShiftDown}{F4}{ShiftUp}
        Sleep, 25
        Send, {F6}

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
        If %Male%
        {
            Send, Mr.
        } else {
            Send, Mrs.
        }
        Send, {Tab}
        Send, %Suffix%

        ;now year, school + type
        Send, {Tab 3}
        Send, 2021
        Send, {Tab}
        Send, MD
        Send, {Tab}
        Send, PA

        ;now the address
        Send, {AltDown}A{AltUp}
        Send, %addressOne%
        Send, {AltDown}Z{AltUp}
        Send, %zip%{TAB}{Down}
        Send, {AltDown}m{AltUp}
        Send, Team 240749
        
        ;elaborate/brute saving process
        Send, {F8}
        Sleep, 300
        ;MsgBox, catch me
        ;enter for parent-child warning, successful window creation, and duplicate warning
        
        ;enter for duplicates
        Send, {Enter}
        Sleep, 300
        
        ;enter again for same
        Send, {Enter}
        Sleep, 300

        ;about child warning
        ;IfWinExist, Warning
        Send, {Enter}
        ;WinWaitClose, Warning
        Sleep, 200
        ;IfWinExist, Adding...
        Send, {Enter}
        ;WinWaitClose, Adding...
        Sleep, 200

        ;IfWinExist, Primary Entity
        Send, {Enter}
        ;WinWaitClose, Primary Entity
    
    
        Send, {Enter}
        Sleep, 1000
        
        Send, {Enter}
        ;MsgBox, we've just saved the entity

        Sleep, 2000
        ;adding details
        ;gathering new ID number
        ;get window text to get ID
        WinGetText, entityWindow 
        Sleep, 200
        ;StringGetPos, numPosition, entityWindow, #
        ;MsgBox, The text is:`n%entityWindow%
        ;while we could find the position of the id num
        ;we know where it is for now, so we will hard code it
        StringMid, idNum, entityWindow, 15, 6
        ;MsgBox, ID is %idNum%
        ;MsgBox, didn't find it!    
        
        ;####################
        Sleep, 100
        ;create informal salutation
        createInformal(firstName)
        Sleep, 100
        ;MsgBox, informal just finished
        ;add child to record
        createChildRel(childID)
        Sleep, 100
        ;MsgBox, child just finished
        Send, {CtrlDown}{F4}{CtrlUp}
        Sleep, 200
        Send, N
        Sleep, 100
        Send, {CtrlDown}{F4}{CtrlUp}
        return %idNum%

    }
marry(spouse1, spouse2, fatherFirst, fatherMiddle, fatherLast, fatherSuffix, motherFirst, motherMiddle, motherLast)
{
    ;pull up spouse1 record
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 50
    Send, %spouse1%
    Send, {ENTER}

    ;call up marriage window
    callVar = Mar
    callGoTo(callVar)
    Sleep, 50

    ;status is Married - M
    Send, M{TAB}
    ;spouse ID
    Send, %spouse2%{TAB}
    
    ;enter joint names
    If fatherLast <> %motherLast% 
    {
        ;we are stacking names
        
        ;find middle initial
        StringTrimRight, fatherInitial, fatherMiddle, StrLen(fatherMiddle)-1
        StringTrimRight, motherInitial, motherMiddle, StrLen(motherMiddle)-1
        
        ;enter appropriately given whether initial exists
        If StrLen(fatherInitial) > 0
        {
            Send, Mr. %fatherFirst% %fatherInitial%. %fatherLast%
        } else
        {
            Send, Mr. %fatherFirst% %fatherLast%
        }
        ;Father suffix (jr, etc)
        If fatherSuffix is not space
        {
            Send, , %fatherSuffix%
            Send, .
        }
        ;now mother
        Send, {TAB}
        If StrLen(motherInitial) > 0
        {
            Send, Ms. %motherFirst% %motherInitial%. %motherLast%
        } else
        {
            Send, Ms. %motherFirst% %motherLast%
        }
        
        ;build jnt sal now
        Send, {TAB}
        Send, Mr. %fatherLast% and Ms. %motherLast%
        Send, {Tab 13}
    } else 
        {
            ;we accept default naming scheme
            Send, {TAB 15}
        }
    Send, Team 240749
    ;save and then uncheck home phone
    Send, {F8}
    Sleep, 50
    Send, {Enter}
    Send, {Space}
    Send, {F8}

    ;add joint informal salutation in sal window
    callVar = SAL
    callGoTo(callVar)
    Send, {F6}
    Send, %fatherFirst% and %motherFirst%
    Send, {Tab 2}ji{F8}
    Send, {SHIFTDOWN}{F3}{SHIFTUP}
    Send, {F6}
    Send, %fatherFirst% and %motherFirst%
    Send, {Tab 2}ji{F8}
    Send, {CTRLDOWN}{F4}{CTRLUP}

    ;close marriage window
    Send, {CTRLDOWN}{F4}{CTRLUP}
    ;close entity window
    Send, {CTRLDOWN}{F4}{CTRLUP}
    return

}
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

;###############################
; STUDENT
studentID = ;
studentLast = ;
studentFirst = ;

;studentID
;capture studentID (col 1)
;should already be selected
studentID := CopyCell()
;MsgBox, studentID = %studentID%

;check to see if the row is empty/without primary donor id
if (studentID < 1) {
    MsgBox, No Student ID
    return
}

;studentLast
;capture studentLast (col 2)
Send, {Tab 1}
studentLast := CopyCell()

;studentFirst
;capture studentFirst (col 3)
Send, {Tab 1}
studentFirst := CopyCell()

;################################
; FATHER
fatherID = ;
fatherDeceased = ;
fatherFirst = ;
fatherMiddle = ;
fatherLast = ;
fatherSuffix = ;

;fatherID is not given, so we will skip it
;(col 4)
Send, {Tab 1}

;fatherDeceased (col 5)
Send, {Tab 1}
fatherDeceased := CopyCell()
fatherDeceased := IsDeceased(fatherDeceased)

;fatherFirst (col 6)
Send, {Tab 1}
fatherFirst := CopyCell()

;fatherMiddle (col 7)
Send, {Tab 1}
fatherMiddle := CopyCell()

;fatherLast (col 8)
Send, {Tab 1}
fatherLast := CopyCell()

;fatherSuffix (col 9)
Send, {Tab 1}
fatherSuffix := CopyCell()

;##################################
; MOTHER
motherID = ;
motherDeceased = ;
motherFirst = ;
motherMiddle = ;
motherLast = ;
motherSuffix = ;

;motherID is not given, so we will skip it
;(col 10)
Send, {Tab 1}

;motherDeceased (col 11)
Send, {Tab 1}
motherDeceased := CopyCell()
motherDeceased := IsDeceased(motherDeceased)

;motherrFirst (col 12)
Send, {Tab 1}
motherFirst := CopyCell()

;motherMiddle (col 13)
Send, {Tab 1}
motherMiddle := CopyCell()

;motherLast (col 14)
Send, {Tab 1}
motherLast := CopyCell()

;motherSuffix (col 15)
Send, {Tab 1}
motherSuffix := CopyCell()

;################################
; ADDRESS
addressOne = ;
addressTwo = ;
city = ;
state = ;
country = ;
zip = ;

;addressOne (col 16)
Send, {Tab 1}
addressOne := CopyCell()

;addressTwo (col 17)
;not being used right now, will skip
Send, {Tab 1}

;city (col 18)
Send, {Tab 1}
city := CopyCell()

;state (col 19)
Send, {Tab 1}
state := CopyCell()

;country (col 20)
Send, {Tab 1}
;skip

;zip (col 21)
Send, {Tab 1}
zip := CopyCell()


;########################################
;ENTER DATA INTO ADVANCE
;########################################

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, , 
WinWaitActive, Production Database, , 
Sleep, 200
father = False
mother = False
If fatherFirst is not space
{
    ;MsgBox, .%fatherFirst%.
    ;create father
    father = True
    fatherID := createEntity(studentID, True, fatherFirst, fatherMiddle, fatherLast, fatherSuffix, addressOne, city, state, zip)
    Sleep, 500
}
If motherFirst is not space
{
    ;create mother
    mother = True
    motherID := createEntity(studentID, False, motherFirst, motherMiddle, motherLast, motherSuffix, addressOne, city, state, zip)
    Sleep, 500
}
If motherFirst is not space
{
    If fatherFirst is not space
    {
        ;marry the two records
        marry(fatherID, motherID, fatherFirst, fatherMiddle, fatherLast, fatherSuffix, motherFirst, motherMiddle, motherLast)
        Sleep, 200
    }
}

;back to excel
WinWait, Excel, 
IfWinNotActive, Excel, , Excel, 
WinActivate, Excel,

;shift tab back to mother id field
Send, {ShiftDown}{Tab 11}{ShiftUp}
Send, %motherID%

;now father id field
Send, {ShiftDown}{Tab 6}{ShiftUp}
Send, %fatherID%

;new line start
Send, {Enter}
Sleep, 200
}
return