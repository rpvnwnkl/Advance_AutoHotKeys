/*This script is designed to enter new parents for 
medical students 
*/



#MaxMem 256 ; Allow 256 MB per variable
dbName = Production Database
comment = Team 248279
source = BIO

MoveTo(destination)
{
    ;Move to destination 
    ;MsgBox, Moving to %destination%
    IfWinNotActive, %destination%,
    ;MsgBox "2"
    ;#WinActivateForce
    WinActivate, %destination%, ,
    ;MsgBox "4"
    WinWaitActive, %destination%
    
}
callGoTo(tableText)
{
    global dbName
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
    MoveTo(dbName) 
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

createEntity(gender, firstName, lastName, recSchool, recYear, recType)
    {
        global comment
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
        ; Send, %middleName%
        Send, {Tab}
        IfEqual, gender, M
        {
            Send, Mr.
        } else {
            Send, Ms.
        }
        ; MsgBox, stopit
        ;now year, school + type
        Send, {Tab 4}
        ; MsgBox, h%recYear%h
        Send, %recYear%
        Send, {Tab}
        ; MsgBox, h%recschool%h
        Send, %recSchool%
        Send, {Tab}
        ; MsgBox, h%recType%h
        Send, %recType%
        ; MsgBox, stopit
        ;add comment
        Send, {AltDown}m{AltUp}
        Send, %comment%
        
        ;elaborate/brute-force saving process
        Send, {F8}
        ; Sleep, 300
        ; ;MsgBox, catch me
        ; ;enter for parent-child warning, successful window creation, and duplicate warning
        
        ; ;enter for duplicates
        Send, {Enter}
        Sleep, 300
        
        ; ;enter again for address
        Send, {Enter}
        Sleep, 300

        ; ;about child warning
        ; ;IfWinExist, Warning
        ; Send, {Enter}
        ; ;WinWaitClose, Warning
        ; Sleep, 200
        ; ;IfWinExist, Adding...
        ; Send, {Enter}
        ; ;WinWaitClose, Adding...
        ; Sleep, 200

        ; ;IfWinExist, Primary Entity
        ; Send, {Enter}
        ; ;WinWaitClose, Primary Entity
    
    
        Send, {Enter}
        Sleep, 300
        
        Send, {Enter}
        ;MsgBox, we've just saved the entity

        Sleep, 1000
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
        return %idNum%
        ;####################
        ; Sleep, 100
        ; ;create informal salutation
        ; createInformal(firstName)
        ; Sleep, 100
        ; ;MsgBox, informal just finished
        ; Send, {CtrlDown}{F4}{CtrlUp}
        ; Sleep, 200
        ; Send, N
        ; Sleep, 100
        ; Send, {CtrlDown}{F4}{CtrlUp}
        ; return %idNum%

    }

addEmail(emailAddress, emailType, prefAddr, emailSrc, emailComment)
{
    callGoTo("EMAIL")
    Send, {F6}
    Send, %emailAddress%
    Send, {TAB}
    If prefAddr
    {
        Send, {SPACE}
    }
    Send, {TAB 3}
    Send, %emailType%
    Send, {TAB 3}
    Send, %emailSrc%
    Send, {TAB 6}
    Send, %emailComment%
    Send, {F8}
    Send, {CtrlDown}{F4}{CtrlUp}
    return
}

addBirthDate(entityDOB, entityBirthCity, entityBirthCountry)
{
    callGoTo("BDET")
    StringSplit, birthDate, entityDOB, /,
    Send, %birthDate1%
    Send, {TAB}
    Send, %birthDate2%
    Send, {Tab}
    SEnd, %birthDate3%
    Send, {TAB}
    Send, %entityBirthCity%, %entityBirthCountry%
    Send, {F8}
    Send, {CtrlDown}{F4}{CtrlUp}
    return
}

addID(idNum, idType)
{
    callGoTo("ID")
    Send, {F6}
    Send, %idType%
    Send, {TAB}
    Send, %idNum%
    Send, {F8}
    Send, {CtrlDown}{F4}{CtrlUp}
    return
}
/*
#####################################
SCRIPT START
*/
; #IfWinActive, Excel
^+G::
SetTitleMatchMode, 2

;#####################################
;COLLECT DATA FROM EXCEL
;#####################################
MoveTo("Excel")
;######################
;NEED TO TRIM EMPTY CHARS FROM ENDS OF NAME
;(PERHAPS ALL FIELDS)
;###############################
;Record defaults
recSchool = FL
recType = CA
recYear = 2017


; ENTITY
entityID = ;
entityLast = ;
entityFirst = ;
entitySIS = ;
entityDOB = ;
entityGender = ;
entityPrefEmail = ;
entityTuftsEmail = ;
entityBirthCountry = ;
entityBirthCity = ;

;studentLast
;capture studentLast (col 2)
entityLast := CopyCell()

;check to see if the row is empty/without primary donor id
IfEqual, entityLast, 
{
    MsgBox, No Entity Name
    return
}

;entityFirst
;capture entityFirst (col 3)
Send, {Tab 1}
entityFirst := CopyCell()

; entitySIS
Send, {TAB 1}
entitySIS := CopyCell()

;entityDOB
Send, {TAB 1}
entityDOB := CopyCell()

;entityGender
Send, {TAB 1}
entityGender := CopyCell()

;entityEmail
Send, {TAB 1}
entityPrefEmail := CopyCell()

;entityTuftsEmail
Send, {TAB 1}
entityTuftsEmail := CopyCell()

;entityBirthCountry
Send, {TAB 1}
entityBirthCountry := CopyCell()

;entityBirthCity
Send, {TAB 1}
entityBirthCity := CopyCell()



;########################################
;ENTER DATA INTO ADVANCE
;########################################

;move to Advance
; MsgBox, %dbName%
MoveTo(dbName)

entityID := createEntity(entityGender, entityFirst, entityLast, recSchool, recYear, recType)
createInformal(entityFirst)
addEmail(entityPrefEmail, "P", True, source, comment)
addEmail(entityTuftsEmail, "P", False, source, comment)
addBirthDate(entityDOB, entityBirthCity, entityBirthCountry)
addID(entitySIS, "SIS")
; return
Sleep, 100
Send, {CtrlDown}{F4}{CtrlUp}
Sleep, 200
Send, N
Sleep, 100
Send, {CtrlDown}{F4}{CtrlUp}

Sleep, 500


;back to excel
MoveTo("Excel")

;shift tab back to id field
Send, {ShiftDown}{Tab 9}{ShiftUp}
Send, %entityID%


Send, {DOWN}{RIGHT}


;new line start
Sleep, 200

return