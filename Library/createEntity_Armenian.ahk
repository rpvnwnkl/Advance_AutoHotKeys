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