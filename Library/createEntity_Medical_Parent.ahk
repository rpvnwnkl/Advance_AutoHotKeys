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