
;v 1.0
SetTitleMatchMode, 2

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

CopyAndCheck()
{
    ;Copy highlighted ID
    global
    entityID := CopyCell()

    IfEqual, entityID, 
    {
        ;MsgBox, Done
        return
    }

    IfWinNotExist, Production Database
    {
        MsgBox, Advance isn't open
        return
    }

    ;move to Advance
    MoveTo("Production Database")
    Sleep, 100 
    return entityID   
}

#IfWinActive, Production Database
^+U::
;open windows
callGoTo("STAFF")
callGoTo("EMAIL")
callGoTo("TEL")
callGoTo("ADDR")
callGoTo("EMPT")
callGoTo("REC")

return
