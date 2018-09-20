callGoToBlank(tableText, entID)
{
    global dbName
    SetTitleMatchMode, 2
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, %tableText%
    Send, {TAB}
    ; Send, {AltDown}Y
    ; Send, {AltUp}
    Send, %entID%
    Send, {ENTER} ;Call up given window
    Sleep, 250
    ;move control back to advance, jik
    MoveTo(dbName) 
    
    
    return
}