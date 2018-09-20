openProsWindow(prospectID, window)
{
    global dbName
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, %window%
    Send, {TAB}
    Send, {TAB}
    Send, %prospectID%
    Send, {ENTER} ;Call up Prospect window
    Sleep, 250
    ;MsgBox, I've just entered the info
    IfWinNotActive, %dbName%
    WinActivate, %dbName% 
    return
}