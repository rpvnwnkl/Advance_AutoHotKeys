; #Include %A_ScriptDir%
; MsgBox %A_ScriptDir%
#Include moveTo.ahk
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
    moveTo(dbName) 
    return
}