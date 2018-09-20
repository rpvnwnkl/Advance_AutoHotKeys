
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetTitleMatchMode, 2

WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database,

callVar = 47
callGoTo(callVar)

Send, SFDC
Send, {TAB}

commentText = Salesforce Communities
FormatTime, commentDate, , M/d/yyy
Send, %commentText% %commentDate%
Send, {TAB}
Send, SF
Send, {TAB 4}
Send, {ENTER}

;close the dataloader window
closeWindows(1)

;open adv loader
callVar = 46
callGoTo(callVar)
Run, SalesForce_Tools.ahk,,,
return

closeWindows(numWindows)
{
    ;close windows
    Send, {CtrlDown}{F4 %numWindows%}{CtrlUp}
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
    WinWait, Production Database, 
    IfWinNotActive, Production Database, , WinActivate, Production Database, 
    WinWaitActive, Production Database, 
    
    
    return
}

