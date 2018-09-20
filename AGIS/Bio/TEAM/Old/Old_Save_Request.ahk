+l::
SetTitleMatchMode, 2

Send, {CTRLDOWN}w{CTRLUP}
WinWait, Team Work Request, 
IfWinNotActive, Team Work Request, , WinActivate, Team Work Request, 
WinWaitActive, Team Work Request, 
Sleep, 50
Send, {TAB}{TAB}{TAB}

Sleep, 50
Send, {DOWN}{DOWN}{ENTER}
Sleep, 25
Send, {CTRLDOWN}s{CTRLUP}
Sleep, 25
Send, {ALTDOWN}{F4}{ALTUP}
return