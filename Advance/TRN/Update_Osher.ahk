
#o:: 
SetTitleMatchMode, 2
WinWait, Primary Gift, 
IfWinNotActive, Primary Gift, , WinActivate, Primary Gift, 
WinWaitActive, Primary Gift,
Send, {ENTER}{ENTER}
WinWait, ADVTRN, 
IfWinNotActive, ADVTRN, , WinActivate, ADVTRN, 
WinWaitActive, ADVTRN, 
Send, {TAB}{TAB}{TAB}{TAB}{TAB}{TAB}
Sleep, 100
Send, osh
Sleep, 100
Send, {F8}
Sleep, 100
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 100
Send, {CTRLDOWN}{F4}{CTRLUP}
Sleep, 100
Send, {ALTDOWN}{ALTUP}
Send, a
Send, g
Send, g