
;Pressing windows key + A reloads the script in case it gets wonky
#A::
Reload
Return

#T::
SetTitleMatchMode, 2

;make sure we are in Team
WinWait, Alexsys Team 2 Pro, 
IfWinNotActive, Alexsys Team 2 Pro, , WinActivate, Alexsys Team 2 Pro, 
WinWaitActive, Alexsys Team 2 Pro,

;Ctrl+W opens team request from list view
Send, {CTRLDOWN}w{CTRLUP}

;make sure ticket window is open
WinWait, Team Work Request, 
IfWinNotActive, Team Work Request, , WinActivate, Team Work Request, 
WinWaitActive, Team Work Request, 
Sleep, 200

;reset to first field
;F9 sends 'edit' command
Send, {F9}
Send, {TAB}{TAB}{TAB}
;MouseClick, left,  221,  162
;Sleep, 25
Send, w
;Send, {DOWN}{DOWN}{ENTER}
;Sleep, 100
;save changes with Ctrl+S
Send, {CTRLDOWN}s{CTRLUP}
Sleep, 100
;Close with Alt+F4
Send, {ALTDOWN}{F4}{ALTUP}
return