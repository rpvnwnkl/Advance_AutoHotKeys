/*
This script copies advance IDs from the Team Request window and pastes it into and excel
sheet. It then moves back to Team and presses page down to move to the next ticket

THe id window must be selected
excel sheet must be open
*/

#U::
SetTitleMatchMode, 2
;Activate Team request window
WinWait, Team Work Request, 
IfWinNotActive, Team Work Request, , WinActivate, Team Work Request, 
WinWaitActive, Team Work Request, 
Sleep, 25

;Copy adv id
Send, {CTRLDOWN}a{CTRLUP}{CTRLDOWN}c{CTRLUP}
ClipWait ;

;open excel sheet
WinWait, Excel, 
IfWinNotActive, Excel, , WinActivate, Excel, 
WinWaitActive, Excel, 
Sleep, 25

;paste adv id
Send, {CTRLDOWN}v{CTRLUP}{ENTER}

;go back to Team
WinWait, Team Work Request, 
IfWinNotActive, Team Work Request, , WinActivate, Team Work Request, 
WinWaitActive, Team Work Request, 

;move to next ticket
Send, {PGDN}

return