;Context: You have an ID highlighted and would like to search it in Advance
;	
;	the script copies ID and moves over to advance 
;	an entity search window is opened and the ID pasted in
;	

;Instructions: have advance open, select the ID you'd like to search 
;	execute this script by pressing the middle mouse button
;	
;v 1.0
;Pressing windows key + A reloads the script in case it gets wonky
#A::
Reload
Return

;The middle mouse button launches the script
MButton::
SetTitleMatchMode, 2
Sleep, 100

 
;Copy highlighted ID
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait, 1 ;
entityID = %clipboard%

;check if you are in excel so the newline can be dropped
IfWinActive, Excel, , , 
{
    ;clean up id from excel
    StringTrimRight, entityID, entityID, 2 ; trim two char from end of string
    Send, {Esc}
}

IfEqual, entityID, 
{
    ;MsgBox, Done
    return
}
;clear clipboard
clipboard = ;

IfWinNotExist, Production Database
{
    MsgBox, Advance isn't open
    return
}

;move to Advance
WinWait, Production Database,
IfWinNotActive, Production Database, , WinActivate, Production Database,
WinWaitActive, Production Database,
Sleep, 100

;open Entity search window
Send, {SHIFTDOWN}{F4}{SHIFTUP}
Sleep, 100

;Paste id into first text box
Send, %entityID%

Sleep, 100
Send, {ENTER}
Sleep, 100

return