;Context: You have an ID highlighted and would like to search it in Advance
;	
;	the script copies ID and moves over to advance 
;	an entity search window is opened and the ID pasted in
;  Afterwards, the DEG and REC windows are opened
;	

;Instructions: have advance open, select the ID you'd like to search 
;	execute this script by pressing the middle mouse button
;	
;v 1.0
;Pressing windows key + A reloads the script in case it gets wonky
#A::
Reload
Return

callGoToBlank(tableText, entID)
{
    SetTitleMatchMode, 2
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, %tableText%
    Send, {TAB}
    Send, %entID%{ENTER} ;Call up given window
    Sleep, 250
    ;MsgBox, I've just entered the info
    WinWait, Production Database, 
    IfWinNotActive, Production Database, , WinActivate, Production Database, 
    WinWaitActive, Production Database, 
    
    
    return
}

pullupEntity(entID)
{
    ;pull up entity window
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    Send, %entID%
    Send, {Enter}
}

;The middle mouse button launches the script
^MButton::
SetTitleMatchMode, 2
Sleep, 100

 
;Copy highlighted ID
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
entityID = %clipboard%

;check if you are in excel so the newline can be dropped
IfWinActive, Excel, , , 
{
    ;clean up id from excel
    StringTrimRight, entityID, entityID, 2 ; trim two char from end of string
    Send, {Esc}
}

;clear clipboard
clipboard = ;

;select row text
;Send, {CtrlDown}{Right}{CtrlUp}
;Send, {ShiftDown}{CtrlDown}{Left}{CtrlUp}{ShiftUp}

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 100

;close previous windows
Send, {CtrlDown}{F4 3}{CtrlUp}
;open Entity search window
;Send, {SHIFTDOWN}{F4}{SHIFTUP}
;Sleep, 100
pullupEntity(entityID)

callVar = DEG
callGoToBlank(callVar, entityID)

callVar = REC
callGoToBlank(callVar, entityID)

;Paste id into first text box
;Send, %entityID%

;Sleep, 100
;Send, {ENTER}
;Sleep, 100

return