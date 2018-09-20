
;Context: You adding new emails to records

;Instructions: have advance and excel open
;               email and id side by side
;               press Windows Key plus G

#A::
Reload
Return


#G::
SetTitleMatchMode, 2

WinWait, Excel, 
IfWinNotActive, Excel, , WinActivate, Excel, 
WinWaitActive, Excel,

;capture ID
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

;clean up from excel
entityID = %clipboard%
StringTrimRight, entityID, entityID, 2 ; trim two char from end of string
clipboard = ;

;move right, Capture email
Send, {RIGHT}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

;clean up from excel
entEmail = %clipboard%
StringTrimRight, entEmail, entEmail, 2 ; trim two char from end of string

;Empty clipboard
clipboard = ;

;reset cursor to next line
Send, {DOWN}{LEFT}
Sleep, 100

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 200

;pull up entity search window, enter ID
Send, {SHIFTDOWN}{F4}{SHIFTUP}
;clipboard = entityID
;ClipWait ;
;MsgBox Clipboard is:`n`n%clipboard%
Sleep, 50

;paste ID in
Send, %entityID%
Sleep, 100

;submit search
Send, {ENTER}
Sleep, 300

;pull up email window
Send, {F5}
Send, EMAIL
Send, {ENTER}
Sleep, 100

;New email
Send, {F6}
Send, %entEmail%

;set preferred
;Send, {TAB}{SPACE}
;or
;don't set preferred
Send, {TAB}
;set status date
FormatTime, timeVar, ,ShortDate
Send, {TAB 2}%timeVar%


;prompt for P or B email type
emailType = P
/*
;#SingleInstance
SetTimer, ChangeButtonNames, 50
MsgBox, 4, Personal or Business, Choose an email type:
IfMsgBox, Yes
    emailType = P
Else
    emailType = B

*/
;set email type
Send, {TAB}%emailType%

;Set donor response
Send, {TAB 3}DON

;add Art Sale to comment
Send, {TAB 6}
Send, Team 240144


;save
Send, {F8}
Sleep, 200

;test for duplicate
/*DetectHiddenText, On
WinGetText, windowCheck
Sleep, 500 
MsgBox, The text is:`n%windowCheck%
;term to check for
windowText = already
;Search for windowText in current windows
IfInString, windowCheck, %windowText%
{
    Send, N
    Send, {F7}
}
*/
;in case of duplicate, send N and refresh
Send, N
Send, {F7}
;close
Send, {CtrlDown}{F4}{F4}{F4}{CtrlUp}

WinWait, Excel, 
IfWinNotActive, Excel, , WinActivate, Excel, 
WinWaitActive, Excel,

return

ChangeButtonNames:
IfWinNotExist, Personal or Business
    return ; keep waiting
SetTimer, ChangeButtonNames, off
WinActivate
ControlSetText, Button1, &Personal
ControlSetText, Button2, &Business
return