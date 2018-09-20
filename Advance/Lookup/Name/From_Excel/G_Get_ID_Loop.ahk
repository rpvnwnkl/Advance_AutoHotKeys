
;Context: You are using first and last names in advance using a spreadsheet as source

;mostly assumes the record exists in advance

;Instructions: have advance and excel open
;               select the cell of the first last name you'd like to search
;               press Windows Key plus G

#R::
Reload
Return


#G::
;MsgBox F started
SetTitleMatchMode, 2

Loop {
WinWait, Excel, 
IfWinNotActive, Excel, , WinActivate, Excel, 
WinWaitActive, Excel,

;capture last name
clipboard = ;
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

;clean up from excel
lastName = %clipboard%
StringTrimRight, lastName, lastName, 2 ; trim two char from end of string

;move lft, Capture first name
clipboard = ;

Send, {LEFT}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

;clean up from excel
firstName = %clipboard%
StringTrimRight, firstName, firstName, 2 ; trim two char from end of string

;Empty clipboard
clipboard = ;

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 100

;pull up entity search window, tab to last name
Send, {SHIFTDOWN}{F4}{SHIFTUP}{TAB}{TAB}{TAB}
clipboard = lastName
ClipWait ;
;MsgBox Clipboard is:`n`n%clipboard%
Sleep, 50

;paste last name in
Send, %lastName%
Sleep, 50

;empty clipboard, add first name to it
clipboard = ;
Sleep, 50
clipboard = %firstName%
ClipWait ;
;MsgBox Clipboard is:`n`n%clipboard%
Sleep, 50

;past first name in
Send, {TAB}
Send, %firstName%
Sleep, 50

;submit search
Send, {ENTER}

;send Y, just in case last name is too short
Send, {Y}
Sleep, 50
clipboard = ;
Sleep, 100
;Send, {ENTER}
;Send, {ENTER}
;Send, {ENTER}
Sleep, 300

;pull up id window
Send, {F3}{F3}
Sleep, 100
MsgBox, Hello
Sleep, 100
;copy id to clipboard
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
advID = %clipboard%
MsgBox, clipboard is %advID%
;close the Id window
Send, {Esc}

;close the entity window
Send, {CTRLDOWN}{F4}{CTRLUP}

;Move back to Excel
WinWait, Excel, 
IfWinNotActive, Excel, , Excel,
WinActivate, Excel, ,WinWaitActive, Excel,

;arrow to next column to paste ID
Send, {LEFT}

;paste ID from clipboard, prepare for next row
Send, %advID%
Send, {Enter}{Tab}{Tab}
} 
return