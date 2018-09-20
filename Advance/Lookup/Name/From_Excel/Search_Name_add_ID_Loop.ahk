
;Context: You are using first and last names in advance using a spreadsheet as source

;Instructions: have advance and excel open
;               select the cell of the first last name you'd like to search
;               press Windows Key plus G

#A::
Reload
Return


#G::
;MsgBox F started
SetTitleMatchMode, 2
SetTitleMatchMode, slow


/*
; Example: Achieve an effect similar to SplashTextOn:

Gui, +AlwaysOnTop +Disabled -SysMenu +Owner  ; +Owner avoids a taskbar button.
Gui, Add, Text,, Searching for IDs...
Gui, Show, NoActivate, Running  ; NoActivate avoids deactivating the currently active window.
*/

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
Send, {Esc}

;check for end of column
if (lastName = "") {
    ;MsgBox, Cooked
    Break
}
;move lft, Capture first name
clipboard = ;

Send, {LEFT}
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;

;clean up from excel
firstName = %clipboard%
StringTrimRight, firstName, firstName, 2 ; trim two char from end of string
Send, {Esc}

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
Send, {ENTER}
Send, {ENTER}
Send, {ENTER}
Sleep, 300

;#####################################
;case for multiple results from lookup
;#####################################
;get window text to check for multiple hits
;DetectHiddenText, On
WinGetText, windowCheck 
;MsgBox, The text is:`n%windowCheck%

;term to check for
windowText = Criteria
;Search for windowText in current windows
IfInString, windowCheck, %windowText%
{
    MsgBox, 262145,, Please open the correct entity record
                
    IfMsgBox Cancel
        break
    ;check if profile was opened or not
    WinGetText, windowCheck
    ;Search for windowText in current windows
    windowText = Custom
    IfNotInString, windowCheck, %windowText%
    {
       Send, {Enter} 
    }
}
;pull up id window
Send, {F3}{F3}
Sleep, 250
;MsgBox, Hello
;Sleep, 100


;copy id to clipboard
Send, {CTRLDOWN}c{CTRLUP}
ClipWait ;
advID = %clipboard%
;MsgBox, clipboard is %advID%
;close the Id window
Send, {Esc}

;close the entity window
Send, {CTRLDOWN}{F4}{CTRLUP}

;close the entity lookup window
Send, {CTRLDOWN}{F4}{CTRLUP}

;Move back to Excel
WinWait, Excel, 
IfWinNotActive, Excel, , Excel,
WinActivate, Excel, ,WinWaitActive, Excel,

;arrow to next column to paste ID
Send, {LEFT}

;paste ID from clipboard
Send, %advID%
Send, {Enter}{Tab}{Tab}
} 
;Gui, Cancel
return