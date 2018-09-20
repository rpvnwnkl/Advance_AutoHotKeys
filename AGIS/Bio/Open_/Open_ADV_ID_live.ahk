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
ogClipContents = %clipboard%

SetTitleMatchMode, 2
CopyCell()
{
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait ;
    ;clean up id from excel
    ;check if you are in excel so the newline can be dropped
    IfWinActive, Excel, , , 
    {
        ;clean up id from excel
        StringTrimRight, clipboard, clipboard, 2
        Send, {Esc}
    }
    cellOfInterest = %clipboard%
    clipboard = ;
    return cellOfInterest
}

;copy entity ID
entityID := CopyCell()

IfEqual, entityID, 
{
    ;MsgBox, Done
    return
}

IfWinNotExist, Production Database
{
    MsgBox, Advance isn't open
    return
} else 
{
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

    ;Sleep, 100
    Send, {ENTER}
    ;Sleep, 100
}

clipboard = %ogClipContents%
return