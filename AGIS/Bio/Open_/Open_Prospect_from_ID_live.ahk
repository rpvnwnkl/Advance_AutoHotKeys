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

MoveTo(destination)
{
    ;Move to destination 
    ;MsgBox, Moving to %destination%
    IfWinNotActive, %destination%,
    ;MsgBox "2"
    ;#WinActivateForce
    WinActivate, %destination%, ,
    ;MsgBox "4"
    WinWaitActive, %destination%
    
}

pullupEntity(entID)
{
    ;pull up entity window
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    Send, %entID%
    Send, {Enter}
}

getToProsA()
{
    ; this gives 'control' to the current window
    ; Send, {ENTER}
    ; Send, {CtrlDown}{F4}{CtrlUp}
    Send, {F7}

    ;now we can tab to prospect window
    ;on the main bio view
    ; Send, {TAB 8}
    Send, {ShiftDown}{TAB 4}
    Send, {ShiftUp}
    Send, {Enter}

    ;now pop up a dialog and exit
    ;allowing us to navigate this window
    Send, {F3}
    Send, {ESC}

    ;now tab to assignment box
    Send, {TAB 3}
    Send, {Enter}

    ;now would be a good time to capture
    ;prospect ID, if you're interested

    ;open Assignment Window
    Send, {AltDown}A{AltUp}
    
    ;move control back to advance, jik
    MoveTo("Production Database") 
    
    
    return
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

    ;pull up entity
    pullupEntity(entityID)

    ;open PROSA window
    getToProsA()
    Sleep, 100

}

clipboard = %ogClipContents%
return